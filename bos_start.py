"""Start von nextcloud-mcp-server mit den Office-Werkzeugen (bos_office.py) als zusaetzlicher App "office".

AVAILABLE_APPS ist das Verzeichnis, aus dem app.py beim Start die Werkzeuge aller Apps registriert; ohne
--enable-app sind das alle Eintraege, also auch "office".

Der Custom Processor (dokumente, MarkItDown) liest .docx mit Fett, Kursiv und nummerierten, eingerueckten Listen.
Der eingebaute Word-Reader (Prioritaet 15) macht daraus flache Aufzaehlungen; der Custom Processor kommt deshalb
mit Prioritaet 16 davor. app.py ruft initialize_document_processors beim Start ueber den Modulnamen auf.

Login Flow v2 speichert ein App-Passwort nur, wenn das Nextcloud-Konto, das den Zugriff gewaehrt, der Aufrufer ist
(GHSA-84qv-22q6-x82r). Der Server vergleicht die Nextcloud-UID mit dem Keycloak-sub und mit der UID, die Nextcloud
fuer das Bearer-Token des Aufrufers nennt. Beides passt hier nicht: Die UID ist der userPrincipalName, sub eine UUID,
und Nextcloud prueft keine Keycloak-Tokens (user_oidc ohne Provider). Zusaetzlich gilt deshalb: Der Login-Name,
mit dem der Zugriff gewaehrt wurde, ist der preferred_username des Aufrufers aus dem Keycloak-Userinfo, und das
App-Passwort meldet sich an Nextcloud an.
"""
import logging
import re
import sys

import httpx

from nextcloud_mcp_server import app, server
from nextcloud_mcp_server.auth import grant_ownership, provision_routes
from nextcloud_mcp_server.auth.scope_authorization import invalidate_scope_cache
from nextcloud_mcp_server.auth.storage import OCS_REQUEST_HEADERS, nextcloud_httpx_client
from nextcloud_mcp_server.client.webdav import WebDAVClient
from nextcloud_mcp_server.config import get_document_processor_config, get_settings
from nextcloud_mcp_server.server import auth_tools

import bos_office

logger = logging.getLogger(__name__)

server.AVAILABLE_APPS["office"] = bos_office.configure_office_tools

_initialize_document_processors = app.initialize_document_processors


def initialize_document_processors():
    _initialize_document_processors()
    custom = get_document_processor_config()["processors"].get("custom")
    if not custom:
        return
    from nextcloud_mcp_server.document_processors import get_registry

    registry = get_registry()
    processor = registry.get_processor(custom["name"])
    if processor is not None:
        registry.register(processor, priority=16)


app.initialize_document_processors = initialize_document_processors

# WebDAV-Suche: upstream schreibt die Sortiereigenschaft als <d:order><oc:size/><d:descending/></d:order>.
# RFC 5323 und searchdav (Order::xmlDeserialize liest {DAV:}prop) verlangen <d:order><d:prop>...</d:prop>, sonst
# antwortet Nextcloud mit 400 "requested order by property is not a valid property for this scope".
_ORDER_WITHOUT_PROP = re.compile(r"<d:order>(<[a-z]+:[A-Za-z_-]+/>)(<d:(?:ascending|descending)/>)</d:order>")
_build_search_xml = WebDAVClient._build_search_xml


def _build_search_xml_with_order_prop(self, *args, **kwargs):
    return _ORDER_WITHOUT_PROP.sub(r"<d:order><d:prop>\1</d:prop>\2</d:order>", _build_search_xml(self, *args, **kwargs))


WebDAVClient._build_search_xml = _build_search_xml_with_order_prop

_caller_identities = grant_ownership.caller_identities
_grant_belongs_to_caller = grant_ownership.grant_belongs_to_caller


async def _preferred_username(access_token: str) -> str | None:
    try:
        async with httpx.AsyncClient(timeout=10) as client:
            discovery = (await client.get(get_settings().oidc_discovery_url)).raise_for_status().json()
            response = await client.get(discovery["userinfo_endpoint"],
                                        headers={"Authorization": f"Bearer {access_token}"})
            name = response.raise_for_status().json().get("preferred_username")
    except Exception as e:
        logger.warning("Keycloak userinfo for a Login Flow grant failed: %s", e)
        return None
    return name if isinstance(name, str) and name else None


async def caller_identities(user_id: str, access_token: str | None = None) -> set[str]:
    identities = await _caller_identities(user_id, access_token)
    name = await _preferred_username(access_token) if access_token else None
    if name:
        identities.add(name.casefold())
    return identities


async def grant_belongs_to_caller(identities: set[str], login_name: str | None, app_password: str) -> bool:
    if (login_name and login_name.casefold() in identities
            and await grant_ownership._ocs_whoami(auth=(login_name, app_password))):
        return True
    return await _grant_belongs_to_caller(identities, login_name, app_password)


# provision_routes und auth_tools haben beide Namen beim Import uebernommen
for module in (grant_ownership, provision_routes, auth_tools):
    module.caller_identities = caller_identities
    module.grant_belongs_to_caller = grant_belongs_to_caller

# In Nextcloud zurueckgezogenes App-Passwort: upstream prueft in den nc_auth-Werkzeugen nur, ob eines gespeichert
# ist. Der Status bleibt dann "provisioned", alle Werkzeuge enden mit 401, und nc_auth_provision_access liefert keinen
# Anmeldelink. Hier prueft Nextcloud das Passwort, bei 401/403 gilt es als nicht vorhanden und wird geloescht. 429 und
# Netzfehler sagen nichts ueber das Passwort. Upstream-Fassung: Branch pr/verify-stored-app-password in thekk1/nextcloud-mcp-server
_get_shared_storage = auth_tools.get_shared_storage


class _CheckedStorage:
    def __init__(self, storage):
        self._storage = storage

    def __getattr__(self, name):
        return getattr(self._storage, name)

    async def get_app_password_with_scopes(self, user_id: str):
        existing = await self._storage.get_app_password_with_scopes(user_id)
        if not existing:
            return existing
        try:
            async with nextcloud_httpx_client(
                base_url=get_settings().nextcloud_host,
                auth=httpx.BasicAuth(existing.get("username") or user_id, existing["app_password"]),
                timeout=10.0,
            ) as client:
                status = (await client.get("/ocs/v2.php/cloud/user", headers=dict(OCS_REQUEST_HEADERS))).status_code
        except Exception as e:
            logger.warning("Could not validate the app password of %s: %s", user_id, e)
            return existing
        if status not in (401, 403):
            return existing
        logger.info("App password of %s rejected by Nextcloud (HTTP %s), removing", user_id, status)
        await self._storage.delete_app_password(user_id)
        invalidate_scope_cache(user_id)
        return None


async def _get_checked_storage():
    return _CheckedStorage(await _get_shared_storage())


auth_tools.get_shared_storage = _get_checked_storage

from nextcloud_mcp_server.cli import cli  # noqa: E402  (nach dem Eintrag, die CLI liest AVAILABLE_APPS beim Import)

sys.exit(cli())
