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
import functools
import logging
import re
import sys

import httpx

from nextcloud_mcp_server import app, server
from nextcloud_mcp_server.auth import grant_ownership, provision_routes
from nextcloud_mcp_server.auth.scope_authorization import invalidate_scope_cache
from nextcloud_mcp_server.auth.storage import OCS_REQUEST_HEADERS, nextcloud_httpx_client
from nextcloud_mcp_server.client.ocs import OCS_STATUS_UNAUTHENTICATED, parse_ocs_envelope
from nextcloud_mcp_server.client.webdav import WebDAVClient
from nextcloud_mcp_server.config import get_document_processor_config, get_settings
from nextcloud_mcp_server.models.auth import ProvisionStatusResponse
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

# In Nextcloud zurueckgezogenes App-Passwort: upstream prueft in den nc_auth-Werkzeugen nur, ob eines gespeichert ist.
# Der Status bleibt dann "provisioned", alle Werkzeuge enden mit 401, und nc_auth_provision_access liefert keinen
# Anmeldelink. Hier fragt der Status bei Nextcloud nach und meldet "revoked", loescht aber nichts (nur lesendes
# Werkzeug). nc_auth_provision_access loescht das abgelehnte Passwort und startet einen neuen Login Flow.
# Als Ablehnung zaehlt nur HTTP 401 mit OCS-Status 997: Ein 401/403 eines Proxys, 429 und Netzfehler sagen nichts ueber
# das Passwort. Zeilen ohne Login-Name bleiben ungeprueft, die UID kann davon abweichen (upstream #980).
# Upstream-Fassung: PR #1761, Branch pr/verify-stored-app-password in thekk1/nextcloud-mcp-server
_APP_PASSWORD_CHECK_TIMEOUT = 5.0


async def _rejected_app_password(user_id: str) -> bytes | None:
    """Das verschluesselte gespeicherte Passwort, wenn Nextcloud es ablehnt, sonst None."""
    storage = await auth_tools.get_shared_storage()
    async with storage._db() as db:
        async with db.execute(
            "SELECT encrypted_password, username FROM app_passwords WHERE user_id = ?", (user_id,)
        ) as cursor:
            row = await cursor.fetchone()
    if not row or not row[1]:
        return None
    encrypted_password, username = row
    try:
        async with nextcloud_httpx_client(
            base_url=get_settings().nextcloud_host,
            auth=httpx.BasicAuth(username, storage.cipher.decrypt(encrypted_password).decode()),
            timeout=_APP_PASSWORD_CHECK_TIMEOUT,
        ) as client:
            response = await client.get("/ocs/v2.php/cloud/user", headers=dict(OCS_REQUEST_HEADERS))
        rejected = (response.status_code == 401
                    and parse_ocs_envelope(response.json()).status_code == OCS_STATUS_UNAUTHENTICATED)
    except Exception as e:
        logger.warning("Could not validate the app password of %s: %s", user_id, e)
        return None
    if not response.is_success and not rejected:
        logger.warning("App password check of %s inconclusive (HTTP %s)", user_id, response.status_code)
    return encrypted_password if rejected else None


async def _delete_app_password(user_id: str, encrypted_password: bytes) -> None:
    """Loescht genau das gepruefte Passwort: Ein waehrend der Pruefung gespeichertes neues bleibt stehen."""
    storage = await auth_tools.get_shared_storage()
    async with storage._db() as db:
        cursor = await db.execute(
            "DELETE FROM app_passwords WHERE user_id = ? AND encrypted_password = ?", (user_id, encrypted_password)
        )
        await db.commit()
        deleted = cursor.rowcount > 0
    if deleted:
        logger.info("App password of %s rejected by Nextcloud, removed", user_id)
        invalidate_scope_cache(user_id)


def _status_reports_revoked(tool):
    @functools.wraps(tool)
    async def wrapper(*args, **kwargs):
        result = await tool(*args, **kwargs)
        if result.status == "provisioned" and await _rejected_app_password(result.user_id):
            return ProvisionStatusResponse(
                status="revoked",
                message=("Nextcloud no longer accepts the stored app password (revoked in Nextcloud). "
                         "Call nc_auth_provision_access to log in again."),
                user_id=result.user_id,
                username=result.username,
            )
        return result
    return wrapper


def _provision_replaces_rejected(tool):
    @functools.wraps(tool)
    async def wrapper(*args, **kwargs):
        result = await tool(*args, **kwargs)
        if result.status == "already_provisioned":
            encrypted_password = await _rejected_app_password(result.user_id)
            if encrypted_password:
                await _delete_app_password(result.user_id, encrypted_password)
                result = await tool(*args, **kwargs)
        return result
    return wrapper


_AUTH_TOOL_CHECKS = {"nc_auth_check_status": _status_reports_revoked,
                     "nc_auth_provision_access": _provision_replaces_rejected}


class _CheckedAuthTools:
    """Steht bei register_auth_tools fuer den MCP-Server und legt die Pruefung um die fertig dekorierten Werkzeuge."""

    def __init__(self, mcp):
        self._mcp = mcp

    def tool(self, *args, **kwargs):
        register = self._mcp.tool(*args, **kwargs)
        check = _AUTH_TOOL_CHECKS.get(kwargs.get("name"))
        return (lambda tool: register(check(tool))) if check else register


_register_auth_tools = auth_tools.register_auth_tools


def register_auth_tools(mcp):
    _register_auth_tools(_CheckedAuthTools(mcp))


app.register_auth_tools = register_auth_tools

from nextcloud_mcp_server.cli import cli  # noqa: E402  (nach dem Eintrag, die CLI liest AVAILABLE_APPS beim Import)

sys.exit(cli())
