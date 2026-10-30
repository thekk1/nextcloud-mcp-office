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
import sys

import httpx

from nextcloud_mcp_server import app, server
from nextcloud_mcp_server.auth import grant_ownership, provision_routes
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

from nextcloud_mcp_server.cli import cli  # noqa: E402  (nach dem Eintrag, die CLI liest AVAILABLE_APPS beim Import)

sys.exit(cli())
