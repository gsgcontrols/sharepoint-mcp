"""SharePoint MCP server for Claude custom connectors.

Auth model: Claude signs the user in with Microsoft Entra ID (OAuth). This server
validates that user's token, then uses the On-Behalf-Of (OBO) flow to call
Microsoft Graph *as that user*. Claude can only reach what the user already can.
"""
import asyncio
import os
from typing import Any
from urllib.parse import urlencode, urlparse

import httpx
import jwt
import msal
from mcp.server.auth.middleware.auth_context import get_access_token
from mcp.server.auth.provider import AccessToken, TokenVerifier
from mcp.server.auth.settings import AuthSettings
from mcp.server.fastmcp import FastMCP
from pydantic import AnyHttpUrl
from starlette.requests import Request
from starlette.responses import JSONResponse, RedirectResponse, Response

TENANT_ID = os.environ["TENANT_ID"]
CLIENT_ID = os.environ["CLIENT_ID"]
PUBLIC_URL = os.environ.get("PUBLIC_URL", "http://localhost:8000").rstrip("/")

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
ISSUER = f"{AUTHORITY}/v2.0"
API_SCOPE = f"api://{CLIENT_ID}/access_as_user"
GRAPH = "https://graph.microsoft.com/v1.0"

_jwks = jwt.PyJWKClient(f"{AUTHORITY}/discovery/v2.0/keys")
_msal_app: msal.ConfidentialClientApplication | None = None


def msal_app() -> msal.ConfidentialClientApplication:
    """Created lazily so the container can start before the secret is configured."""
    global _msal_app
    if _msal_app is None:
        secret = os.environ.get("CLIENT_SECRET")
        if not secret:
            raise RuntimeError("CLIENT_SECRET is not configured on the server.")
        _msal_app = msal.ConfidentialClientApplication(
            CLIENT_ID, authority=AUTHORITY, client_credential=secret
        )
    return _msal_app


class EntraTokenVerifier(TokenVerifier):
    """Validates Entra ID v2 access tokens issued for this API."""

    async def verify_token(self, token: str) -> AccessToken | None:
        try:
            key = await asyncio.to_thread(lambda: _jwks.get_signing_key_from_jwt(token).key)
            claims = jwt.decode(
                token,
                key,
                algorithms=["RS256"],
                audience=[CLIENT_ID, f"api://{CLIENT_ID}"],
                issuer=ISSUER,
            )
        except Exception:
            return None
        # Entra puts short scope names in "scp"; expand to the full URI the metadata advertises.
        scopes = [f"api://{CLIENT_ID}/{s}" for s in claims.get("scp", "").split()]
        return AccessToken(
            token=token,
            client_id=claims.get("azp", ""),
            scopes=scopes,
            expires_at=claims.get("exp"),
        )


mcp = FastMCP(
    "SharePoint",
    host="0.0.0.0",
    port=int(os.environ.get("PORT", "8000")),
    stateless_http=True,
    token_verifier=EntraTokenVerifier(),
    auth=AuthSettings(
        # This server fronts Entra's sign-in (see "OAuth proxy" below), so it is
        # the authorization server Claude talks to.
        issuer_url=AnyHttpUrl(PUBLIC_URL),
        resource_server_url=AnyHttpUrl(f"{PUBLIC_URL}/mcp"),
        required_scopes=[API_SCOPE],
    ),
)


async def graph(method: str, path: str, **kwargs: Any) -> dict:
    """Call Microsoft Graph on behalf of the signed-in user."""
    user = get_access_token()
    if user is None:
        raise RuntimeError("No signed-in user.")
    result = await asyncio.to_thread(
        msal_app().acquire_token_on_behalf_of,
        user.token,
        ["https://graph.microsoft.com/.default"],
    )
    if "access_token" not in result:
        raise RuntimeError(f"Could not get Graph token: {result.get('error_description', result)}")
    async with httpx.AsyncClient(timeout=30) as client:
        r = await client.request(
            method,
            GRAPH + path,
            headers={"Authorization": f"Bearer {result['access_token']}"},
            **kwargs,
        )
    if r.status_code >= 400:
        raise RuntimeError(f"Graph error {r.status_code}: {r.text[:800]}")
    return r.json() if r.content else {}


def _site_summary(s: dict) -> dict:
    return {k: s.get(k) for k in ("id", "displayName", "name", "webUrl")}


# ---------------------------------------------------------------- tools

@mcp.tool()
async def search_sites(query: str) -> list[dict]:
    """Search SharePoint sites the user can access by keyword."""
    data = await graph("GET", "/sites", params={"search": query})
    return [_site_summary(s) for s in data.get("value", [])]


@mcp.tool()
async def get_site_by_url(site_url: str) -> dict:
    """Look up a site from its URL, e.g. https://contoso.sharepoint.com/sites/Marketing."""
    u = urlparse(site_url)
    path = u.path.rstrip("/")
    data = await graph("GET", f"/sites/{u.hostname}:{path}" if path else f"/sites/{u.hostname}")
    return _site_summary(data)


@mcp.tool()
async def list_lists(site_id: str) -> list[dict]:
    """List the lists and document libraries in a site."""
    data = await graph("GET", f"/sites/{site_id}/lists")
    return [
        {"id": l["id"], "displayName": l.get("displayName"),
         "template": l.get("list", {}).get("template"), "webUrl": l.get("webUrl")}
        for l in data.get("value", [])
    ]


@mcp.tool()
async def list_columns(site_id: str, list_id: str) -> list[dict]:
    """Show a list's columns (internal names are what create/update expect)."""
    data = await graph("GET", f"/sites/{site_id}/lists/{list_id}/columns")
    return [
        {"name": c.get("name"), "displayName": c.get("displayName"), "readOnly": c.get("readOnly")}
        for c in data.get("value", [])
    ]


@mcp.tool()
async def get_list_items(site_id: str, list_id: str, top: int = 50) -> list[dict]:
    """Get items (with field values) from a list. top is capped at 200."""
    data = await graph(
        "GET",
        f"/sites/{site_id}/lists/{list_id}/items",
        params={"$expand": "fields", "$top": min(top, 200)},
    )
    return [{"id": i["id"], "fields": i.get("fields", {})} for i in data.get("value", [])]


@mcp.tool()
async def create_list_item(site_id: str, list_id: str, fields: dict) -> dict:
    """Create a list item. fields maps column internal names to values."""
    data = await graph("POST", f"/sites/{site_id}/lists/{list_id}/items", json={"fields": fields})
    return {"id": data.get("id"), "fields": data.get("fields", {})}


@mcp.tool()
async def update_list_item(site_id: str, list_id: str, item_id: str, fields: dict) -> dict:
    """Update fields on an existing list item. Only the given columns change."""
    return await graph(
        "PATCH", f"/sites/{site_id}/lists/{list_id}/items/{item_id}/fields", json=fields
    )


@mcp.tool()
async def list_pages(site_id: str) -> list[dict]:
    """List the modern site pages in a site."""
    data = await graph("GET", f"/sites/{site_id}/pages")
    return [
        {"id": p.get("id"), "name": p.get("name"), "title": p.get("title"), "webUrl": p.get("webUrl")}
        for p in data.get("value", [])
    ]


# ---------------------------------------------------------------- OAuth proxy
# Claude sends the RFC 8707 "resource" parameter, which Entra rejects alongside
# our api:// scope (AADSTS9010010). These endpoints forward Claude's sign-in
# requests to Entra unchanged except for dropping "resource" and filling in
# scopes Entra needs. Tokens still come from Entra and are verified above.

ENTRA_AUTHORIZE = f"{AUTHORITY}/oauth2/v2.0/authorize"
ENTRA_TOKEN = f"{AUTHORITY}/oauth2/v2.0/token"
DEFAULT_SCOPE = f"{API_SCOPE} offline_access"


def _with_scope(scope: str | None) -> str:
    parts = (scope or "").split()
    for needed in (API_SCOPE, "offline_access"):
        if needed not in parts:
            parts.append(needed)
    return " ".join(parts)


@mcp.custom_route("/.well-known/oauth-authorization-server", methods=["GET"])
@mcp.custom_route("/.well-known/openid-configuration", methods=["GET"])
async def as_metadata(request: Request) -> Response:
    issuer = str(AnyHttpUrl(PUBLIC_URL))  # same string the resource metadata advertises
    return JSONResponse({
        "issuer": issuer,
        "authorization_endpoint": f"{PUBLIC_URL}/authorize",
        "token_endpoint": f"{PUBLIC_URL}/token",
        "response_types_supported": ["code"],
        "grant_types_supported": ["authorization_code", "refresh_token"],
        "code_challenge_methods_supported": ["S256"],
        "token_endpoint_auth_methods_supported": ["client_secret_post", "client_secret_basic"],
        "scopes_supported": [API_SCOPE, "offline_access"],
    })


@mcp.custom_route("/authorize", methods=["GET"])
async def authorize(request: Request) -> Response:
    params = {k: v for k, v in request.query_params.items() if k != "resource"}
    params["scope"] = _with_scope(params.get("scope"))
    return RedirectResponse(f"{ENTRA_AUTHORIZE}?{urlencode(params)}", status_code=302)


@mcp.custom_route("/token", methods=["POST"])
async def token(request: Request) -> Response:
    form = {k: v for k, v in (await request.form()).items() if k != "resource"}
    form["scope"] = _with_scope(form.get("scope"))
    headers = {}
    if "authorization" in request.headers:  # client_secret_basic
        headers["Authorization"] = request.headers["authorization"]
    async with httpx.AsyncClient(timeout=30) as client:
        r = await client.post(ENTRA_TOKEN, data=form, headers=headers)
    return Response(r.content, status_code=r.status_code,
                    media_type=r.headers.get("content-type", "application/json"))


if __name__ == "__main__":
    mcp.run(transport="streamable-http")
