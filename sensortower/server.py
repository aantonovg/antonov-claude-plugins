# /// script
# requires-python = ">=3.12"
# dependencies = [
#     "fastmcp>=2.0",
#     "httpx[socks]",
#     "pyyaml",
# ]
# ///

import json
import os
import sys
import tempfile
import time
from pathlib import Path

import httpx
import yaml
from fastmcp import FastMCP

SPEC_PATH = Path(__file__).parent / "sensortower_openapi.yaml"
_STATIC_PREFIX = "/api/docs/static/"
_STATIC_HOST = "app.sensortower.com"

with open(SPEC_PATH) as f:
    spec = yaml.safe_load(f)

# Strip response schemas so FastMCP won't validate bodies replaced by _smart_response
for _path_data in spec.get("paths", {}).values():
    for _method_data in _path_data.values():
        if isinstance(_method_data, dict):
            for _resp in _method_data.get("responses", {}).values():
                for _content in _resp.get("content", {}).values():
                    _content.pop("schema", None)

api_key = os.environ.get("SENSORTOWER_API_KEY", "")
if api_key:
    print(f"✔️ SENSORTOWER_API_KEY loaded ({len(api_key)} chars)", file=sys.stderr)
else:
    print("⚠️ SENSORTOWER_API_KEY is empty — API calls will fail", file=sys.stderr)


# Endpoints/params that REQUIRE repeated keys (?k=a&k=b), not CSV.
# Everything else is collapsed to CSV (SensorTower default).
_REPEATED_PARAMS: set[tuple[str, str]] = set()

# Endpoint-specific param RENAMES. OpenAPI spec drift: some endpoints document
# `countries` but the API wants a different shape. Note: app_overlap `countries`
# must be sent as a Rails-style bracketed array: `countries[]=US`.
_PARAM_RENAMES: dict[str, dict[str, str]] = {
    "/v1/unified/app_overlap": {"countries": "countries[]"},
}


def _unpack_array_value(v: str) -> list[str]:
    """FastMCP sometimes sends array params as JSON-encoded strings like `["US"]`.
    Decode them back to a list of individual values."""
    if len(v) >= 2 and v.startswith("[") and v.endswith("]"):
        try:
            arr = json.loads(v)
            if isinstance(arr, list):
                return [str(x) for x in arr]
        except json.JSONDecodeError:
            pass
    return [v]


async def _fix_array_params(request: httpx.Request):
    """SensorTower expects comma-separated arrays (a=1,2), but FastMCP sends
    arrays either as repeated keys (a=1&a=2) or as a single JSON-encoded string
    (a=["1","2"]). Normalize to CSV — except endpoints/params listed in
    `_REPEATED_PARAMS`, which REQUIRE repeated keys (a=1&a=2)."""
    seen: dict[str, list[str]] = {}
    for key, value in request.url.params.multi_items():
        for v in _unpack_array_value(value):
            seen.setdefault(key, []).append(v)
    path = request.url.path
    renames = _PARAM_RENAMES.get(path, {})
    fixed_multi: list[tuple[str, str]] = []
    for k, vs in seen.items():
        out_key = renames.get(k, k)
        if (path, out_key) in _REPEATED_PARAMS:
            for v in vs:
                fixed_multi.append((out_key, v))
        else:
            fixed_multi.append((out_key, ",".join(vs) if len(vs) > 1 else vs[0]))
    request.url = request.url.copy_with(params=fixed_multi)


async def _inject_auth(request: httpx.Request):
    """Add auth_token to API requests. Static lookups are public — skip them."""
    if request.url.path.startswith(_STATIC_PREFIX):
        return
    request.url = request.url.copy_merge_params({"auth_token": api_key})


_INLINE_THRESHOLD = 4000  # chars; larger responses are saved to file
_DUMP_DIR = Path(tempfile.gettempdir()) / "claude" / "sensortower"


def _smart_response(cleaned: str, request: httpx.Request) -> str:
    """Return response inline if small, otherwise save to file."""
    if len(cleaned) <= _INLINE_THRESHOLD:
        return cleaned
    _DUMP_DIR.mkdir(parents=True, exist_ok=True)
    slug = request.url.path.strip("/").replace("/", "_")
    path = _DUMP_DIR / f"{slug}_{int(time.time())}.json"
    path.write_text(cleaned, encoding="utf-8")
    return json.dumps({
        "saved_to": str(path),
        "size_chars": len(cleaned),
        "hint": (
            "Response too large for inline. Read the file to inspect. "
            "Ask the user if they want to copy results to a project folder "
            "for future sessions — include the query parameters used."
        ),
    })


_STRIP_KEYS = {
    "custom_tags",
    "canonical_country",
}


def _strip_bloat(obj):
    """Recursively strip bulky/redundant fields from API responses."""
    if isinstance(obj, dict):
        for key in _STRIP_KEYS & obj.keys():
            del obj[key]
        for v in obj.values():
            _strip_bloat(v)
    elif isinstance(obj, list):
        for item in obj:
            _strip_bloat(item)


class _RoutingTransport(httpx.AsyncBaseTransport):
    """Static lookups live on app.sensortower.com (public, no auth).
    Everything else goes to api.sensortower.com.

    httpx ignores host changes via copy_with() at the transport level,
    so we use a separate AsyncClient for static requests.

    All JSON responses are cleaned: custom_tags are stripped to reduce payload."""

    def __init__(self):
        self._api = httpx.AsyncHTTPTransport()
        self._app = httpx.AsyncClient(
            base_url=f"https://{_STATIC_HOST}", timeout=60,
        )

    async def handle_async_request(self, request: httpx.Request) -> httpx.Response:
        if request.url.path.startswith(_STATIC_PREFIX):
            resp = await self._app.get(str(request.url.path))
        else:
            resp = await self._api.handle_async_request(request)

        # Strip custom_tags to keep responses manageable
        await resp.aread()
        try:
            data = json.loads(resp.content)
            _strip_bloat(data)
            cleaned = json.dumps(data, ensure_ascii=False)
            body = _smart_response(cleaned, request)
            headers = {
                k: v for k, v in resp.headers.items()
                if k.lower() not in ("content-encoding", "content-length")
            }
            return httpx.Response(
                status_code=resp.status_code,
                headers=headers,
                content=body.encode("utf-8"),
            )
        except (json.JSONDecodeError, UnicodeDecodeError):
            return resp


client = httpx.AsyncClient(
    base_url="https://api.sensortower.com",
    transport=_RoutingTransport(),
    event_hooks={"request": [_fix_array_params, _inject_auth]},
    timeout=60,
)

mcp = FastMCP.from_openapi(
    openapi_spec=spec,
    client=client,
    name="Sensor Tower",
)

if __name__ == "__main__":
    mcp.run()
