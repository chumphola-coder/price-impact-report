"""One-off / scheduled utility: download the RAW Master Net Change Validation
Report .xlsx directly from SharePoint (bypassing the MCP tool's row-limited
text/table parsing, which is far too slow for a 100k+ row, 100MB+ file).

Reuses the sharepoint-msagadev MCP server's own cached cookie session so no
new interactive login is needed as long as that session hasn't expired.

Run with the sharepoint-msagadev project's own venv (has httpx/pyyaml):
    "/Users/camornpi/Documents/Cisco MCP project/sharepoint-msagadev/.venv/bin/python3" \
        fetch_master_report_raw.py "<sharepoint file path or URL>" <output_path>
"""
from __future__ import annotations

import asyncio
import os
import sys
from pathlib import Path

# Must match the "sharepoint" MCP server's env in .vscode/mcp.json, otherwise
# Config() resolves no tenant, uses a different (empty) session cache dir, and
# silently tries a full interactive re-auth (opens a real browser) instead of
# reusing the already-authenticated cached cookie session.
os.environ.setdefault("SHAREPOINT_TENANT", "cisco")

SHAREPOINT_MCP_SRC = "/Users/camornpi/Documents/Cisco MCP project/sharepoint-msagadev/src"
sys.path.insert(0, SHAREPOINT_MCP_SRC)

from sharepoint_mcp.client import (  # noqa: E402
    SharePointClient,
    _encode_sp_path,
    extract_web_path,
    normalize_file_path_input,
    resolve_origin_for_site_path,
)
from sharepoint_mcp.config import Config  # noqa: E402


async def download_raw(file_path: str, out_path: Path) -> None:
    config = Config()
    client = SharePointClient(config)
    try:
        info = client._resolve_site_info(path_or_file=file_path)
        cached = config.get_session(info.host)
        if not cached or not Config.is_session_valid(cached):
            raise RuntimeError(
                f"No valid cached SharePoint session for host {info.host!r} — "
                "refusing to proceed (would otherwise silently open an interactive "
                "browser login). Re-authenticate via the sharepoint MCP tool first."
            )
        normalized = normalize_file_path_input(file_path, info.site_path, info.library)
        web_path = extract_web_path(normalized) or info.site_path or ""
        site_url_for_origin = info.source_url or config.get_site(info.kind).resolution.url
        resolved_origin = resolve_origin_for_site_path(site_url_for_origin or info.host, web_path)
        host = resolved_origin or info.host
        session = await client._ensure_session(host, web_path)

        encoded = _encode_sp_path(normalized)
        url = f"{host}{web_path}/_api/web/GetFileByServerRelativeUrl('{encoded}')/$value"
        resp = await client._request(
            "GET", url, session, host,
            headers={"Accept": "application/octet-stream"},
            timeout=600.0,
        )
        if resp.status_code >= 400:
            raise RuntimeError(f"Download failed: HTTP {resp.status_code}")

        out_path.write_bytes(resp.content)
        print(f"Saved {len(resp.content):,} bytes to {out_path}")
    finally:
        await client.close()


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print(f"Usage: {sys.argv[0]} <sharepoint_file_path_or_url> <output_path>")
        sys.exit(1)
    asyncio.run(download_raw(sys.argv[1], Path(sys.argv[2])))
