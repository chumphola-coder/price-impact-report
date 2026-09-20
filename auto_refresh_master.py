"""
auto_refresh_master.py
=======================
Fully unattended pipeline: check the Webex "MyAgent" bot DM for a newer
Master Net Change Validation Report announcement, and if one is found,
download it from SharePoint, distill it, update the bundled default, and
commit + push -- with zero manual steps (per user request, Sep 20 2026).

Run manually to test:
    ./.venv/bin/python3 auto_refresh_master.py [--dry-run]

Scheduled via launchd, see com.camornpi.ccwr-master-refresh.plist
(runs this script roughly every 20 days, matching the user's established
cadence -- pricing changes are monthly at most).

Logs every run to auto_refresh_master.log in this directory.
"""
from __future__ import annotations

import asyncio
import json
import logging
import os
import re
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
LOG_PATH = HERE / "auto_refresh_master.log"

logging.basicConfig(
    filename=LOG_PATH, level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
)
log = logging.getLogger("auto_refresh_master")

# Reuse the webex-mcp / sharepoint-msagadev packages' own OAuth/cookie auth
# directly (same technique as fetch_master_report_raw.py) instead of
# re-implementing login flows here.
WEBEX_MCP_DIR = Path("/Users/camornpi/Documents/Cisco MCP project/webex-mcp")
SHAREPOINT_MCP_SRC = "/Users/camornpi/Documents/Cisco MCP project/sharepoint-msagadev/src"
sys.path.insert(0, str(WEBEX_MCP_DIR / "src"))
sys.path.insert(0, SHAREPOINT_MCP_SRC)

# webex_mcp's server.py calls load_dotenv() with no path, which only finds
# a .env in the CURRENT working directory -- load it explicitly instead of
# relying on cwd, since this script's cwd is the CCWR project, not webex-mcp.
from dotenv import load_dotenv  # noqa: E402
load_dotenv(WEBEX_MCP_DIR / ".env")

# Must be set BEFORE importing sharepoint_mcp.config -- see
# fetch_master_report_raw.py for why (mirrors the "sharepoint" MCP server's
# own env in ~/Documents/Cisco MCP project/.vscode/mcp.json).
os.environ.setdefault("SHAREPOINT_TENANT", "cisco")

from webex_mcp.client import WebexClient  # noqa: E402
from sharepoint_mcp.client import (  # noqa: E402
    SharePointClient,
    _encode_sp_path,
    extract_web_path,
    normalize_file_path_input,
    resolve_origin_for_site_path,
)
from sharepoint_mcp.config import Config as SharePointConfig  # noqa: E402

sys.path.insert(0, str(HERE))
from distill_master_report import distill  # noqa: E402

MYAGENT_ROOM_ID = "Y2lzY29zcGFyazovL3VzL1JPT00vZjY1YmRjNzAtYTYxNy0xMWYxLTk5N2MtNDUzMWYxYmVmYTIz"
SHAREPOINT_SITE = "https://cisco.sharepoint.com/sites/CXPricingOperations"
DEFAULT_MASTER_CSV = HERE / "master_net_change_validation_report.csv.gz"
DEFAULT_MASTER_META = HERE / "master_net_change_validation_report.meta.json"
RAW_TMP = HERE / "_auto_refresh_raw.xlsx"

# Safety guard: refuse to auto-apply an update whose distilled SKU count is
# suspiciously low vs. the current bundled file (guards against a truncated
# or malformed download silently degrading the live app). Current real
# baseline is 728,410 SKUs (Sep 2026).
MIN_SKU_FRACTION_OF_CURRENT = 0.5

# Matches the real MyAgent message format, e.g.:
#   "... Effective Date: October 3, 2026 ... Excel Report File:
#    [Master Net Change Validation Report 03-Oct-2026.xlsx](https://...)"
ANNOUNCEMENT_RE = re.compile(
    r"Effective Date:\s*([A-Za-z]+ \d{1,2},\s*\d{4}).*?"
    r"Excel Report File:\s*\n?\[?([^\]\n(]+?\.xlsx)",
    re.DOTALL,
)


async def _find_latest_announcement() -> tuple[str, str] | None:
    """Return (report_filename, effective_date_str) from the most recent
    real (non-test) MyAgent announcement, or None if none is found."""
    async with WebexClient() as client:
        result = await client.get(
            "/messages", params={"roomId": MYAGENT_ROOM_ID, "max": 20}, paginate=False
        )
    messages = result.get("items", []) if isinstance(result, dict) else result
    for msg in messages:  # Webex returns newest-first
        text = msg.get("text", "") or msg.get("markdown", "") or ""
        if "test notification" in text.lower():
            continue
        m = ANNOUNCEMENT_RE.search(text)
        if m:
            return m.group(2).strip(), m.group(1).strip()
    return None


async def _download_raw(file_path: str, out_path: Path) -> None:
    """Download the real report file straight from SharePoint (raw bytes,
    bypassing the row-limited table-text parser). Refuses to proceed if
    there's no already-valid cached cookie session -- this pipeline is only
    meant to run unattended when auth is already warm; it must never block
    on an interactive login."""
    config = SharePointConfig()
    client = SharePointClient(config)
    try:
        info = client._resolve_site_info(path_or_file=file_path)
        cached = config.get_session(info.host)
        if not cached or not SharePointConfig.is_session_valid(cached):
            raise RuntimeError(
                f"No valid cached SharePoint session for host {info.host!r} -- "
                "refusing to proceed unattended (needs a fresh interactive login "
                "via the sharepoint MCP tool first)."
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
            headers={"Accept": "application/octet-stream"}, timeout=600.0,
        )
        if resp.status_code >= 400:
            raise RuntimeError(f"Download failed: HTTP {resp.status_code}")
        out_path.write_bytes(resp.content)
    finally:
        await client.close()


def _run_git(*args: str) -> subprocess.CompletedProcess:
    # GIT_TERMINAL_PROMPT=0 makes git fail fast instead of hanging/blocking
    # if it would otherwise need an interactive credential prompt.
    env = {**os.environ, "GIT_TERMINAL_PROMPT": "0"}
    return subprocess.run(["git", *args], cwd=HERE, capture_output=True, text=True, env=env)


def _current_sku_count() -> int:
    if not DEFAULT_MASTER_CSV.exists():
        return 0
    import pandas as pd
    return len(pd.read_csv(DEFAULT_MASTER_CSV, compression="infer"))


async def _notify_self(text: str) -> None:
    """Best-effort confirmation message back to the same MyAgent DM thread
    so the user has visibility into what the automation did, without
    needing to check logs. Never lets a notify failure fail the whole run."""
    try:
        async with WebexClient() as client:
            await client.post("/messages", json={"roomId": MYAGENT_ROOM_ID, "markdown": text})
    except Exception:
        log.exception("Failed to send Webex notification (non-fatal).")


async def main(dry_run: bool = False) -> int:
    log.info("=== auto_refresh_master run start (dry_run=%s) ===", dry_run)
    try:
        found = await _find_latest_announcement()
        if not found:
            log.info("No real price-change announcement found in MyAgent room. Nothing to do.")
            return 0
        report_filename, effective_date = found

        current_meta = (
            json.loads(DEFAULT_MASTER_META.read_text()) if DEFAULT_MASTER_META.exists() else {}
        )
        if current_meta.get("original_filename") == report_filename:
            log.info("Already up to date (%s). Nothing to do.", report_filename)
            return 0

        log.info(
            "New report detected: %s (effective %s) -- current bundled: %s",
            report_filename, effective_date, current_meta.get("original_filename"),
        )

        file_path = f"{SHAREPOINT_SITE}/Shared Documents/{report_filename}"
        await _download_raw(file_path, RAW_TMP)

        before_count = _current_sku_count()
        new_count = distill(str(RAW_TMP), report_filename, DEFAULT_MASTER_CSV, DEFAULT_MASTER_META)
        RAW_TMP.unlink(missing_ok=True)

        if before_count and new_count < before_count * MIN_SKU_FRACTION_OF_CURRENT:
            log.error(
                "SAFETY ABORT: new SKU count %d is under %.0f%% of current %d -- "
                "reverting, needs manual review (NOT committed/pushed).",
                new_count, MIN_SKU_FRACTION_OF_CURRENT * 100, before_count,
            )
            _run_git("checkout", "--", str(DEFAULT_MASTER_CSV), str(DEFAULT_MASTER_META))
            await _notify_self(
                f"\u26a0\ufe0f Auto-refresh SAFETY ABORT: {report_filename} distilled to only "
                f"{new_count:,} SKUs vs current {before_count:,} -- left unchanged, please check manually."
            )
            return 1

        log.info("Distilled %d unique SKUs (previous: %d).", new_count, before_count)

        if dry_run:
            log.info("Dry run: not committing or pushing.")
            _run_git("checkout", "--", str(DEFAULT_MASTER_CSV), str(DEFAULT_MASTER_META))
            return 0

        _run_git("add", "--", str(DEFAULT_MASTER_CSV), str(DEFAULT_MASTER_META))
        commit = _run_git(
            "commit", "-m",
            f"Auto-refresh master report: {report_filename} ({new_count:,} SKUs, effective {effective_date})",
        )
        log.info("git commit: rc=%d stdout=%s stderr=%s",
                 commit.returncode, commit.stdout.strip(), commit.stderr.strip())
        if commit.returncode != 0:
            await _notify_self(f"\u26a0\ufe0f Auto-refresh: git commit failed for {report_filename}, check logs.")
            return 1

        push = _run_git("push", "origin", "main")
        log.info("git push: rc=%d stdout=%s stderr=%s",
                 push.returncode, push.stdout.strip(), push.stderr.strip())
        if push.returncode != 0:
            log.error("Push failed -- commit is local only, needs a manual 'git push origin main'.")
            await _notify_self(
                f"\u26a0\ufe0f Auto-refresh: downloaded + committed {report_filename} "
                f"({new_count:,} SKUs) but git push failed -- please run "
                "'git push origin main' manually."
            )
            return 1

        log.info("Success: pushed new master report (%s, %d SKUs) live.", report_filename, new_count)
        await _notify_self(
            f"\u2705 Price Change Impact web app auto-updated: **{report_filename}** "
            f"({new_count:,} SKUs, effective {effective_date}) is now live."
        )
        return 0
    except Exception:
        log.exception("auto_refresh_master run failed")
        return 1


if __name__ == "__main__":
    rc = asyncio.run(main(dry_run="--dry-run" in sys.argv))
    sys.exit(rc)
