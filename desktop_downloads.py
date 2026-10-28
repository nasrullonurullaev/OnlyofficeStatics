#!/usr/bin/env python3
"""Desktop Editors and Document Server downloads; one workbook, one worksheet."""

import argparse
from datetime import datetime
import logging
import os
from pathlib import Path
import re
import tempfile
from urllib.parse import quote, urlparse
from zoneinfo import ZoneInfo

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell, xl_cell_to_rowcol, xl_col_to_name


LOG = logging.getLogger("desktop-downloads")
GROUPS = ("Installations", "Updates", "Add-ons")
PLATFORMS = ("Windows", "Linux", "macOS")
REPOSITORY = "ONLYOFFICE/DesktopEditors"


class ReportError(Exception):
    pass


def session(token=None, github=False):
    client = requests.Session()
    client.headers.update({"User-Agent": "onlyoffice-desktop-downloads/1.0"})
    if token:
        client.headers["Authorization"] = f"Bearer {token}"
    if github:
        client.headers.update({"Accept": "application/vnd.github+json",
                               "X-GitHub-Api-Version": "2022-11-28"})
    # Never retry a POST: a lost response can mean the version was already saved.
    client.mount("https://", HTTPAdapter(max_retries=Retry(
        total=3, backoff_factor=1, status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=frozenset(("GET",)), respect_retry_after_header=True)))
    return client


def request_json(client, method, url, **kwargs):
    response = client.request(method, url, timeout=(15, 120),
                              allow_redirects=False, **kwargs)
    # Do not log response bodies, authorization headers or signed links.
    if not 200 <= response.status_code < 300:
        raise ReportError(f"{method} {urlparse(url).path}: HTTP {response.status_code}")
    try:
        return response.json()
    except ValueError as exc:
        raise ReportError("API returned an invalid JSON response") from exc


def classify(name):
    lower = name.lower()
    if lower.startswith("addon-"):
        group = "Add-ons"
    elif lower.endswith(".zip"):
        group = "Updates"
    else:
        group = "Installations"

    if lower.startswith("addon-win-") or lower.endswith((".exe", ".msi")):
        platform = "Windows"
    elif lower.startswith("addon-mac-") or lower.endswith((".dmg", ".pkg")):
        platform = "macOS"
    elif lower.startswith("addon-linux-") or lower.endswith(
            (".deb", ".rpm", ".appimage", ".tar.xz", ".tar.gz", ".snap", ".flatpak")):
        platform = "Linux"
    elif re.fullmatch(r"desktopeditors_(x64|x86|arm|arm64)\.zip", lower):
        platform = "Windows"
    else:
        raise ReportError(f"Unknown platform for asset {name!r}; update classification")
    return group, platform


def collect(client, tag="", repository=REPOSITORY):
    root = f"https://api.github.com/repos/{repository}"
    endpoint = f"tags/{quote(tag, safe='')}" if tag else "latest"
    release = request_json(client, "GET", f"{root}/releases/{endpoint}")
    if release.get("draft") or release.get("prerelease") or not release.get("published_at"):
        raise ReportError("A published stable release is required")
    release_id = release["id"]
    assets, page = [], 1
    # Explicitly paginate assets; the list embedded in a release can be truncated.
    while True:
        batch = request_json(client, "GET", f"{root}/releases/{release_id}/assets",
                             params={"per_page": 100, "page": page})
        if not isinstance(batch, list):
            raise ReportError("Invalid GitHub assets response")
        assets.extend(batch)
        if len(batch) < 100:
            break
        page += 1
    if not assets:
        raise ReportError("Release has no assets; existing report will not be replaced")
    seen = set()
    for asset in assets:
        if asset["id"] in seen:
            raise ReportError("GitHub returned a duplicate asset ID")
        seen.add(asset["id"])
        count = asset.get("download_count")
        if type(count) is not int or count < 0 or asset.get("state") != "uploaded":
            raise ReportError(f"Invalid counter or incomplete asset: {asset['name']}")
        classify(asset["name"])
    return release, assets


def write_product(book, sheet, release, assets, collected_at, column_offset=0,
                  product="Desktop Editors", repository=REPOSITORY):
    """Write an independent five-column product block on the shared worksheet."""
    buckets = {(group, platform): [] for group in GROUPS for platform in PLATFORMS}
    for asset in assets:
        buckets[classify(asset["name"])].append(asset)

    def address(value):
        def shift(match):
            column = xl_cell_to_rowcol(match.group(1) + "1")[1]
            return xl_col_to_name(column + column_offset) + match.group(2)
        return re.sub(r"([A-Z]+)([0-9]*)", shift, value)

    download_column = xl_col_to_name(column_offset + 4)
    sheet.set_column(address("A:A"), 24)
    sheet.set_column(address("B:D"), 18)
    sheet.set_column(address("E:E"), 20)
    base = {"font_name": "Calibri", "font_size": 11, "font_color": "#243746", "valign": "vcenter"}
    title = book.add_format(dict(base, bold=True, font_size=22, font_color="#17365D"))
    note = book.add_format(dict(base, font_size=10, font_color="#64748B"))
    section = book.add_format(dict(base, bold=True, font_size=13, font_color="#FFFFFF", bg_color="#24476B"))
    platform_style = book.add_format(dict(base, bold=True, bg_color="#E8EFF5"))
    header = book.add_format(dict(base, bold=True, bg_color="#24476B", font_color="#FFFFFF", align="right"))
    label_header = book.add_format(dict(base, bold=True, bg_color="#24476B", font_color="#FFFFFF"))
    text = book.add_format(base)
    number = book.add_format(dict(base, num_format="#,##0", align="right"))
    alternate = book.add_format(dict(base, bg_color="#F5F8FB"))
    alternate_number = book.add_format(dict(base, bg_color="#F5F8FB", num_format="#,##0", align="right"))
    total = book.add_format(dict(base, bold=True, top=1, top_color="#CBD5E1", num_format="#,##0", bg_color="#E8EFF5"))
    sheet.merge_range(address("A1:E2"), f"{product} {release['tag_name']}", title)
    published = datetime.fromisoformat(release["published_at"].replace("Z", "+00:00"))
    sheet.merge_range(address("A3:E3"), f"Released: {published.astimezone(collected_at.tzinfo):%d.%m.%Y}   ·   Updated: {collected_at:%d.%m.%Y %H:%M %Z}", note)
    sheet.merge_range(address("A4:E4"), f"Source: GitHub / {repository}. Lifetime download counts.", note)
    sheet.merge_range(address("A6:E6"), "Download summary", platform_style)
    sheet.write(address("A7"), "Platform", label_header)
    for column, label in enumerate((*GROUPS, "Total"), 1):
        sheet.write(6, column_offset + column, label, header)
    sheet.merge_range(address("A13:E13"), "Downloads by file", platform_style)
    row, group_totals, platform_cells = 14, {}, {}
    for group in GROUPS:
        if not any(buckets[group, platform] for platform in PLATFORMS):
            continue
        sheet.merge_range(row, column_offset, row, column_offset + 3, group, section)
        sheet.write(row, column_offset + 4, "Downloads", header)
        row += 1
        platform_totals = []
        group_count = 0
        for platform in PLATFORMS:
            entries = sorted(buckets[group, platform], key=lambda item: item["name"].lower())
            if not entries:
                continue
            sheet.merge_range(row, column_offset, row, column_offset + 4, platform, platform_style)
            row += 1
            start = row
            for index, asset in enumerate(entries):
                sheet.merge_range(row, column_offset, row, column_offset + 3, asset["name"], alternate if index % 2 == 0 else text)
                sheet.write_number(row, column_offset + 4, asset["download_count"], alternate_number if index % 2 == 0 else number)
                row += 1
            count = sum(item["download_count"] for item in entries)
            sheet.merge_range(row, column_offset, row, column_offset + 3, f"Total {platform}", total)
            sheet.write_formula(row, column_offset + 4, f"=SUM({download_column}{start + 1}:{download_column}{row})", total, count)
            platform_cells[group, platform] = xl_rowcol_to_cell(row, column_offset + 4)
            platform_totals.append(xl_rowcol_to_cell(row, column_offset + 4))
            group_count += count
            row += 1
        sheet.merge_range(row, column_offset, row, column_offset + 3, f"Total — {group.lower()}", total)
        sheet.write_formula(row, column_offset + 4, f"=SUM({','.join(platform_totals)})", total, group_count)
        group_totals[group] = xl_rowcol_to_cell(row, column_offset + 4)
        row += 2
    # Summary links to detail subtotals, so each file is counted exactly once.
    for summary_row, platform in enumerate(PLATFORMS, 7):
        sheet.write(summary_row, column_offset, platform, text)
        for column, group in enumerate(GROUPS, 1):
            count = sum(item["download_count"] for item in buckets[group, platform])
            cell = platform_cells.get((group, platform))
            sheet.write_formula(summary_row, column_offset + column, f"={cell}" if cell else "=0", number, count)
        count = sum(item["download_count"] for group in GROUPS for item in buckets[group, platform])
        sheet.write_formula(summary_row, column_offset + 4, f"=SUM({address(f'B{summary_row + 1}')}:{address(f'D{summary_row + 1}')})", total, count)
    sheet.write(address("A11"), "Total downloads", total)
    for column, group in enumerate(GROUPS, 1):
        cell = group_totals.get(group)
        count = sum(item["download_count"] for platform in PLATFORMS for item in buckets[group, platform])
        sheet.write_formula(10, column_offset + column, f"={cell}" if cell else "=0", total, count)
    sheet.write_formula(address("E11"), f"=SUM({address('B11:D11')})", total, sum(item["download_count"] for item in assets))
    return row


def write_report(path, release, assets, collected_at, server=None):
    """Rebuild rather than import: old sheets and stale rows must disappear."""
    path.parent.mkdir(parents=True, exist_ok=True)
    handle, temporary = tempfile.mkstemp(suffix=".xlsx", dir=path.parent)
    os.close(handle)
    try:
        with xlsxwriter.Workbook(temporary, {"strings_to_formulas": False,
                                             "strings_to_urls": False}) as book:
            sheet = book.add_worksheet("Statistics")
            sheet.hide_gridlines(2)
            sheet.set_default_row(24)
            sheet.set_zoom(70 if server else 90)
            last_row = write_product(book, sheet, release, assets, collected_at)
            if server:
                sheet.set_column("F:F", 4)
                server_release, server_assets = server
                last_row = max(last_row, write_product(
                    book, sheet, server_release, server_assets, collected_at,
                    column_offset=6, product="Document Server", repository="ONLYOFFICE/DocumentServer"))
            sheet.set_landscape() if server else sheet.set_portrait()
            sheet.fit_to_pages(1, 0)
            sheet.print_area(0, 0, last_row - 1, 10 if server else 4)
        os.replace(temporary, path)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)


class DocSpace:
    def __init__(self, base_url, token, folder_id):
        parsed = urlparse(base_url)
        if parsed.scheme != "https" or not parsed.netloc or parsed.username or parsed.query or parsed.fragment:
            raise ReportError("DocSpace URL must be an HTTPS URL without credentials or query")
        if not str(folder_id).isdigit() or int(folder_id) <= 0:
            raise ReportError("Set a positive native DocSpace folder ID")
        self.root = base_url.rstrip("/") + "/api/2.0"
        self.folder_id = str(folder_id)
        self.client = session(token)

    def call(self, method, route, **kwargs):
        data = request_json(self.client, method, self.root + route, **kwargs)
        if not isinstance(data, dict) or "response" not in data:
            raise ReportError("Invalid DocSpace response envelope")
        return data["response"]

    def find(self, name):
        matches, start = [], 0
        while True:
            content = self.call("GET", f"/files/{self.folder_id}", params={
                "count": 100, "startIndex": start, "withSubFolders": "false"})
            files, folders = content["files"], content.get("folders", [])
            matches.extend(item for item in files if item["title"].casefold() == name.casefold())
            received = len(files) + len(folders)
            start += received
            if received == 0 or start >= content["total"]:
                break
        if len(matches) > 1:
            raise ReportError("Several report files have the same name; resolve duplicates in DocSpace")
        if not matches:
            return None
        info = self.call("GET", f"/files/file/{matches[0]['id']}")
        if str(info["folderId"]) != self.folder_id or info["title"] != name:
            raise ReportError("Report name or destination folder does not match")
        return info

    def publish(self, path, expected):
        current = self.find(path.name)
        fingerprint = lambda item: (item["id"], item["version"]) if item else None
        if fingerprint(current) != fingerprint(expected):
            raise ReportError("Report changed during collection; rerun to avoid overwriting a concurrent change")
        with path.open("rb") as source:
            result = self.call("POST", f"/files/{self.folder_id}/upload",
                               params={"createNewIfExist": "false"},
                               files={"file": (path.name, source,
                                      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")})
        if not isinstance(result, list) or len(result) != 1:
            raise ReportError("Unexpected upload result; inspect DocSpace before retrying")
        stored = result[0]
        if expected and str(stored["id"]) != str(expected["id"]):
            raise ReportError("Upload created a different file; inspect DocSpace for a lock or duplicate")
        verified = self.find(path.name)
        if not verified or str(verified["id"]) != str(stored["id"]) or verified["version"] != stored["version"]:
            raise ReportError("Uploaded report could not be verified")
        if expected and verified["version"] <= expected["version"]:
            raise ReportError("Report version did not advance")
        LOG.info("DocSpace report updated: fileId=%s, version=%s", verified["id"], verified["version"])
        return verified


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=Path("reports/ONLYOFFICE-Download-Statistics.xlsx"))
    parser.add_argument("--tag", default="", help="Desktop Editors tag; empty means latest stable release")
    parser.add_argument("--server-tag", default="", help="Document Server tag; empty means latest stable release")
    parser.add_argument("--timezone", default="Asia/Tashkent")
    parser.add_argument("--publish", action="store_true")
    parser.add_argument("--docspace-url", default=os.getenv("DOCSPACE_URL", "https://nct2.onlyoffice.com"))
    parser.add_argument("--folder-id", default=os.getenv("DOCSPACE_FOLDER_ID", ""))
    args = parser.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    if args.output.name != "ONLYOFFICE-Download-Statistics.xlsx":
        parser.error("Output filename must be ONLYOFFICE-Download-Statistics.xlsx to keep the remote document stable")
    try:
        portal, expected = None, None
        if args.publish:
            token = os.getenv("DOCSPACE_API_KEY")
            if not token:
                raise ReportError("Set DOCSPACE_API_KEY via GitHub Actions Secrets")
            portal = DocSpace(args.docspace_url, token, args.folder_id)
            expected = portal.find(args.output.name)
        github = session(os.getenv("GITHUB_TOKEN"), github=True)
        release, assets = collect(github, args.tag)
        server_release, server_assets = collect(github, args.server_tag, "ONLYOFFICE/DocumentServer")
        now = datetime.now(ZoneInfo(args.timezone))
        write_report(args.output, release, assets, now, server=(server_release, server_assets))
        LOG.info("Desktop Editors %s: %d assets, %s downloads; report: %s", release["tag_name"],
                 len(assets), sum(item["download_count"] for item in assets), args.output)
        LOG.info("Document Server %s: %d assets, %s downloads", server_release["tag_name"],
                 len(server_assets), sum(item["download_count"] for item in server_assets))
        if portal:
            portal.publish(args.output, expected)
    except (ReportError, requests.RequestException, OSError, ValueError, KeyError) as exc:
        # requests exceptions can contain URLs; omit their content to avoid leaking secrets.
        message = str(exc) if isinstance(exc, ReportError) else type(exc).__name__
        LOG.error("%s", message)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
