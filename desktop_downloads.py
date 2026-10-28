#!/usr/bin/env python3
"""Weekly snapshots of the latest two stable releases of both products."""

import argparse
from copy import copy
from datetime import datetime
import logging
import os
from pathlib import Path
import re
import tempfile
from urllib.parse import quote, urlparse, urljoin
from zoneinfo import ZoneInfo

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import xlsxwriter
from openpyxl import load_workbook
from openpyxl.workbook.properties import CalcProperties
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


def collect_recent(client, tag="", repository=REPOSITORY):
    latest = collect(client, tag, repository)
    cutoff = latest[0]["published_at"]
    candidates, page = [], 1
    while True:
        batch = request_json(client, "GET", f"https://api.github.com/repos/{repository}/releases",
                             params={"per_page": 100, "page": page})
        if not isinstance(batch, list):
            raise ReportError("Invalid GitHub releases response")
        candidates.extend(item for item in batch if not item.get("draft")
                          and not item.get("prerelease") and item.get("published_at")
                          and item["id"] != latest[0]["id"] and item["published_at"] <= cutoff)
        if len(batch) < 100:
            break
        page += 1
    if not candidates:
        return [latest]
    previous = max(candidates, key=lambda item: (item["published_at"], item["id"]))
    return [latest, collect(client, previous["tag_name"], repository)]


def write_product(book, sheet, release, assets, collected_at, column_offset=0,
                  product="Desktop Editors", repository=REPOSITORY, row_offset=0,
                  release_label=""):
    """Write an independent five-column product block on the shared worksheet."""
    buckets = {(group, platform): [] for group in GROUPS for platform in PLATFORMS}
    for asset in assets:
        buckets[classify(asset["name"])].append(asset)

    def address(value):
        def shift(match):
            column = xl_cell_to_rowcol(match.group(1) + "1")[1]
            row = str(int(match.group(2)) + row_offset) if match.group(2) else ""
            return xl_col_to_name(column + column_offset) + row
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
    text = book.add_format(base)
    number = book.add_format(dict(base, num_format="#,##0", align="right"))
    alternate = book.add_format(dict(base, bg_color="#F5F8FB"))
    alternate_number = book.add_format(dict(base, bg_color="#F5F8FB", num_format="#,##0", align="right"))
    total = book.add_format(dict(base, bold=True, top=1, top_color="#CBD5E1", num_format="#,##0", bg_color="#E8EFF5"))
    label = f" — {release_label}" if release_label else ""
    sheet.merge_range(address("A1:E2"), f"{product} {release['tag_name']}{label}", title)
    published = datetime.fromisoformat(release["published_at"].replace("Z", "+00:00"))
    sheet.merge_range(address("A3:E3"), f"Released: {published.astimezone(collected_at.tzinfo):%d.%m.%Y}   ·   Updated: {collected_at:%d.%m.%Y %H:%M %Z}", note)
    sheet.merge_range(address("A4:E4"), f"Source: GitHub / {repository}. Lifetime download counts.", note)
    row = row_offset + 5
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
            platform_totals.append(xl_rowcol_to_cell(row, column_offset + 4))
            group_count += count
            row += 1
        sheet.merge_range(row, column_offset, row, column_offset + 3, f"Total — {group.lower()}", total)
        sheet.write_formula(row, column_offset + 4, f"=SUM({','.join(platform_totals)})", total, group_count)
        row += 2
    return row


def snapshot_name(existing, collected_at):
    name = collected_at.strftime("%Y-%m-%d")
    if name not in existing:
        return name
    base = collected_at.strftime("%Y-%m-%d %H-%M-%S")
    name, number = base, 2
    while name in existing:
        name = f"{base}-{number}"
        number += 1
    return name


def append_snapshot(history, snapshot, name):
    source_book = load_workbook(snapshot)
    source = source_book.active
    target = history.create_sheet(name)
    for row in source:
        for cell in row:
            new = target.cell(cell.row, cell.column, cell.value)
            new.data_type = cell.data_type
            if cell.has_style:
                for attribute in ("font", "fill", "border", "alignment", "protection", "number_format"):
                    setattr(new, attribute, copy(getattr(cell, attribute)))
    for merged in source.merged_cells.ranges:
        target.merge_cells(str(merged))
    for key, dimension in source.column_dimensions.items():
        target.column_dimensions[key] = copy(dimension)
    for key, dimension in source.row_dimensions.items():
        target.row_dimensions[key] = copy(dimension)
    for attribute in ("sheet_format", "sheet_properties", "views", "page_setup", "page_margins", "print_options"):
        setattr(target, attribute, copy(getattr(source, attribute)))
    target.print_area = f"A1:K{source.max_row}"
    for sheet in history:
        sheet.sheet_view.tabSelected = False
    target.sheet_view.tabSelected = True
    history.active = len(history.worksheets) - 1
    history.calculation = CalcProperties(calcMode="auto", fullCalcOnLoad=True, forceFullCalc=True)
    source_book.close()
    history.save(snapshot)


def write_report(path, release, assets, collected_at, server=None, previous=None,
                 server_previous=None, history_path=None):
    """Append a dated snapshot; never replace the existing history with an empty workbook."""
    history = None
    if history_path is not None:
        try:
            history = load_workbook(history_path)
        except Exception as exc:
            raise ReportError("Existing report cannot be read; history will not be overwritten") from exc
    name = snapshot_name(history.sheetnames if history else [], collected_at)
    path.parent.mkdir(parents=True, exist_ok=True)
    handle, temporary = tempfile.mkstemp(suffix=".xlsx", dir=path.parent)
    os.close(handle)
    try:
        with xlsxwriter.Workbook(temporary, {"strings_to_formulas": False,
                                             "strings_to_urls": False}) as book:
            sheet = book.add_worksheet(name)
            sheet.hide_gridlines(2)
            sheet.set_default_row(24)
            sheet.set_zoom(70 if server else 90)
            last_row = write_product(book, sheet, release, assets, collected_at, release_label="Latest")
            if server:
                sheet.set_column("F:F", 4)
                server_release, server_assets = server
                last_row = max(last_row, write_product(
                    book, sheet, server_release, server_assets, collected_at,
                    column_offset=6, product="Document Server", repository="ONLYOFFICE/DocumentServer",
                    release_label="Latest"))
            if previous or server_previous:
                start = last_row + 1
                if previous:
                    last_row = write_product(book, sheet, *previous, collected_at,
                                             row_offset=start, release_label="Previous")
                if server_previous:
                    last_row = max(last_row, write_product(book, sheet, *server_previous, collected_at,
                        column_offset=6, product="Document Server", repository="ONLYOFFICE/DocumentServer",
                        row_offset=start, release_label="Previous"))
            sheet.set_landscape() if server else sheet.set_portrait()
            sheet.fit_to_pages(1, 0)
            sheet.print_area(0, 0, last_row - 1, 10 if server else 4)
        if history:
            append_snapshot(history, temporary, name)
        os.replace(temporary, path)
    finally:
        if history:
            history.close()
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

    def download(self, path, info):
        LOG.info("Downloading existing workbook: fileId=%s", info['id'])
        link = self.call("GET", f"/files/file/{info['id']}/presigneduri")
        if not isinstance(link, str):
            raise ReportError("DocSpace returned an invalid download link")
        url = urljoin(self.root, link)
        portal = urlparse(self.root)
        for _ in range(6):
            parsed = urlparse(url)
            if parsed.scheme != "https" or not parsed.netloc or parsed.username is not None:
                raise ReportError("DocSpace returned an invalid download URL")
            same_origin = (parsed.hostname, parsed.port or 443) == (portal.hostname, portal.port or 443)
            # Authorize portal handlers, but never send the API key to storage hosts.
            client = self.client if same_origin else requests
            with client.get(url, timeout=(15, 120), stream=True, allow_redirects=False) as response:
                if response.status_code in (301, 302, 303, 307, 308):
                    location = response.headers.get("Location")
                    if not location:
                        raise ReportError("History download redirect has no destination")
                    url = urljoin(url, location)
                    continue
                if not 200 <= response.status_code < 300:
                    raise ReportError(f"History download failed: HTTP {response.status_code} "
                                      f"from {parsed.hostname}; fileId={info['id']}")
                with path.open("wb") as output:
                    for chunk in response.iter_content(1024 * 1024):
                        output.write(chunk)
                return
        raise ReportError("History download exceeded the redirect limit")

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
        with tempfile.TemporaryDirectory(prefix="onlyoffice-history-") as directory:
            history = args.output if args.output.exists() and not args.publish else None
            if portal and expected:
                history = Path(directory) / "history.xlsx"
                portal.download(history, expected)
            github = session(os.getenv("GITHUB_TOKEN"), github=True)
            desktop = collect_recent(github, args.tag)
            server = collect_recent(github, args.server_tag, "ONLYOFFICE/DocumentServer")
            now = datetime.now(ZoneInfo(args.timezone))
            write_report(args.output, *desktop[0], now, server=server[0],
                         previous=desktop[1] if len(desktop) > 1 else None,
                         server_previous=server[1] if len(server) > 1 else None,
                         history_path=history)
            for product, releases in (("Desktop Editors", desktop), ("Document Server", server)):
                for release, assets in releases:
                    LOG.info("%s %s: %d assets, %s downloads", product, release["tag_name"],
                             len(assets), sum(item["download_count"] for item in assets))
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
