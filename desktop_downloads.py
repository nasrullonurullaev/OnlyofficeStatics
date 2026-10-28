#!/usr/bin/env python3
"""Current DesktopEditors release downloads; one workbook, one worksheet."""

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
from xlsxwriter.utility import xl_rowcol_to_cell


LOG = logging.getLogger("desktop-downloads")
GROUPS = ("Инсталляции", "Обновления", "Аддоны")
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
        group = "Аддоны"
    elif lower.endswith(".zip"):
        group = "Обновления"
    else:
        group = "Инсталляции"

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


def collect(client, tag=""):
    root = f"https://api.github.com/repos/{REPOSITORY}"
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


def write_report(path, release, assets, collected_at):
    """Rebuild rather than import: old sheets and stale rows must disappear."""
    buckets = {(group, platform): [] for group in GROUPS for platform in PLATFORMS}
    for asset in assets:
        buckets[classify(asset["name"])].append(asset)
    path.parent.mkdir(parents=True, exist_ok=True)
    handle, temporary = tempfile.mkstemp(suffix=".xlsx", dir=path.parent)
    os.close(handle)
    try:
        with xlsxwriter.Workbook(temporary, {"strings_to_formulas": False,
                                             "strings_to_urls": False}) as book:
            sheet = book.add_worksheet("Статистика")
            sheet.hide_gridlines(2)
            sheet.freeze_panes(8, 0)
            sheet.set_column("A:A", 66)
            sheet.set_column("B:B", 20)
            sheet.set_column("C:C", 17)
            sheet.set_default_row(22)
            base = {"font_name": "Calibri", "font_size": 11}
            title = book.add_format(dict(base, bold=True, font_size=18, font_color="#FFFFFF", bg_color="#17365D"))
            section = book.add_format(dict(base, bold=True, font_size=13, font_color="#FFFFFF", bg_color="#24476B"))
            platform_style = book.add_format(dict(base, bold=True, bg_color="#E8EFF5"))
            text = book.add_format(base)
            number = book.add_format(dict(base, num_format="#,##0"))
            total = book.add_format(dict(base, bold=True, top=1, num_format="#,##0", bg_color="#E8EFF5"))
            link = book.add_format(dict(base, font_color="#0563C1", underline=1))
            sheet.merge_range("A1:C2", f"Desktop Editors {release['tag_name']}", title)
            published = datetime.fromisoformat(release["published_at"].replace("Z", "+00:00"))
            sheet.write("A3", f"Релиз опубликован: {published.astimezone(collected_at.tzinfo):%d.%m.%Y}", text)
            sheet.write("A4", f"Статистика собрана: {collected_at:%d.%m.%Y %H:%M %Z}", text)
            sheet.write("A5", "Накопительное число скачиваний файлов с GitHub", text)
            sheet.write_url("A6", release["html_url"], link, "Источник: GitHub — релиз DesktopEditors")
            sheet.write_row("A8", ["Файл / платформа", "Скачиваний", "Ссылка"], platform_style)
            row, group_totals = 9, []
            for group in GROUPS:
                sheet.merge_range(row, 0, row, 2, group, section)
                row += 1
                platform_totals = []
                group_count = 0
                for platform in PLATFORMS:
                    sheet.merge_range(row, 0, row, 2, platform, platform_style)
                    row += 1
                    start = row
                    entries = sorted(buckets[group, platform], key=lambda item: item["name"].lower())
                    for asset in entries:
                        sheet.write_string(row, 0, asset["name"], text)
                        sheet.write_number(row, 1, asset["download_count"], number)
                        sheet.write_url(row, 2, asset["browser_download_url"], link, "Скачать")
                        row += 1
                    count = sum(item["download_count"] for item in entries)
                    sheet.write(row, 0, f"Всего {platform}", total)
                    formula = f"=SUM(B{start + 1}:B{row})" if entries else "=0"
                    sheet.write_formula(row, 1, formula, total, count)
                    platform_totals.append(xl_rowcol_to_cell(row, 1))
                    group_count += count
                    row += 2
                sheet.write(row, 0, f"Всего — {group.lower()}", total)
                sheet.write_formula(row, 1, f"=SUM({','.join(platform_totals)})", total, group_count)
                group_totals.append(xl_rowcol_to_cell(row, 1))
                row += 3
            sheet.write(row, 0, "Всего скачиваний", total)
            sheet.write_formula(row, 1, f"=SUM({','.join(group_totals)})", total,
                                sum(item["download_count"] for item in assets))
            sheet.set_landscape()
            sheet.fit_to_pages(1, 0)
            sheet.print_area(0, 0, row, 2)
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
    parser.add_argument("--output", type=Path, default=Path("reports/DesktopEditors.xlsx"))
    parser.add_argument("--tag", default="", help="Empty means latest stable GitHub release")
    parser.add_argument("--timezone", default="Asia/Tashkent")
    parser.add_argument("--publish", action="store_true")
    parser.add_argument("--docspace-url", default=os.getenv("DOCSPACE_URL", "https://nct2.onlyoffice.com"))
    parser.add_argument("--folder-id", default=os.getenv("DOCSPACE_FOLDER_ID", ""))
    args = parser.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    if args.output.name != "DesktopEditors.xlsx":
        parser.error("Output filename must be DesktopEditors.xlsx to keep the remote document stable")
    try:
        portal, expected = None, None
        if args.publish:
            token = os.getenv("DOCSPACE_API_KEY")
            if not token:
                raise ReportError("Set DOCSPACE_API_KEY via GitHub Actions Secrets")
            portal = DocSpace(args.docspace_url, token, args.folder_id)
            expected = portal.find(args.output.name)
        release, assets = collect(session(os.getenv("GITHUB_TOKEN"), github=True), args.tag)
        now = datetime.now(ZoneInfo(args.timezone))
        write_report(args.output, release, assets, now)
        LOG.info("%s: %d assets, %s downloads; report: %s", release["tag_name"],
                 len(assets), sum(item["download_count"] for item in assets), args.output)
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
