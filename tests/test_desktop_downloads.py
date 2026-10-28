from datetime import datetime, timezone
import importlib.util
from pathlib import Path
import tempfile
import unittest
from unittest.mock import Mock, patch
from xml.etree import ElementTree as ET
from zipfile import ZipFile

SCRIPT = Path(__file__).resolve().parents[1] / "desktop_downloads.py"
SPEC = importlib.util.spec_from_file_location("desktop_downloads", SCRIPT)
report = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(report)
NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def asset(name, count, identifier=1):
    return {"id": identifier, "name": name, "download_count": count,
            "state": "uploaded"}


RELEASE = {"id": 10, "tag_name": "v10.0.0", "draft": False, "prerelease": False,
           "published_at": "2026-10-13T10:18:58Z"}


class ClassificationTests(unittest.TestCase):
    def test_groups_and_platforms(self):
        cases = {
            "DesktopEditors_x64.exe": ("Installations", "Windows"),
            "DesktopEditors_arm64.zip": ("Updates", "Windows"),
            "addon-win-x64.zip": ("Add-ons", "Windows"),
            "addon-linux-aarch64.zip": ("Add-ons", "Linux"),
            "addon-mac-v8.zip": ("Add-ons", "macOS"),
            "ONLYOFFICE-arm.dmg": ("Installations", "macOS"),
            "DesktopEditors-x86_64.AppImage": ("Installations", "Linux"),
            "onlyoffice-desktopeditors-help_all.deb": ("Installations", "Linux"),
            "onlyoffice-desktopeditors-help-any.tar.xz": ("Installations", "Linux"),
            "onlyoffice-desktopeditors.suse12.x86_64.rpm": ("Installations", "Linux"),
            "onlyoffice-documentserver.exe": ("Installations", "Windows"),
            "onlyoffice-documentserver_amd64.deb": ("Installations", "Linux"),
            "onlyoffice-documentserver_arm64.deb": ("Installations", "Linux"),
            "onlyoffice-documentserver.x86_64.rpm": ("Installations", "Linux"),
            "onlyoffice-documentserver.aarch64.rpm": ("Installations", "Linux"),
        }
        for name, expected in cases.items():
            with self.subTest(name=name):
                self.assertEqual(report.classify(name), expected)

    def test_unknown_asset_is_not_silently_dropped(self):
        with self.assertRaises(report.ReportError):
            report.classify("checksums.txt")


class CollectionTests(unittest.TestCase):
    @patch.object(report, "request_json")
    def test_documentserver_repository_is_used(self, get):
        get.side_effect = [RELEASE, [asset("onlyoffice-documentserver.exe", 5)]]
        report.collect(Mock(), repository="ONLYOFFICE/DocumentServer")
        for call in get.call_args_list:
            self.assertIn("/repos/ONLYOFFICE/DocumentServer/releases/", call.args[2])

    @patch.object(report, "request_json")
    def test_assets_are_paginated(self, get):
        first = [asset("DesktopEditors_x64.exe", n, n) for n in range(100)]
        get.side_effect = [RELEASE, first, [asset("addon-mac-arm.zip", 5, 101)]]
        release, assets = report.collect(Mock())
        self.assertEqual(len(assets), 101)
        self.assertEqual(release["tag_name"], "v10.0.0")
        self.assertEqual(get.call_args.kwargs["params"]["page"], 2)

    @patch.object(report, "request_json")
    def test_invalid_input_preserves_existing_report(self, get):
        for items in ([], [asset("DesktopEditors_x64.exe", -1)],
                      [asset("DesktopEditors_x64.exe", True)],
                      [asset("DesktopEditors_x64.exe", 1), asset("DesktopEditors_x86.exe", 2)]):
            get.side_effect = [RELEASE, items]
            with self.subTest(items=items), self.assertRaises(report.ReportError):
                report.collect(Mock())

    @patch.object(report, "request_json")
    def test_prerelease_rejected(self, get):
        get.return_value = dict(RELEASE, prerelease=True)
        with self.assertRaises(report.ReportError):
            report.collect(Mock(), "v10.1.0-beta")


class WorkbookTests(unittest.TestCase):
    def test_products_share_one_sheet_with_independent_totals(self):
        desktop = [asset("DesktopEditors_x64.exe", 100)]
        server = [asset("onlyoffice-documentserver.exe", 7),
                  asset("onlyoffice-documentserver_amd64.deb", 11, 2)]
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "ONLYOFFICE-Download-Statistics.xlsx"
            report.write_report(path, RELEASE, desktop, datetime.now(timezone.utc),
                                server=(dict(RELEASE, tag_name="v9.9.9"), server))
            with ZipFile(path) as archive:
                book = ET.fromstring(archive.read("xl/workbook.xml"))
                self.assertEqual(len(book.findall("s:sheets/s:sheet", NS)), 1)
                xml = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
                cells = {c.get("r"): c for c in xml.findall(".//s:c", NS)}
                for cell, count in {"E9": 100, "E10": 100, "K9": 7, "K12": 11,
                                    "K13": 18}.items():
                    self.assertEqual(int(cells[cell].find("s:v", NS).text), count)
                self.assertEqual(cells["K13"].find("s:f", NS).text, "SUM(K9,K12)")
                self.assertEqual(cells["K9"].find("s:f", NS).text, "SUM(K8:K8)")
                for name, cell in cells.items():
                    formula = cell.find("s:f", NS)
                    if name[0] in "GHIJK" and formula is not None:
                        self.assertNotRegex(formula.text, r"\b[A-E][0-9]+")
                strings = archive.read("xl/sharedStrings.xml").decode()
                self.assertIn("Document Server v9.9.9", strings)
                self.assertIn("Desktop Editors v10.0.0", strings)
                self.assertNotIn("Download summary", strings)
                self.assertNotIn("Total downloads", strings)
                self.assertEqual(len(xml.findall("s:hyperlinks/s:hyperlink", NS)), 0)

    def test_group_totals_without_summary_links_and_replacement(self):
        items = [asset("DesktopEditors_x64.exe", 1234),
                 asset("DesktopEditors_x86.exe", 567, 2),
                 asset("DesktopEditors_arm64.exe", 89, 3),
                 asset("DesktopEditors_x64.zip", 20, 4),
                 asset("addon-mac-arm.zip", 10, 5),
                 asset("onlyoffice-desktopeditors_amd64.deb", 30, 6)]
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "ONLYOFFICE-Download-Statistics.xlsx"
            report.write_report(path, RELEASE, items, datetime.now(timezone.utc))
            with ZipFile(path) as archive:
                book = ET.fromstring(archive.read("xl/workbook.xml"))
                self.assertEqual(len(book.findall("s:sheets/s:sheet", NS)), 1)
                xml = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
                strings = ET.fromstring(archive.read("xl/sharedStrings.xml"))
                strings = ["".join(node.itertext()) for node in strings]
                totals = {}
                for row in xml.findall("s:sheetData/s:row", NS):
                    cells = row.findall("s:c", NS)
                    if len(cells) >= 2 and cells[0].get("t") == "s":
                        label = strings[int(cells[0].find("s:v", NS).text)]
                        if label.startswith("Total"):
                            numeric = next(c for c in cells if c.get("r").startswith("E"))
                            totals.setdefault(label, []).append(int(numeric.find("s:v", NS).text))
                            self.assertIsNotNone(numeric.find("s:f", NS))
                self.assertEqual(totals["Total Windows"], [1890, 20])
                self.assertEqual(totals["Total — installations"], [1920])
                self.assertEqual(totals["Total — updates"], [20])
                self.assertEqual(totals["Total — add-ons"], [10])
                self.assertNotIn("Total downloads", totals)
                self.assertNotIn("Download summary", strings)
                self.assertNotIn("Downloads by file", strings)
                self.assertEqual(len(xml.findall("s:hyperlinks/s:hyperlink", NS)), 0)
                self.assertNotIn("Link", strings)
                self.assertNotIn("Download", strings)
            report.write_report(path, dict(RELEASE, tag_name="v10.1.0"), items[:1],
                                datetime.now(timezone.utc))
            with ZipFile(path) as archive:
                strings = archive.read("xl/sharedStrings.xml").decode()
                self.assertIn("v10.1.0", strings)
                self.assertNotIn("v10.0.0", strings)
                self.assertNotIn("DesktopEditors_x86.exe", strings)


class DocSpaceTests(unittest.TestCase):
    def portal(self):
        return report.DocSpace("https://nct2.onlyoffice.com", "test-secret", "42")

    def test_folder_listing_pagination(self):
        portal = self.portal()
        portal.call = Mock(side_effect=[
            {"files": [{"title": "Other.xlsx", "id": 1}], "folders": [], "total": 2},
            {"files": [{"title": "ONLYOFFICE-Download-Statistics.xlsx", "id": 5}], "folders": [], "total": 2},
            {"id": 5, "folderId": 42, "title": "ONLYOFFICE-Download-Statistics.xlsx", "version": 1}])
        self.assertEqual(portal.find("ONLYOFFICE-Download-Statistics.xlsx")["id"], 5)
        self.assertEqual(portal.call.call_args_list[1].kwargs["params"]["startIndex"], 1)

    def test_duplicate_names_rejected(self):
        portal = self.portal()
        portal.call = Mock(return_value={"files": [
            {"title": "ONLYOFFICE-Download-Statistics.xlsx", "id": 1},
            {"title": "ONLYOFFICE-Download-Statistics.xlsx", "id": 2}], "total": 2})
        with self.assertRaises(report.ReportError):
            portal.find("ONLYOFFICE-Download-Statistics.xlsx")

    def test_update_checks_id_and_version(self):
        original = {"id": 5, "version": 1}
        updated = {"id": 5, "version": 2}
        portal = self.portal()
        portal.find = Mock(side_effect=[original, updated])
        portal.call = Mock(return_value=[updated])
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "ONLYOFFICE-Download-Statistics.xlsx"
            path.write_bytes(b"test")
            self.assertEqual(portal.publish(path, original), updated)
        self.assertEqual(portal.call.call_args.kwargs["params"], {"createNewIfExist": "false"})

    def test_first_upload(self):
        portal = self.portal()
        stored = {"id": 5, "version": 1}
        portal.find = Mock(side_effect=[None, stored])
        portal.call = Mock(return_value=[stored])
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "ONLYOFFICE-Download-Statistics.xlsx"
            path.write_bytes(b"test")
            self.assertEqual(portal.publish(path, None), stored)

    def test_concurrent_change_rejected_before_upload(self):
        portal = self.portal()
        portal.find = Mock(return_value={"id": 5, "version": 2})
        portal.call = Mock()
        with self.assertRaises(report.ReportError):
            portal.publish(Path("ONLYOFFICE-Download-Statistics.xlsx"), {"id": 5, "version": 1})
        portal.call.assert_not_called()

    def test_duplicate_created_by_upload_fails(self):
        portal = self.portal()
        original = {"id": 5, "version": 1}
        portal.find = Mock(return_value=original)
        portal.call = Mock(return_value=[{"id": 6, "version": 1}])
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "ONLYOFFICE-Download-Statistics.xlsx"
            path.write_bytes(b"test")
            with self.assertRaises(report.ReportError):
                portal.publish(path, original)


if __name__ == "__main__":
    unittest.main()
