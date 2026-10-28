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
            "state": "uploaded", "browser_download_url": f"https://github.com/example/{name}"}


RELEASE = {"id": 10, "tag_name": "v10.0.0", "draft": False, "prerelease": False,
           "published_at": "2026-10-13T10:18:58Z",
           "html_url": "https://github.com/ONLYOFFICE/DesktopEditors/releases/tag/v10.0.0"}


class ClassificationTests(unittest.TestCase):
    def test_groups_and_platforms(self):
        cases = {
            "DesktopEditors_x64.exe": ("Инсталляции", "Windows"),
            "DesktopEditors_arm64.zip": ("Обновления", "Windows"),
            "addon-win-x64.zip": ("Аддоны", "Windows"),
            "addon-linux-aarch64.zip": ("Аддоны", "Linux"),
            "addon-mac-v8.zip": ("Аддоны", "macOS"),
            "ONLYOFFICE-arm.dmg": ("Инсталляции", "macOS"),
            "DesktopEditors-x86_64.AppImage": ("Инсталляции", "Linux"),
            "onlyoffice-desktopeditors-help_all.deb": ("Инсталляции", "Linux"),
            "onlyoffice-desktopeditors-help-any.tar.xz": ("Инсталляции", "Linux"),
            "onlyoffice-desktopeditors.suse12.x86_64.rpm": ("Инсталляции", "Linux"),
        }
        for name, expected in cases.items():
            with self.subTest(name=name):
                self.assertEqual(report.classify(name), expected)

    def test_unknown_asset_is_not_silently_dropped(self):
        with self.assertRaises(report.ReportError):
            report.classify("checksums.txt")


class CollectionTests(unittest.TestCase):
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
    def test_totals_summary_no_links_and_replacement(self):
        items = [asset("DesktopEditors_x64.exe", 1234),
                 asset("DesktopEditors_x86.exe", 567, 2),
                 asset("DesktopEditors_arm64.exe", 89, 3),
                 asset("DesktopEditors_x64.zip", 20, 4),
                 asset("addon-mac-arm.zip", 10, 5),
                 asset("onlyoffice-desktopeditors_amd64.deb", 30, 6)]
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "DesktopEditors.xlsx"
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
                        if label.startswith("Всего"):
                            numeric = next(c for c in cells if c.get("r").startswith("E"))
                            totals.setdefault(label, []).append(int(numeric.find("s:v", NS).text))
                            self.assertIsNotNone(numeric.find("s:f", NS))
                self.assertEqual(totals["Всего Windows"], [1890, 20])
                self.assertEqual(totals["Всего — инсталляции"], [1920])
                self.assertEqual(totals["Всего скачиваний"], [1950])
                self.assertEqual(len(xml.findall("s:hyperlinks/s:hyperlink", NS)), 0)
                self.assertNotIn("Ссылка", strings)
                self.assertNotIn("Скачать", strings)
                values = {c.get("r"): c.find("s:v", NS).text for c in xml.findall(".//s:c", NS) if c.find("s:f", NS) is not None}
                self.assertEqual([int(values[cell]) for cell in ("B8", "C8", "D8", "E8")], [1890, 20, 0, 1910])
                self.assertEqual([int(values[cell]) for cell in ("B11", "C11", "D11", "E11")], [1920, 20, 10, 1950])
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
            {"files": [{"title": "DesktopEditors.xlsx", "id": 5}], "folders": [], "total": 2},
            {"id": 5, "folderId": 42, "title": "DesktopEditors.xlsx", "version": 1}])
        self.assertEqual(portal.find("DesktopEditors.xlsx")["id"], 5)
        self.assertEqual(portal.call.call_args_list[1].kwargs["params"]["startIndex"], 1)

    def test_duplicate_names_rejected(self):
        portal = self.portal()
        portal.call = Mock(return_value={"files": [
            {"title": "DesktopEditors.xlsx", "id": 1},
            {"title": "DesktopEditors.xlsx", "id": 2}], "total": 2})
        with self.assertRaises(report.ReportError):
            portal.find("DesktopEditors.xlsx")

    def test_update_checks_id_and_version(self):
        original = {"id": 5, "version": 1}
        updated = {"id": 5, "version": 2}
        portal = self.portal()
        portal.find = Mock(side_effect=[original, updated])
        portal.call = Mock(return_value=[updated])
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "DesktopEditors.xlsx"
            path.write_bytes(b"test")
            self.assertEqual(portal.publish(path, original), updated)
        self.assertEqual(portal.call.call_args.kwargs["params"], {"createNewIfExist": "false"})

    def test_first_upload(self):
        portal = self.portal()
        stored = {"id": 5, "version": 1}
        portal.find = Mock(side_effect=[None, stored])
        portal.call = Mock(return_value=[stored])
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "DesktopEditors.xlsx"
            path.write_bytes(b"test")
            self.assertEqual(portal.publish(path, None), stored)

    def test_concurrent_change_rejected_before_upload(self):
        portal = self.portal()
        portal.find = Mock(return_value={"id": 5, "version": 2})
        portal.call = Mock()
        with self.assertRaises(report.ReportError):
            portal.publish(Path("DesktopEditors.xlsx"), {"id": 5, "version": 1})
        portal.call.assert_not_called()

    def test_duplicate_created_by_upload_fails(self):
        portal = self.portal()
        original = {"id": 5, "version": 1}
        portal.find = Mock(return_value=original)
        portal.call = Mock(return_value=[{"id": 6, "version": 1}])
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "DesktopEditors.xlsx"
            path.write_bytes(b"test")
            with self.assertRaises(report.ReportError):
                portal.publish(path, original)


if __name__ == "__main__":
    unittest.main()
