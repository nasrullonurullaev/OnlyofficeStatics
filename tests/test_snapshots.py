from pathlib import Path
import tempfile
import unittest
from unittest.mock import MagicMock, Mock, patch

from test_desktop_downloads import report


class SnapshotTests(unittest.TestCase):
    def test_failed_history_download_does_not_publish(self):
        portal = Mock()
        portal.find.return_value = {'id': 5, 'version': 1}
        portal.download.side_effect = report.requests.ConnectionError('Download failed')
        with patch('sys.argv', ['report', '--publish']), \
                patch.dict('os.environ', {'DOCSPACE_API_KEY': 'test'}), \
                patch.object(report, 'DocSpace', return_value=portal), \
                patch.object(report, 'collect_tracked') as collect, self.assertLogs(report.LOG, level='ERROR'):
            self.assertEqual(report.main(), 1)
        collect.assert_not_called()
        portal.publish.assert_not_called()

    def test_signed_download_does_not_forward_the_api_key(self):
        portal = report.DocSpace('https://nct2.onlyoffice.com', 'secret', '42')
        portal.call = Mock(return_value='https://storage.example.com/signed-report')
        response = MagicMock()
        response.__enter__.return_value.status_code = 200
        response.__enter__.return_value.iter_content.return_value = [b'workbook bytes']
        with tempfile.TemporaryDirectory() as directory, patch.object(report.requests, 'get', return_value=response) as get:
            path = Path(directory) / 'history.xlsx'
            portal.download(path, {'id': 5})
            self.assertEqual(path.read_bytes(), b'workbook bytes')
        self.assertNotIn('headers', get.call_args.kwargs)
        self.assertNotIn('auth', get.call_args.kwargs)

    def test_portal_download_uses_authenticated_session(self):
        portal = report.DocSpace('https://nct2.onlyoffice.com', 'secret', '42')
        portal.call = Mock(return_value='https://nct2.onlyoffice.com/download?key=private')
        response = MagicMock()
        response.__enter__.return_value.status_code = 200
        response.__enter__.return_value.iter_content.return_value = [b'workbook bytes']
        portal.client.get = Mock(return_value=response)
        with tempfile.TemporaryDirectory() as directory, patch.object(report.requests, 'get') as external:
            path = Path(directory) / 'history.xlsx'
            portal.download(path, {'id': 5})
            self.assertEqual(path.read_bytes(), b'workbook bytes')
        portal.client.get.assert_called_once()
        self.assertEqual(portal.client.headers['Authorization'], 'Bearer secret')
        external.assert_not_called()

    def test_download_error_reports_status_without_signed_url(self):
        portal = report.DocSpace('https://nct2.onlyoffice.com', 'secret', '42')
        portal.call = Mock(return_value='https://storage.example.com/report?key=private')
        response = MagicMock()
        response.__enter__.return_value.status_code = 403
        with tempfile.TemporaryDirectory() as directory, patch.object(report.requests, 'get', return_value=response):
            path = Path(directory) / 'history.xlsx'
            path.write_bytes(b'previous report')
            with self.assertRaisesRegex(report.ReportError, 'HTTP 403') as failure:
                portal.download(path, {'id': 5})
            self.assertNotIn('private', str(failure.exception))
            self.assertEqual(path.read_bytes(), b'previous report')

    def test_redirect_to_storage_does_not_forward_authorization(self):
        portal = report.DocSpace('https://nct2.onlyoffice.com', 'secret', '42')
        portal.call = Mock(return_value='https://nct2.onlyoffice.com/download')
        redirect, success = MagicMock(), MagicMock()
        redirect.__enter__.return_value.status_code = 302
        redirect.__enter__.return_value.headers = {'Location': 'https://storage.example.com/report'}
        success.__enter__.return_value.status_code = 200
        success.__enter__.return_value.iter_content.return_value = [b'workbook bytes']
        portal.client.get = Mock(return_value=redirect)
        with tempfile.TemporaryDirectory() as directory, patch.object(report.requests, 'get', return_value=success) as external:
            portal.download(Path(directory) / 'history.xlsx', {'id': 5})
        self.assertNotIn('headers', external.call_args.kwargs)
        self.assertNotIn('auth', external.call_args.kwargs)
        self.assertFalse(external.call_args.kwargs['allow_redirects'])


if __name__ == '__main__':
    unittest.main()
