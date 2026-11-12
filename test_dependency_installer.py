from io import BytesIO
import hashlib
from pathlib import Path
import subprocess
import tempfile
import threading
import unittest
from unittest.mock import patch
import urllib.error
import zipfile

import dependency_installer as installer


class Response(BytesIO):
    def __init__(self, content, url=installer.PANDOC_URL):
        super().__init__(content)
        self.url = url
        self.headers = {'Content-Length': str(len(content))}

    def geturl(self):
        return self.url


class InstallerTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        buffer = BytesIO()
        with zipfile.ZipFile(buffer, 'w') as archive:
            for name in installer.PORTABLE_FILES:
                archive.writestr(f'pandoc-{installer.PANDOC_VERSION}/{name}', b'fixture data')
        self.archive = buffer.getvalue()

    def test_download_checks_sha256(self):
        destination = self.root / 'download.zip'
        with patch.object(installer.urllib.request, 'urlopen', return_value=Response(self.archive)), \
             patch.object(installer, 'PANDOC_ZIP_SHA256', hashlib.sha256(self.archive).hexdigest()):
            installer.download_archive(destination, lambda text: None, threading.Event())
        self.assertEqual(self.archive, destination.read_bytes())

    def test_bad_checksum_is_not_installed(self):
        with patch.object(installer.urllib.request, 'urlopen', return_value=Response(self.archive)), \
             patch.object(installer, 'PANDOC_ZIP_SHA256', '0' * 64):
            with self.assertRaisesRegex(RuntimeError, 'SHA-256'):
                installer.install_pandoc(self.root / 'bin')
        self.assertFalse((self.root / 'bin/pandoc.exe').exists())

    def test_non_official_redirect_is_rejected(self):
        with patch.object(installer.urllib.request, 'urlopen', return_value=Response(self.archive, 'https://example.com/package.zip')):
            with self.assertRaisesRegex(RuntimeError, '非官方'):
                installer.download_archive(self.root / 'download.zip', lambda text: None, threading.Event())

    def test_install_publishes_executable_and_copyright_files(self):
        if installer.os.name != 'nt':
            self.skipTest('Windows only')
        with patch.object(installer.urllib.request, 'urlopen', return_value=Response(self.archive)), \
             patch.object(installer, 'PANDOC_ZIP_SHA256', hashlib.sha256(self.archive).hexdigest()), \
             patch.object(installer.subprocess, 'run', return_value=subprocess.CompletedProcess([], 0, f'pandoc.exe {installer.PANDOC_VERSION}\n', '')):
            executable = installer.install_pandoc(self.root / 'bin')
        self.assertTrue(executable.exists())
        for name in installer.PORTABLE_FILES:
            self.assertEqual(b'fixture data', (self.root / 'bin' / name).read_bytes())
        self.assertFalse(list((self.root / 'bin').glob('.pandoc-install-*')))

    def test_missing_zip_members_do_not_replace_existing_files(self):
        if installer.os.name != 'nt':
            self.skipTest('Windows only')
        buffer = BytesIO()
        with zipfile.ZipFile(buffer, 'w') as archive:
            archive.writestr('../pandoc.exe', b'wrong data')
        payload = buffer.getvalue()
        (self.root / 'bin').mkdir()
        (self.root / 'bin/pandoc.exe').write_bytes(b'original')
        with patch.object(installer.urllib.request, 'urlopen', return_value=Response(payload)), \
             patch.object(installer, 'PANDOC_ZIP_SHA256', hashlib.sha256(payload).hexdigest()):
            with self.assertRaisesRegex(RuntimeError, '内容异常'):
                installer.install_pandoc(self.root / 'bin')
        self.assertEqual(b'original', (self.root / 'bin/pandoc.exe').read_bytes())

    def test_network_errors_are_retried_and_reported(self):
        with patch.object(installer.urllib.request, 'urlopen', side_effect=urllib.error.URLError('offline')) as request:
            with self.assertRaisesRegex(RuntimeError, '检查网络'):
                installer.download_archive(self.root / 'download.zip', lambda text: None, threading.Event())
        self.assertEqual(2, request.call_count)

    def test_cancelled_install_does_not_download(self):
        stop = threading.Event()
        stop.set()
        with patch.object(installer.urllib.request, 'urlopen') as request:
            with self.assertRaisesRegex(RuntimeError, '取消'):
                installer.download_archive(self.root / 'download.zip', lambda text: None, stop)
        request.assert_not_called()


if __name__ == '__main__':
    unittest.main()
