import hashlib
import json
from pathlib import Path
import tempfile
import unittest
import zipfile

import build_release


class ReleaseTests(unittest.TestCase):
    def test_release_includes_pandoc_but_excludes_personal_files(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for filename in build_release.ROOT_FILES:
                (root / filename).write_text('fixture', encoding='utf-8')
            for filename in build_release.BIN_FILES:
                path = root / 'bin' / filename
                path.parent.mkdir(exist_ok=True)
                path.write_bytes(b'fixture')
            for filename in ('web/index.html', 'lua/filter.lua', 'templates/template_标题不编号-列表第二行顶格.docx'):
                path = root / filename
                path.parent.mkdir(exist_ok=True)
                path.write_bytes(b'fixture')
            (root / '.web-ui-state.json').write_text('{"output_dir":"PRIVATE"}', encoding='utf-8')
            (root / 'config.json').write_text('{"pandoc":"PRIVATE"}', encoding='utf-8')
            (root / '.env').write_text('SECRET', encoding='utf-8')
            release, checksum = build_release.build_release(root)
            with zipfile.ZipFile(release) as archive:
                self.assertIn('Markdown-to-Word/bin/pandoc.exe', archive.namelist())
                self.assertIn('Markdown-to-Word/dependency_installer.py', archive.namelist())
                self.assertFalse(any('.env' in n or '.web-ui-state' in n or '/dist/' in n for n in archive.namelist()))
                config = json.loads(archive.read('Markdown-to-Word/config.json'))
                self.assertEqual('', config['output_dir'])
                self.assertEqual('', config['pandoc'])
                self.assertIsNone(archive.testzip())
            self.assertEqual(hashlib.sha256(release.read_bytes()).hexdigest(), checksum.read_text().split()[0])

    def test_missing_pandoc_is_reported(self):
        with tempfile.TemporaryDirectory() as directory:
            with self.assertRaises(FileNotFoundError):
                build_release.build_release(Path(directory))


if __name__ == '__main__':
    unittest.main()
