"""Build a portable GitHub Releases ZIP without personal state, caches or Git files."""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import tempfile
import zipfile

ROOT = Path(__file__).resolve().parent
ROOT_FILES = ('启动页面.cmd', 'web_ui.py', 'md_to_word.py', 'dependency_installer.py',
              'beautify_docx_tables.py', 'markdown-to-docx.lua', 'requirements.txt', 'README.md')
ASSET_DIRS = {'web': '*.html', 'templates': '*.docx', 'lua': '*.lua'}
BIN_FILES = ('pandoc.exe', 'COPYING.rtf', 'COPYRIGHT.txt')
FOLDER = 'Markdown-to-Word'


def build_release(root=ROOT, output=None):
    root = Path(root).resolve()
    output = Path(output or root / 'dist' / 'Markdown-to-Word-Windows-full.zip').resolve()
    files = [root / name for name in ROOT_FILES]
    files += [root / 'bin' / name for name in BIN_FILES]
    for directory, pattern in ASSET_DIRS.items():
        assets = sorted((root / directory).glob(pattern))
        if not assets:
            raise FileNotFoundError('缺少发布资源目录：' + directory)
        files.extend(assets)
    for path in files:
        if not path.is_file() or not path.resolve().is_relative_to(root) or path.is_symlink():
            raise FileNotFoundError('缺少发布文件或资源路径无效：' + str(path))
    # Release defaults are independent of the packager's chosen folders and external executables.
    config = {'output_dir': '', 'template': 'templates/template_标题不编号-列表第二行顶格.docx',
              'pandoc': '', 'pipeline': 'auto', 'beautify_tables': False,
              'syntax_highlighting': False, 'table_style': 'Table'}
    if not (root / config['template']).is_file():
        raise FileNotFoundError('缺少默认 Word 模板。')
    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='.release-build-', dir=output.parent) as directory:
        temporary = Path(directory) / 'release.zip'
        with zipfile.ZipFile(temporary, 'w', compression=zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
            for path in files:
                archive.write(path, FOLDER + '/' + path.relative_to(root).as_posix())
            archive.writestr(FOLDER + '/config.json', json.dumps(config, ensure_ascii=False, indent=2))
        with zipfile.ZipFile(temporary) as archive:
            if archive.testzip() is not None:
                raise RuntimeError('发布 ZIP 校验失败。')
        temporary.replace(output)
    checksum = hashlib.sha256()
    with output.open('rb') as stream:
        while chunk := stream.read(1024 * 1024):
            checksum.update(chunk)
    checksum_file = output.with_suffix(output.suffix + '.sha256')
    checksum_file.write_text(checksum.hexdigest() + '  ' + output.name + '\n', encoding='utf-8')
    return output, checksum_file


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='生成包含 Pandoc 的 Windows 完整版发布 ZIP')
    parser.add_argument('--output', help='指定输出 ZIP 路径，默认 dist/Markdown-to-Word-Windows-full.zip')
    args = parser.parse_args()
    release, checksum = build_release(output=args.output)
    print('发布包：' + str(release))
    print('SHA-256：' + str(checksum))
