"""Download the pinned official Windows Pandoc portable package, using only stdlib."""
from __future__ import annotations

import hashlib
import os
from pathlib import Path
import platform
import shutil
import subprocess
import tempfile
import threading
import urllib.error
import urllib.parse
import urllib.request
import zipfile

PANDOC_VERSION = '3.6.4'
PANDOC_URL = ('https://github.com/jgm/pandoc/releases/download/' + PANDOC_VERSION +
              '/pandoc-' + PANDOC_VERSION + '-windows-x86_64.zip')
# SHA-256 of the official 3.6.4 Windows x86_64 ZIP, verified during release preparation.
PANDOC_ZIP_SHA256 = 'a9e5feb3d56d2fb0e3e765d1c33b8ee6b72e6963d7de31504edeec8cd1be34b1'
MAX_DOWNLOAD_BYTES = 100 * 1024 * 1024
PORTABLE_FILES = ('COPYING.rtf', 'COPYRIGHT.txt', 'pandoc.exe')
CREATE_NO_WINDOW = getattr(subprocess, 'CREATE_NO_WINDOW', 0)


def cancelled(event):
    if event.is_set():
        raise RuntimeError('工具已关闭，依赖安装已取消。')


def download_archive(destination, log, event):
    for attempt in range(1, 3):
        cancelled(event)
        try:
            request = urllib.request.Request(PANDOC_URL, headers={'User-Agent': 'Markdown-to-Word/1.0'})
            with urllib.request.urlopen(request, timeout=30) as response, destination.open('wb') as stream:
                location = urllib.parse.urlsplit(response.geturl())
                if location.scheme != 'https' or not (location.hostname == 'github.com' or
                        (location.hostname or '').endswith('.githubusercontent.com')):
                    raise RuntimeError('Pandoc 下载被重定向到非官方地址。')
                total = int(response.headers.get('Content-Length') or 0)
                if total > MAX_DOWNLOAD_BYTES:
                    raise RuntimeError('Pandoc 下载包大小超出预期。')
                received, last_percent = 0, -1
                digest = hashlib.sha256()
                while True:
                    cancelled(event)
                    chunk = response.read(1024 * 1024)
                    if not chunk:
                        break
                    received += len(chunk)
                    if received > MAX_DOWNLOAD_BYTES:
                        raise RuntimeError('Pandoc 下载包大小超出预期。')
                    stream.write(chunk)
                    digest.update(chunk)
                    percent = int(received * 100 / total) if total else 0
                    if percent // 10 != last_percent // 10 or not total:
                        log(f'Pandoc 下载：{received / 1024 / 1024:.1f} MB' + (f' / {total / 1024 / 1024:.1f} MB（{percent}%）' if total else ''))
                        last_percent = percent
                if not received or (total and received != total):
                    raise RuntimeError('Pandoc 下载未完成，请重试。')
                if digest.hexdigest() != PANDOC_ZIP_SHA256:
                    raise RuntimeError('Pandoc 下载包 SHA-256 校验失败，未安装，请重试或手动安装官方 Pandoc。')
            return
        except (urllib.error.URLError, TimeoutError, ConnectionError) as error:
            if attempt == 2:
                raise RuntimeError('无法下载官方 Pandoc，请检查网络后重试。\n' + str(error)) from error
            log('Pandoc 网络下载失败，正在重试…')


def install_pandoc(bin_dir, log=lambda text: None, stop_event=None):
    if os.name != 'nt' or platform.machine().lower() not in ('amd64', 'x86_64', 'arm64', 'aarch64'):
        raise RuntimeError('自动下载支持 Windows 64 位系统；其他系统请手动安装 Pandoc。')
    event = stop_event or threading.Event()
    bin_dir = Path(bin_dir).resolve()
    log(f'正在下载官方 Pandoc {PANDOC_VERSION} Windows 便携版…')
    with tempfile.TemporaryDirectory(prefix='mdword-pandoc-', ignore_cleanup_errors=True) as directory:
        temporary = Path(directory)
        archive_path = temporary / 'pandoc.zip'
        download_archive(archive_path, log, event)
        cancelled(event)
        log('下载完成，SHA-256 校验通过，正在准备 Pandoc…')
        staged = temporary / 'staged'
        staged.mkdir()
        with zipfile.ZipFile(archive_path) as archive:
            for filename in PORTABLE_FILES:
                expected = f'pandoc-{PANDOC_VERSION}/{filename}'
                matches = [info for info in archive.infolist() if info.filename == expected]
                if len(matches) != 1 or matches[0].file_size > (400 * 1024 * 1024 if filename.endswith('.exe') else 1024 * 1024):
                    raise RuntimeError('Pandoc 便携包内容异常，未安装。')
                with archive.open(matches[0]) as source, (staged / filename).open('wb') as output:
                    while chunk := source.read(1024 * 1024):
                        cancelled(event)
                        output.write(chunk)
        executable = staged / 'pandoc.exe'
        result = subprocess.run([str(executable), '--version'], capture_output=True, text=True,
                                encoding='utf-8', errors='replace', timeout=30,
                                creationflags=CREATE_NO_WINDOW)
        first_line = result.stdout.splitlines()[0].strip() if result.stdout else ''
        if result.returncode or first_line not in ('pandoc ' + PANDOC_VERSION, 'pandoc.exe ' + PANDOC_VERSION):
            raise RuntimeError('Pandoc 程序校验失败，未安装。')
        cancelled(event)
        bin_dir.mkdir(parents=True, exist_ok=True)
        # Stage on the same filesystem and publish the executable last.
        with tempfile.TemporaryDirectory(prefix='.pandoc-install-', dir=bin_dir,
                                         ignore_cleanup_errors=True) as final_directory:
            final = Path(final_directory)
            for filename in PORTABLE_FILES:
                cancelled(event)
                shutil.copyfile(staged / filename, final / filename)
            cancelled(event)
            for filename in PORTABLE_FILES:
                os.replace(final / filename, bin_dir / filename)
    log('Pandoc 安装成功：' + str(bin_dir / 'pandoc.exe'))
    return bin_dir / 'pandoc.exe'
