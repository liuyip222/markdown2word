"""Local-only web UI; needs no third-party packages to start."""
from __future__ import annotations

import argparse
from contextlib import contextmanager
import json
import os
from pathlib import Path
import secrets
import subprocess
import sys
import threading
import time
import tempfile
import shutil
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlsplit

import md_to_word as converter
from dependency_installer import install_pandoc

ROOT = Path(__file__).resolve().parent
STATE_FILE = ROOT / '.web-ui-state.json'
CREATE_NO_WINDOW = getattr(subprocess, 'CREATE_NO_WINDOW', 0)


def run_python(arguments, owner=None, **kwargs):
    env = dict(os.environ, PYTHONIOENCODING='utf-8', PYTHONUTF8='1')
    if owner is not None:
        timeout = kwargs.pop('timeout', None)
        with subprocess.Popen([sys.executable, *arguments], cwd=ROOT, env=env,
                              creationflags=CREATE_NO_WINDOW, encoding='utf-8', errors='replace',
                              stdout=subprocess.PIPE, stderr=subprocess.PIPE, **kwargs) as process:
            with owner.track_process(process):
                try:
                    stdout, stderr = process.communicate(timeout=timeout)
                except BaseException:
                    process.kill()
                    process.communicate()
                    raise
                return subprocess.CompletedProcess(process.args, process.returncode, stdout, stderr)
    return subprocess.run([sys.executable, *arguments], cwd=ROOT, env=env,
                          creationflags=CREATE_NO_WINDOW, encoding='utf-8',
                          errors='replace', capture_output=True, **kwargs)


def check_dependencies():
    # A fresh interpreter also detects packages installed while this server is running.
    code = ("import importlib.metadata, re; from docx import Document; "
            "v=importlib.metadata.version('python-docx'); "
            r"assert tuple(map(int,re.match(r'(\d+)\.(\d+)',v).groups())) >= (1,1); print(v)")
    result = run_python(['-c', code], timeout=30)
    python_ready = result.returncode == 0
    pandoc_error = ''
    try:
        pandoc = converter.find_pandoc(converter.load_config().get('pandoc'))
    except Exception as error:
        pandoc, pandoc_error = '', str(error)
    return {'python_ready': python_ready, 'python_version': result.stdout.strip() if python_ready else '',
            'pandoc_ready': bool(pandoc), 'pandoc_error': pandoc_error,
            'ready': python_ready and bool(pandoc)}


def choose_path(kind, initial, owner=None):
    # Tk lives in its own main-thread process, not in an HTTP worker thread.
    code = '''import json, sys, tkinter as tk
from tkinter import filedialog
root = tk.Tk(); root.withdraw(); root.attributes('-topmost', True)
try:
    if sys.argv[1] == 'files':
        value = list(filedialog.askopenfilenames(parent=root, title='选择 Markdown 文件（可多选）', initialdir=sys.argv[2], filetypes=[('Markdown 文件', '*.md *.markdown'), ('所有文件', '*.*')]))
    elif sys.argv[1] == 'template':
        value = filedialog.askopenfilename(parent=root, title='选择 Word 模板', initialdir=sys.argv[2], filetypes=[('Word 模板', '*.docx')])
    else:
        value = filedialog.askdirectory(parent=root, title='选择 Word 输出目录', initialdir=sys.argv[2], mustexist=True)
    print(json.dumps(value, ensure_ascii=False))
finally:
    root.destroy()
'''
    result = run_python(['-c', code, kind, initial], owner=owner)
    if result.returncode:
        raise RuntimeError('无法打开选择窗口，请确认 Python 已安装 tkinter。\n' + result.stderr.strip())
    return json.loads(result.stdout)


class Application:
    def __init__(self, files=()):
        self.lock = threading.RLock()
        self.process_lock = threading.Lock()
        self.stopping = threading.Event()
        self.active_processes = set()
        self.dialog_lock = threading.Lock()
        self.files = list(dict.fromkeys(str(Path(p).resolve()) for p in files))
        config = converter.load_config()
        self.output_dir = config.get('output_dir') or str(converter.default_output_dir())
        template = Path(config.get('template') or converter.DEFAULT_TEMPLATE).expanduser()
        self.template = str((template if template.is_absolute() else ROOT / template).resolve())
        try:
            saved = json.loads(STATE_FILE.read_text(encoding='utf-8'))
            if isinstance(saved.get('output_dir'), str) and saved['output_dir'].strip():
                self.output_dir = saved['output_dir']
        except (OSError, ValueError, AttributeError):
            pass
        self.dependencies = check_dependencies()
        self.job = self.empty_job()

    @staticmethod
    def empty_job():
        return {'kind': '', 'running': False, 'total': 0, 'done': 0, 'results': [], 'log': '', 'error': ''}

    def snapshot(self):
        with self.lock:
            return json.loads(json.dumps({'files': self.files, 'output_dir': self.output_dir,
                                         'template': self.template,
                                         'dependencies': self.dependencies, 'job': self.job}))

    def log(self, value):
        with self.lock:
            self.job['log'] = (self.job['log'] + value + '\n')[-100000:]

    def save_output(self, value):
        path = Path(value).expanduser()
        if not value.strip() or not path.is_absolute():
            raise ValueError('请选择输出目录，或填写完整的绝对路径。')
        if path.exists() and not path.is_dir():
            raise ValueError('输出路径不是文件夹。')
        self.output_dir = str(path.resolve())
        try:
            # Store an auto-resolve marker for the default, not one user's absolute Downloads path.
            saved_output = '' if path.resolve() == converter.default_output_dir().resolve() else self.output_dir
            STATE_FILE.write_text(json.dumps({'output_dir': saved_output}, ensure_ascii=False), encoding='utf-8')
        except OSError:
            pass  # Read-only tool folders can still convert to a writable output folder.

    def action(self, route, data):
        if route in ('/api/pick-files', '/api/pick-output', '/api/pick-template'):
            if not self.dialog_lock.acquire(blocking=False):
                raise ValueError('已有选择窗口打开，请先完成选择。')
            try:
                with self.lock:
                    if self.job['running']:
                        raise ValueError('请等待当前任务完成。')
                    if route.endswith('template'):
                        initial = str(ROOT / 'templates')
                    else:
                        initial = self.output_dir if route.endswith('output') else (str(Path(self.files[0]).parent) if self.files else str(Path.home()))
                if not Path(initial).is_dir():
                    initial = str(Path.home())
                kind = route.removeprefix('/api/pick-')
                value = choose_path(kind, initial, owner=self)
                with self.lock:
                    if value:
                        if route.endswith('files'):
                            self.files = list(dict.fromkeys(value))
                        elif route.endswith('template'):
                            self.set_template(value)
                        else:
                            self.save_output(value)
            finally:
                self.dialog_lock.release()
        elif route == '/api/remove-file':
            with self.lock:
                if self.job['running'] or self.dialog_lock.locked():
                    raise ValueError('请等待当前任务或选择窗口完成。')
                path = data.get('path')
                if not isinstance(path, str) or path not in self.files:
                    raise ValueError('该文件不在当前已选列表中。')
                self.files.remove(path)  # Remove the selection only; never touch the source file.
        elif route in ('/api/install', '/api/convert'):
            with self.lock:
                if self.job['running'] or self.dialog_lock.locked():
                    raise ValueError('请等待当前任务或选择窗口完成。')
                self.dependencies = check_dependencies()
                if route.endswith('install'):
                    if self.dependencies['ready']:
                        return self.snapshot()
                    self.job = dict(self.empty_job(), kind='install', running=True)
                    worker = self.install
                else:
                    if not self.dependencies['ready']:
                        raise ValueError('请先点击“安装依赖”，安装缺失的 Python 包和 Pandoc。')
                    if not self.files:
                        raise ValueError('请先选择 Markdown 文件。')
                    self.set_template(self.template)
                    for value in self.files:
                        path = Path(value)
                        if not path.is_file() or path.suffix.lower() not in ('.md', '.markdown'):
                            raise ValueError('文件不存在或不是 Markdown：' + value)
                    output = data.get('output_dir', self.output_dir)
                    if not isinstance(output, str):
                        raise ValueError('输出目录格式不正确。')
                    self.save_output(output)
                    self.job = dict(self.empty_job(), kind='convert', running=True, total=len(self.files))
                    worker = self.convert
                threading.Thread(target=self.run_job, args=(worker,), daemon=True).start()
        elif route == '/api/open-output':
            if not Path(self.output_dir).is_dir():
                raise ValueError('输出目录尚不存在，请先转换或选择已有目录。')
            os.startfile(self.output_dir)
        else:
            raise ValueError('未知操作。')
        return self.snapshot()

    def set_template(self, value):
        path = Path(value).expanduser().resolve()
        if not path.is_file() or path.suffix.lower() != '.docx':
            raise ValueError('请选择存在的 .docx Word 模板：' + str(path))
        self.template = str(path)

    @contextmanager
    def track_process(self, process):
        with self.process_lock:
            if self.stopping.is_set():
                process.kill()
                raise RuntimeError('工具正在退出。')
            self.active_processes.add(process)
        try:
            yield process
        finally:
            with self.process_lock:
                self.active_processes.discard(process)

    def stop(self):
        self.stopping.set()
        with self.process_lock:
            for process in tuple(self.active_processes):
                if process.poll() is None:
                    try:
                        process.kill()
                    except OSError:
                        pass

    def run_job(self, worker):
        try:
            worker()
        except Exception as error:
            with self.lock:
                self.job['error'] = str(error)
            self.log(str(error))
        finally:
            with self.lock:
                self.job['running'] = False

    def install(self):
        missing = dict(self.dependencies)
        try:
            if not missing['python_ready']:
                self.install_python_packages()
            if not missing['pandoc_ready']:
                install_pandoc(ROOT / 'bin', log=self.log, stop_event=self.stopping)
        finally:
            if not self.stopping.is_set():
                with self.lock:
                    self.dependencies = check_dependencies()
        if not self.dependencies['ready']:
            raise RuntimeError('依赖安装未完成，请检查日志后重试。')
        self.log('全部依赖已就绪，安装入口已隐藏。')

    def install_python_packages(self):
        self.log('正在安装 requirements.txt 中的依赖，请稍候…')
        env = dict(os.environ, PYTHONIOENCODING='utf-8', PYTHONUTF8='1', PIP_DISABLE_PIP_VERSION_CHECK='1')
        command = [sys.executable, '-m', 'pip', 'install', '-r', str(ROOT / 'requirements.txt')]
        with subprocess.Popen(command, cwd=ROOT, env=env, stdout=subprocess.PIPE,
                              stderr=subprocess.STDOUT, text=True, encoding='utf-8',
                              errors='replace', creationflags=CREATE_NO_WINDOW) as process, self.track_process(process):
            for line in process.stdout:
                self.log(line.rstrip())
            returncode = process.wait()
        if returncode:
            raise RuntimeError('依赖安装未完成，请查看日志并检查网络或 Python 环境，然后重试。')
        self.log('Python 依赖安装成功。')

    def convert(self):
        files, output, template = list(self.files), self.output_dir, self.template
        self.log('使用 Word 模板：' + template)
        for index, source in enumerate(files, 1):
            if self.stopping.is_set():
                break
            self.log(f'[{index}/{len(files)}] 正在转换 {Path(source).name}')
            try:
                result = run_python([str(ROOT / 'md_to_word.py'), source, '--output-dir', output,
                                     '--template', template], owner=self)
                if result.returncode:
                    raise RuntimeError(result.stderr.strip() or result.stdout.strip() or '转换失败')
                report = json.loads(result.stdout)
                entry = {'source': source, 'output': report['output'], 'success': True, 'warnings': report.get('warnings', [])}
                self.log('已生成：' + report['output'])
                for warning in entry['warnings']:
                    self.log('提示：' + warning)
            except Exception as error:
                entry = {'source': source, 'success': False, 'error': str(error)}
                self.log('失败：' + str(error))
            with self.lock:
                self.job['results'].append(entry)
                self.job['done'] = index
        success = sum(item['success'] for item in self.job['results'])
        self.log(f'转换完成：成功 {success} 个，失败 {len(files) - success} 个。')


def create_server(app, port=0):
    token = secrets.token_urlsafe(32)

    class Handler(BaseHTTPRequestHandler):
        def log_message(self, *args):
            pass

        def respond(self, status, value, content_type='application/json; charset=utf-8'):
            payload = json.dumps(value, ensure_ascii=False).encode('utf-8') if 'json' in content_type else value
            self.send_response(status)
            self.send_header('Content-Type', content_type)
            self.send_header('Content-Length', str(len(payload)))
            self.send_header('Cache-Control', 'no-store')
            self.send_header('X-Content-Type-Options', 'nosniff')
            self.send_header('Referrer-Policy', 'no-referrer')
            self.end_headers()
            self.wfile.write(payload)

        def authorized(self):
            host = f'127.0.0.1:{self.server.server_port}'
            return (self.headers.get('Host') == host and self.headers.get('X-Session-Token') == token
                    and self.headers.get('Origin', f'http://{host}') == f'http://{host}')

        def do_GET(self):
            path = urlsplit(self.path).path
            if self.headers.get('Host') != f'127.0.0.1:{self.server.server_port}':
                self.respond(403, {'error': '仅允许本机访问。'})
            elif path == '/':
                self.respond(200, (ROOT / 'web' / 'index.html').read_bytes(), 'text/html; charset=utf-8')
            elif path == '/api/state' and self.authorized():
                self.respond(200, app.snapshot())
            else:
                self.respond(404, {'error': '页面不存在或会话无效。'})

        def do_POST(self):
            if not self.authorized():
                self.respond(403, {'error': '会话无效，请从启动页面打开。'})
                return
            try:
                length = int(self.headers.get('Content-Length', '0'))
                if not 0 <= length <= 65536:
                    raise ValueError('请求过大。')
                data = json.loads(self.rfile.read(length) or b'{}')
                if not isinstance(data, dict):
                    raise ValueError('请求格式不正确。')
                route = urlsplit(self.path).path
                if route == '/api/quit':
                    with app.lock:
                        if app.job['running'] or app.dialog_lock.locked():
                            raise ValueError('请等待当前任务完成再退出。')
                    self.respond(200, {'closed': True})
                    threading.Thread(target=self.server.shutdown, daemon=True).start()
                else:
                    self.respond(200, app.action(route, data))
            except (ValueError, OSError, RuntimeError) as error:
                self.respond(400, {'error': str(error)})
            except Exception as error:
                self.respond(500, {'error': str(error)})

    server = ThreadingHTTPServer(('127.0.0.1', port), Handler)
    server.daemon_threads = True
    return server, f'http://127.0.0.1:{server.server_port}/#token={token}'


def find_app_browser():
    """Find an existing Chromium browser without changing its normal profile."""
    candidates = []
    if os.name == 'nt':
        import winreg
        for name in ('msedge.exe', 'chrome.exe'):
            for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
                for view in (winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY):
                    try:
                        with winreg.OpenKey(hive, 'SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\' + name,
                                            0, winreg.KEY_READ | view) as key:
                            candidates.append(winreg.QueryValue(key, None).strip('"'))
                    except OSError:
                        pass
        for base in (os.environ.get('ProgramFiles'), os.environ.get('ProgramFiles(x86)'), os.environ.get('LOCALAPPDATA')):
            if base:
                candidates.extend(str(Path(base) / path) for path in (
                    'Microsoft/Edge/Application/msedge.exe', 'Google/Chrome/Application/chrome.exe'))
    candidates.extend(filter(None, (shutil.which('msedge'), shutil.which('chrome'), shutil.which('chromium'))))
    return next((value for value in candidates if Path(value).is_file()), None)


class WindowsBrowserProcess:
    """Hold a real process handle, including browsers that relaunch their entry process."""
    def __init__(self, pid):
        import ctypes
        from ctypes import wintypes
        self.pid = pid
        self.kernel = ctypes.WinDLL('kernel32', use_last_error=True)
        self.kernel.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
        self.kernel.OpenProcess.restype = wintypes.HANDLE
        self.kernel.WaitForSingleObject.argtypes = [wintypes.HANDLE, wintypes.DWORD]
        self.kernel.WaitForSingleObject.restype = wintypes.DWORD
        self.kernel.CloseHandle.argtypes = [wintypes.HANDLE]
        self.kernel.TerminateProcess.argtypes = [wintypes.HANDLE, wintypes.UINT]
        self.handle = self.kernel.OpenProcess(0x100000 | 0x1000 | 0x1, False, pid)
        if not self.handle:
            raise OSError(ctypes.get_last_error(), '无法跟踪工具浏览器进程')

    def poll(self):
        return None if self.kernel.WaitForSingleObject(self.handle, 0) == 258 else 0

    def wait(self, timeout=None):
        milliseconds = 0xffffffff if timeout is None else min(int(timeout * 1000), 0xfffffffe)
        if self.kernel.WaitForSingleObject(self.handle, milliseconds) == 258:
            raise subprocess.TimeoutExpired('tool browser', timeout)
        return 0

    def kill(self):
        self.kernel.TerminateProcess(self.handle, 1)

    terminate = kill

    def release(self):
        if self.handle:
            self.kernel.CloseHandle(self.handle)
            self.handle = None

    def __del__(self):
        if getattr(self, 'handle', None):
            self.release()


def find_browser_process(profile, url):
    """Match only this launch's unique profile AND local URL; exclude browser helpers."""
    powershell = str(Path(os.environ.get('SystemRoot', 'C:/Windows')) /
                     'System32/WindowsPowerShell/v1.0/powershell.exe')
    command = '''$items=@(Get-CimInstance Win32_Process -Filter "Name='msedge.exe' OR Name='chrome.exe'" |
Where-Object {$_.CommandLine -and $_.CommandLine.Contains($env:MDWORD_BROWSER_PROFILE) -and
$_.CommandLine.Contains($env:MDWORD_BROWSER_URL) -and $_.CommandLine -notmatch '--type='} |
Sort-Object CreationDate -Descending | Select-Object -ExpandProperty ProcessId);
ConvertTo-Json -InputObject $items -Compress'''
    env = dict(os.environ, MDWORD_BROWSER_PROFILE=profile, MDWORD_BROWSER_URL=url)
    for _ in range(5):
        result = subprocess.run([powershell, '-NoProfile', '-NonInteractive', '-Command', command],
                                env=env, capture_output=True, text=True, creationflags=CREATE_NO_WINDOW,
                                timeout=15)
        if not result.returncode:
            for pid in json.loads(result.stdout.strip() or '[]'):
                try:
                    return WindowsBrowserProcess(int(pid))
                except OSError:
                    pass
        time.sleep(.2)
    raise RuntimeError('无法识别工具专属浏览器进程，请重新启动工具。')


class BrowserWindow:
    """Own only the tool's browser process tree; never close the user's normal browser."""
    def __init__(self):
        self.process = None
        self.profile = None

    def open(self, url):
        executable = find_app_browser()
        if not executable:
            webbrowser.open(url)
            print('未找到 Edge/Chrome，使用默认浏览器；若浏览器限制自动关闭，请手动关闭标签页。', flush=True)
            return
        self.profile = tempfile.TemporaryDirectory(prefix='mdword-browser-', ignore_cleanup_errors=True)
        try:
            self.process = subprocess.Popen(
                [executable, '--app=' + url, '--user-data-dir=' + self.profile.name,
                 '--no-first-run', '--no-default-browser-check', '--disable-background-mode',
                 '--window-size=960,900'], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                creationflags=CREATE_NO_WINDOW)
            if os.name == 'nt':
                launcher = self.process
                self.process = find_browser_process(self.profile.name, url)
                # Reap the short-lived compatibility launcher when applicable.
                launcher.poll()
        except Exception:
            self.close()
            raise

    def close(self):
        if self.process is not None and self.process.poll() is None:
            if os.name == 'nt':
                # PID is the still-running Popen child launched with a unique temporary profile.
                taskkill = str(Path(os.environ.get('SystemRoot', 'C:/Windows')) / 'System32/taskkill.exe')
                subprocess.run([taskkill, '/PID', str(self.process.pid), '/T', '/F'],
                               capture_output=True, creationflags=CREATE_NO_WINDOW, timeout=15)
            else:
                self.process.terminate()
            try:
                self.process.wait(timeout=10)
            except subprocess.TimeoutExpired:
                self.process.kill()
                self.process.wait(timeout=5)
        if self.profile is not None:
            self.profile.cleanup()
        if isinstance(self.process, WindowsBrowserProcess):
            self.process.release()
        self.process = self.profile = None

    def watch_closed(self, server, stop_event):
        """Closing the dedicated browser also ends the local service and launcher."""
        process = self.process
        if process is None:
            return None

        def watch():
            process.wait()
            if not stop_event.is_set():
                server.shutdown()

        thread = threading.Thread(target=watch, name='browser-close-watcher', daemon=True)
        thread.start()
        return thread


def main():
    parser = argparse.ArgumentParser(description='Markdown 转 Word 本地页面')
    parser.add_argument('files', nargs='*')
    parser.add_argument('--no-browser', action='store_true')
    parser.add_argument('--port', type=int, default=0)
    args = parser.parse_args()
    app = Application(args.files)
    server, url = create_server(app, args.port)
    print('本地转换页面：' + url, flush=True)
    print('保持此窗口打开；点击工具窗口右上角的 × 即可退出并关闭启动终端，也可按 Ctrl+C。', flush=True)
    browser = BrowserWindow()
    stop_event = threading.Event()
    try:
        if not args.no_browser:
            browser.open(url)
            browser.watch_closed(server, stop_event)
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        stop_event.set()
        app.stop()
        server.server_close()
        browser.close()


if __name__ == '__main__':
    main()
