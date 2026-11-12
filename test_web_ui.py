import json
from pathlib import Path
import tempfile
import threading
import time
import unittest
from unittest.mock import patch
from urllib.error import HTTPError
from urllib.parse import urlsplit, parse_qs
from urllib.request import Request, urlopen

import web_ui

READY = {'python_ready': True, 'python_version': '1.2.0', 'pandoc_ready': True,
         'pandoc_error': '', 'ready': True}


class WebUITests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.state_patch = patch.object(web_ui, 'STATE_FILE', Path(self.temp.name) / 'state.json')
        self.state_patch.start()
        self.addCleanup(self.state_patch.stop)

    def app(self, ready=READY):
        with patch.object(web_ui, 'check_dependencies', return_value=ready):
            return web_ui.Application()

    def wait_job(self, app):
        deadline = time.monotonic() + 40
        while app.snapshot()['job']['running'] and time.monotonic() < deadline:
            time.sleep(.05)
        self.assertFalse(app.snapshot()['job']['running'])

    def test_dependencies_are_actual_environment(self):
        self.assertTrue(web_ui.check_dependencies()['ready'])

    def test_windows_launcher_has_crlf_line_endings(self):
        script = (web_ui.ROOT / '启动页面.cmd').read_bytes()
        self.assertTrue(script.startswith(b'@echo off\r\n'))
        self.assertNotIn(b'\n', script.replace(b'\r\n', b''))
        self.assertNotIn('安装依赖.cmd'.encode('utf-8'), script)
        self.assertNotIn('启动转换.cmd'.encode('utf-8'), script)

    def test_dependency_check_handles_missing_package(self):
        from subprocess import CompletedProcess
        with patch.object(web_ui, 'run_python', return_value=CompletedProcess([], 1, '', 'No module named docx')):
            result = web_ui.check_dependencies()
        self.assertFalse(result['python_ready'])
        self.assertFalse(result['ready'])

    def test_installed_environment_skips_install(self):
        app = self.app()
        with patch.object(web_ui, 'check_dependencies', return_value=READY), patch.object(app, 'install') as install:
            app.action('/api/install', {})
        install.assert_not_called()
        self.assertFalse(app.job['running'])

    def test_missing_dependencies_block_conversion(self):
        missing = dict(READY, python_ready=False, ready=False)
        app = self.app(missing)
        with patch.object(web_ui, 'check_dependencies', return_value=missing):
            with self.assertRaisesRegex(ValueError, '依赖'):
                app.action('/api/convert', {})

    def test_output_remembered_and_config_unchanged(self):
        config_before = (web_ui.ROOT / 'config.json').read_bytes()
        app = self.app()
        destination = str(Path(self.temp.name) / '新 输出')
        app.save_output(destination)
        self.assertEqual(destination, self.app().output_dir)
        self.assertEqual(config_before, (web_ui.ROOT / 'config.json').read_bytes())
        with self.assertRaises(ValueError):
            app.save_output('relative/folder')

    def test_web_default_uses_system_downloads_folder(self):
        downloads = Path(self.temp.name) / '系统下载'
        with patch.object(web_ui.converter, 'default_output_dir', return_value=downloads):
            self.assertEqual(str(downloads), self.app().output_dir)

    def test_default_output_marker_is_portable_between_users(self):
        first = Path(self.temp.name) / '用户甲下载'
        second = Path(self.temp.name) / '用户乙下载'
        with patch.object(web_ui.converter, 'default_output_dir', return_value=first):
            app = self.app()
            app.save_output(str(first))
            self.assertEqual('', json.loads(web_ui.STATE_FILE.read_text(encoding='utf-8'))['output_dir'])
        with patch.object(web_ui.converter, 'default_output_dir', return_value=second):
            self.assertEqual(str(second), self.app().output_dir)

    def test_relocated_windows_downloads_folder_is_respected(self):
        if web_ui.os.name != 'nt':
            self.skipTest('Windows only')
        relocated = Path(self.temp.name) / '其他磁盘下载'
        with patch.object(web_ui.converter, 'windows_downloads_dir', return_value=relocated):
            self.assertEqual(relocated, web_ui.converter.default_output_dir())

    def test_converter_uses_system_downloads_when_no_output_is_specified(self):
        source = Path(self.temp.name) / '默认输出验证.md'
        source.write_text('# 默认输出验证\n\n正文。\n', encoding='utf-8')
        downloads = Path(self.temp.name) / '系统下载'
        args = web_ui.converter.parser().parse_args([str(source)])
        config = web_ui.converter.load_config()
        config['output_dir'] = ''
        with patch.object(web_ui.converter, 'default_output_dir', return_value=downloads):
            report = web_ui.converter.convert(source, args, config)
        self.assertEqual(downloads / '默认输出验证.docx', Path(report['output']))
        self.assertTrue(Path(report['output']).is_file())

    def test_cancel_selection_keeps_previous_values(self):
        app = self.app()
        app.files = ['existing.md']
        output = app.output_dir
        with patch.object(web_ui, 'choose_path', return_value=[]):
            app.action('/api/pick-files', {})
        with patch.object(web_ui, 'choose_path', return_value=''):
            app.action('/api/pick-output', {})
        self.assertEqual(['existing.md'], app.files)
        self.assertEqual(output, app.output_dir)

    def test_install_success_refreshes_environment(self):
        from io import StringIO
        from unittest.mock import MagicMock
        app = self.app(dict(READY, python_ready=False, ready=False))
        process = MagicMock()
        process.__enter__.return_value = process
        process.stdout = StringIO('Successfully installed python-docx\n')
        process.wait.return_value = 0
        with patch.object(web_ui.subprocess, 'Popen', return_value=process), patch.object(web_ui, 'check_dependencies', return_value=READY):
            app.install()
        self.assertTrue(app.dependencies['python_ready'])
        self.assertIn('安装成功', app.job['log'])

    def test_install_failure_allows_retry(self):
        app = self.app()
        def fail():
            raise RuntimeError('network unavailable')
        app.job['running'] = True
        app.run_job(fail)
        self.assertFalse(app.job['running'])
        self.assertEqual('network unavailable', app.job['error'])

    def test_missing_pandoc_is_installed_without_reinstalling_python(self):
        missing = dict(READY, pandoc_ready=False, ready=False)
        app = self.app(missing)
        with patch.object(app, 'install_python_packages') as python_install, \
             patch.object(web_ui, 'install_pandoc') as pandoc_install, \
             patch.object(web_ui, 'check_dependencies', return_value=READY):
            app.install()
        python_install.assert_not_called()
        pandoc_install.assert_called_once_with(web_ui.ROOT / 'bin', log=app.log, stop_event=app.stopping)
        self.assertTrue(app.dependencies['ready'])

    def test_both_missing_dependencies_are_installed(self):
        app = self.app(dict(READY, python_ready=False, pandoc_ready=False, ready=False))
        with patch.object(app, 'install_python_packages') as python_install, \
             patch.object(web_ui, 'install_pandoc') as pandoc_install, \
             patch.object(web_ui, 'check_dependencies', return_value=READY):
            app.install()
        python_install.assert_called_once()
        pandoc_install.assert_called_once()
        self.assertTrue(app.dependencies['ready'])

    def test_partial_install_failure_preserves_completed_dependency_state(self):
        missing = dict(READY, python_ready=False, pandoc_ready=False, ready=False)
        partial = dict(READY, pandoc_ready=False, ready=False)
        app = self.app(missing)
        with patch.object(app, 'install_python_packages'), \
             patch.object(web_ui, 'install_pandoc', side_effect=RuntimeError('download failed')), \
             patch.object(web_ui, 'check_dependencies', return_value=partial):
            with self.assertRaisesRegex(RuntimeError, 'download failed'):
                app.install()
        self.assertTrue(app.dependencies['python_ready'])
        self.assertFalse(app.dependencies['ready'])

    def test_install_action_is_enabled_when_only_pandoc_is_missing(self):
        missing = dict(READY, pandoc_ready=False, ready=False)
        app = self.app(missing)
        with patch.object(web_ui, 'check_dependencies', return_value=missing), \
             patch.object(app, 'install') as install:
            app.action('/api/install', {})
            self.wait_job(app)
        install.assert_called_once()
        page = (web_ui.ROOT / 'web/index.html').read_text(encoding='utf-8')
        self.assertIn("$('install-card').hidden=dep.ready", page)

    def test_http_authentication_and_page(self):
        app = self.app()
        server, url = web_ui.create_server(app)
        threading.Thread(target=server.serve_forever, daemon=True).start()
        self.addCleanup(server.server_close)
        self.addCleanup(server.shutdown)
        base = url.split('#')[0].rstrip('/')
        token = parse_qs(urlsplit(url).fragment)['token'][0]
        with urlopen(base) as response:
            page = response.read()
            self.assertIn('选择 Word 模板'.encode(), page)
            self.assertNotIn('把 Markdown'.encode(), page)
            self.assertIn('<span class="step">03</span>'.encode(), page)
        request = Request(base + '/api/state', headers={'X-Session-Token': token})
        with urlopen(request) as response:
            self.assertTrue(json.load(response)['dependencies']['ready'])
        for headers in ({}, {'X-Session-Token': token, 'Origin': 'https://external.example'},
                        {'X-Session-Token': token, 'Host': 'external.example'}):
            with self.assertRaises(HTTPError) as error:
                urlopen(Request(base + '/api/install', data=b'{}', headers=headers))
            self.assertEqual(403, error.exception.code)
            error.exception.close()
        with self.assertRaises(HTTPError) as error:
            urlopen(Request(base + '/api/install', data=b'[]', headers={'X-Session-Token': token}))
        self.assertEqual(400, error.exception.code)
        error.exception.close()

    def test_real_conversion_batch_and_no_overwrite(self):
        app = self.app()
        directory = Path(self.temp.name)
        good = directory / '含 空格的示例.md'
        good.write_text('# 示例\n\n正文。\n\n| 标题 | 值 |\n| --- | --- |\n| 项目 | 内容 |\n', encoding='utf-8')
        bad = directory / '失败.md'
        bad.write_text('```text\nunclosed\n', encoding='utf-8')
        output = directory / '输出 目录'
        app.files = [str(good), str(bad)]
        with patch.object(web_ui, 'check_dependencies', return_value=READY):
            app.action('/api/convert', {'output_dir': str(output)})
        self.wait_job(app)
        self.assertEqual([True, False], [r['success'] for r in app.job['results']])
        self.assertTrue((output / '含 空格的示例.docx').is_file())
        self.assertTrue((output / '含 空格的示例.conversion-report.json').is_file())
        app.files = [str(good)]
        with patch.object(web_ui, 'check_dependencies', return_value=READY):
            app.action('/api/convert', {'output_dir': str(output)})
        self.wait_job(app)
        self.assertTrue((output / '含 空格的示例 (1).docx').is_file())

    def test_default_template_matches_configuration(self):
        app = self.app()
        configured = Path(web_ui.converter.load_config().get('template') or web_ui.converter.DEFAULT_TEMPLATE)
        configured = configured if configured.is_absolute() else web_ui.ROOT / configured
        self.assertEqual(str(configured.resolve()), app.snapshot()['template'])

    def test_template_picker_always_starts_in_templates(self):
        app = self.app()
        chosen = web_ui.ROOT / 'templates' / 'template_标题编号-列表第二行缩进.docx'
        with patch.object(web_ui, 'choose_path', return_value=str(chosen)) as choose:
            app.action('/api/pick-template', {})
        choose.assert_called_once_with('template', str(web_ui.ROOT / 'templates'), owner=app)
        self.assertEqual(str(chosen.resolve()), app.template)
        with patch.object(web_ui, 'choose_path', return_value=''):
            app.action('/api/pick-template', {})
        self.assertEqual(str(chosen.resolve()), app.template)

    def test_invalid_template_preserves_selection(self):
        app = self.app()
        original = app.template
        with self.assertRaises(ValueError):
            app.set_template(str(Path(self.temp.name) / 'missing.docx'))
        self.assertEqual(original, app.template)

    def test_selected_template_is_used_for_conversion(self):
        app = self.app()
        source = Path(self.temp.name) / '模板验证.md'
        source.write_text('# 模板验证\n\n正文内容。\n', encoding='utf-8')
        output = Path(self.temp.name) / 'output'
        chosen = web_ui.ROOT / 'templates' / 'template_标题编号-列表第二行缩进.docx'
        app.set_template(str(chosen))
        app.files = [str(source)]
        with patch.object(web_ui, 'check_dependencies', return_value=READY):
            app.action('/api/convert', {'output_dir': str(output)})
        self.wait_job(app)
        self.assertTrue(app.job['results'][0]['success'])
        report = json.loads((output / '模板验证.conversion-report.json').read_text(encoding='utf-8'))
        self.assertEqual(str(chosen.resolve()), report['template'])

    def test_browser_uses_separate_profile_and_closes_only_own_process(self):
        from unittest.mock import MagicMock
        browser = web_ui.BrowserWindow()
        process = MagicMock()
        process.poll.return_value = None
        process.pid = 12345
        profile = MagicMock()
        profile.name = str(Path(self.temp.name) / 'isolated-browser')
        with patch.object(web_ui, 'find_app_browser', return_value='C:/browser/msedge.exe'), \
             patch.object(web_ui.tempfile, 'TemporaryDirectory', return_value=profile), \
             patch.object(web_ui.subprocess, 'Popen', return_value=process) as popen, \
             patch.object(web_ui, 'find_browser_process', return_value=process), \
             patch.object(web_ui.subprocess, 'run') as run:
            browser.open('http://127.0.0.1:1234/#token=test')
            args = popen.call_args.args[0]
            self.assertIn('--user-data-dir=' + profile.name, args)
            self.assertIn('--app=http://127.0.0.1:1234/#token=test', args)
            browser.close()
        if web_ui.os.name == 'nt':
            self.assertEqual(['/PID', '12345', '/T', '/F'], run.call_args.args[0][1:])
        else:
            process.terminate.assert_called_once()
        profile.cleanup.assert_called_once()

    def test_browser_does_not_kill_exited_process(self):
        from unittest.mock import MagicMock
        browser = web_ui.BrowserWindow()
        browser.process = MagicMock()
        browser.process.poll.return_value = 0
        with patch.object(web_ui.subprocess, 'run') as run:
            browser.close()
        run.assert_not_called()

    def test_quit_stops_http_server(self):
        server, url = web_ui.create_server(self.app())
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        self.addCleanup(server.server_close)
        self.addCleanup(server.shutdown)
        base = url.split('#')[0].rstrip('/')
        token = parse_qs(urlsplit(url).fragment)['token'][0]
        with urlopen(Request(base + '/api/quit', data=b'{}', headers={'X-Session-Token': token})) as response:
            self.assertTrue(json.load(response)['closed'])
        thread.join(timeout=3)
        self.assertFalse(thread.is_alive())

    def test_remove_file_only_removes_selection(self):
        app = self.app()
        first, second = Path(self.temp.name) / '第一份.md', Path(self.temp.name) / '第二份.md'
        first.write_text('正文', encoding='utf-8')
        second.write_text('正文', encoding='utf-8')
        app.files = [str(first), str(second)]
        result = app.action('/api/remove-file', {'path': str(first)})
        self.assertEqual([str(second)], result['files'])
        self.assertTrue(first.is_file())
        self.assertTrue(second.is_file())
        self.assertEqual([], app.action('/api/remove-file', {'path': str(second)})['files'])

    def test_remove_unknown_file_and_busy_task_are_rejected(self):
        app = self.app()
        app.files = ['file.md']
        for value in ('unknown.md', None, ['file.md']):
            with self.assertRaises(ValueError):
                app.action('/api/remove-file', {'path': value})
        app.job['running'] = True
        with self.assertRaises(ValueError):
            app.action('/api/remove-file', {'path': 'file.md'})
        app.job['running'] = False
        app.dialog_lock.acquire()
        try:
            with self.assertRaises(ValueError):
                app.action('/api/remove-file', {'path': 'file.md'})
        finally:
            app.dialog_lock.release()
        self.assertEqual(['file.md'], app.files)

    def test_page_removes_explanations_and_shows_bold_template_path(self):
        page = (web_ui.ROOT / 'web/index.html').read_text(encoding='utf-8')
        for explanation in ('支持批量转换', '默认使用配置中的模板', '自动记住上次选择',
                            '选择窗口默认打开工具的 templates 目录，支持 .docx 模板。',
                            'Word 文档与转换报告保存在此目录。同名文件自动添加序号，不覆盖原文件。',
                            '支持 .md / .markdown，可一次选择多个文件'):
            self.assertNotIn(explanation, page)
        self.assertNotIn('template-name', page)
        self.assertIn('<strong id="template-path"', page)
        self.assertIn("remove.onclick=()=>act('/api/remove-file',{path})", page)

    def test_quit_button_and_its_event_handler_are_removed(self):
        page = (web_ui.ROOT / 'web/index.html').read_text(encoding='utf-8')
        self.assertNotIn('id="quit"', page)
        self.assertNotIn("$('quit')", page)
        self.assertNotIn("'output','quit'", page)
        self.assertNotIn('退出工具', page)
        self.assertIn('id="convert" class="primary action-button"', page)
        self.assertNotIn('footer button{', page)

    def test_browser_close_triggers_server_shutdown(self):
        from unittest.mock import MagicMock
        browser = web_ui.BrowserWindow()
        browser.process = MagicMock()
        server = MagicMock()
        thread = browser.watch_closed(server, threading.Event())
        thread.join(timeout=2)
        self.assertFalse(thread.is_alive())
        server.shutdown.assert_called_once()

    def test_browser_watcher_ignores_normal_service_shutdown(self):
        from unittest.mock import MagicMock
        browser = web_ui.BrowserWindow()
        server = MagicMock()
        self.assertIsNone(browser.watch_closed(server, threading.Event()))
        browser.process = MagicMock()
        stopping = threading.Event()
        stopping.set()
        thread = browser.watch_closed(server, stopping)
        thread.join(timeout=2)
        server.shutdown.assert_not_called()

    def test_relaunched_browser_is_identified_by_own_profile_and_url(self):
        from unittest.mock import MagicMock
        from subprocess import CompletedProcess
        process = MagicMock()
        with patch.object(web_ui.subprocess, 'run', return_value=CompletedProcess([], 0, '[789]', '')) as run, \
             patch.object(web_ui, 'WindowsBrowserProcess', return_value=process) as windows_process:
            found = web_ui.find_browser_process('C:/temp/own-profile', 'http://127.0.0.1:1234/#token=own')
        self.assertIs(found, process)
        windows_process.assert_called_once_with(789)
        env = run.call_args.kwargs['env']
        self.assertEqual('C:/temp/own-profile', env['MDWORD_BROWSER_PROFILE'])
        self.assertEqual('http://127.0.0.1:1234/#token=own', env['MDWORD_BROWSER_URL'])
        self.assertIn("-notmatch '--type='", run.call_args.args[0][-1])

    def test_shutdown_stops_only_owned_active_job_processes(self):
        from unittest.mock import MagicMock
        app = self.app()
        process = MagicMock()
        process.poll.return_value = None
        with app.track_process(process):
            app.stop()
        process.kill.assert_called_once()
        self.assertTrue(app.stopping.is_set())
        self.assertEqual(set(), app.active_processes)
        late_process = MagicMock()
        with self.assertRaises(RuntimeError):
            with app.track_process(late_process):
                pass
        late_process.kill.assert_called_once()


if __name__ == '__main__':
    unittest.main()
