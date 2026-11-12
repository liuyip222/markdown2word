#!/usr/bin/env python3
"""Offline Markdown -> DOCX. Python 3.10+, Pandoc, python-docx. No AI/API."""
from __future__ import annotations
import argparse
from html.parser import HTMLParser
import collections
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
import urllib.parse
import zipfile
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent
DEFAULT_TEMPLATE = ROOT / 'templates' / 'template_标题不编号-列表第二行顶格.docx'
FENCE = re.compile(r'^( {0,3})(`{3,}|~{3,})(.*)$')
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def load_config():
    path = ROOT / 'config.json'
    return json.loads(path.read_text(encoding='utf-8-sig')) if path.exists() else {}


def windows_downloads_dir():
    """Resolve Windows' known Downloads folder, including relocated user folders."""
    import ctypes
    import uuid
    folder_id = ctypes.create_string_buffer(uuid.UUID('374DE290-123F-4565-9164-39C4925E467B').bytes_le)
    shell = ctypes.WinDLL('shell32')
    ole = ctypes.WinDLL('ole32')
    shell.SHGetKnownFolderPath.argtypes = [ctypes.c_void_p, ctypes.c_uint32, ctypes.c_void_p,
                                         ctypes.POINTER(ctypes.c_void_p)]
    shell.SHGetKnownFolderPath.restype = ctypes.c_long
    ole.CoTaskMemFree.argtypes = [ctypes.c_void_p]
    address = ctypes.c_void_p()
    try:
        result = shell.SHGetKnownFolderPath(folder_id, 0, None, ctypes.byref(address))
        if result < 0 or not address.value:
            raise OSError('无法读取 Windows 下载目录')
        return Path(ctypes.wstring_at(address.value))
    finally:
        if address.value:
            ole.CoTaskMemFree(address)


def default_output_dir():
    if os.name == 'nt':
        try:
            return windows_downloads_dir()
        except (OSError, AttributeError):
            # Registry fallback also respects Downloads moved to another drive.
            try:
                import winreg
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                    r'Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders') as key:
                    value, _ = winreg.QueryValueEx(key, '{374DE290-123F-4565-9164-39C4925E467B}')
                    if value:
                        return Path(os.path.expandvars(value)).expanduser()
            except OSError:
                pass
    return Path.home() / 'Downloads'


def find_pandoc(explicit=None):
    candidates = [explicit, shutil.which('pandoc'), str(ROOT / 'bin/pandoc.exe'),
                  str(Path(os.environ.get('LOCALAPPDATA', '')) / 'Pandoc/pandoc.exe'),
                  str(Path(os.environ.get('ProgramFiles', 'C:/Program Files')) / 'Pandoc/pandoc.exe')]
    for candidate in candidates:
        if candidate and Path(candidate).is_file():
            return str(Path(candidate).resolve())
    raise RuntimeError('未找到本机 Pandoc。安装 Pandoc 后重开终端，或使用 --pandoc 指定 pandoc.exe 的完整路径。')


def convert_wiki_images(line):
    """Convert Obsidian image embeds outside inline code; keep unsupported embeds intact."""
    output, count, index = [], 0, 0
    while index < len(line):
        if line[index] == chr(92) and index + 1 < len(line):
            output.append(line[index:index + 2]); index += 2; continue
        if line[index] == '`':
            end = index + 1
            while end < len(line) and line[end] == '`': end += 1
            marker = line[index:end]
            close = re.search(r'(?<!`)'+re.escape(marker)+r'(?!`)', line[end:])
            if close:
                stop = end + close.end()
                output.append(line[index:stop]); index = stop; continue
            output.append(marker); index = end; continue
        match = re.match(r'!\[\[([^\]\r\n]+)\]\]', line[index:])
        if match:
            fields = match[1].split('|', 1)
            target = fields[0].strip().replace(chr(92), '/')
            if re.search(r'\.(?:png|jpe?g|gif|bmp|tiff?|svg|webp|emf)(?:#.*)?$', target, re.I):
                label = fields[1].strip() if len(fields) > 1 else ''
                size = re.fullmatch(r'(\d+)(?:x(\d+))?', label)
                attrs = ''
                if size:
                    attrs = '{width=' + size[1] + 'px' + (' height=' + size[2] + 'px' if size[2] else '') + '}'
                    label = ''
                label = label.replace('[', r'\[').replace(']', r'\]')
                url = urllib.parse.quote(target, safe='/%:')
                output.append('!['+label+'](<'+url+'>)'+attrs)
                count += 1
                index += len(match[0]); continue
        output.append(line[index]); index += 1
    return ''.join(output), count


def normalize(text, close_fence=False):
    """Conservative line fixes; never infer a heading from fenced code."""
    lines = text.replace('\r\n', '\n').replace('\r', '\n').splitlines()
    out, fixes, warnings = [], [], []
    active = None
    front = bool(lines and lines[0].strip() == '---')
    front_end = next((i for i in range(1, len(lines)) if lines[i].strip() in ('---', '...')), None) if front else None
    i = 0

    def before():
        if out and out[-1].strip():
            out.append('')

    while i < len(lines):
        line = lines[i]
        number = i + 1
        if front and front_end is not None and i <= front_end:
            out.append(line)
            i += 1
            continue
        m = FENCE.match(line)
        if active:
            out.append(line)
            if m and m[2][0] == active[0] and len(m[2]) >= active[1] and not m[3].strip():
                active = None
                out.append('')
            i += 1
            continue
        if m:
            # A backtick info string may not itself contain backticks.
            if m[2][0] == '`' and '`' in m[3]:
                out.append(line)
                i += 1
                continue
            before()
            out.append(line)
            active = (m[2][0], len(m[2]), number)
            i += 1
            continue
        # Indented code and raw HTML are not rewritten.
        if line.startswith(('    ', '\t')) or re.match(r'^\s*<', line):
            out.append(line)
            if re.match(r'^\s*<(pre|script|style|table)\b', line, re.I):
                warnings.append(f'第 {number} 行含原始 HTML，请检查转换效果。')
            i += 1
            continue
        line, converted_images = convert_wiki_images(line)
        if converted_images:
            fixes.append(f'第 {number} 行：将 {converted_images} 个 Obsidian 图片嵌入转为标准 Markdown。')
        heading = re.match(r'^( {0,3})(#{1,6})(?!#)(.*)$', line)
        if heading and heading[3].strip() and not heading[3].lstrip().startswith(('!', '#')):
            new = f'{heading[1]}{heading[2]} {heading[3].strip()}'
            if new != line:
                fixes.append(f'第 {number} 行：规范标题标记空格。')
            if out and out[-1].strip():
                fixes.append(f'第 {number} 行：标题前添加空行。')
            before()
            out.extend([new, ''])
            i += 1
            continue
        # Only repair clear prose markers; do not rewrite 1.2 / -1 / **bold**.
        numbered = re.match(r'^( {0,3})(\d+[.)])([^\s\d].*)$', line)
        bullet = re.match(r'^( {0,3})([-+])([A-Za-z\u3400-\u9fff].*)$', line)
        if numbered or bullet:
            m = numbered or bullet
            line = f'{m[1]}{m[2]} {m[3]}'
            fixes.append(f'第 {number} 行：列表标记后添加空格。')
        empty = re.match(r'^( {0,3})(\d+[.)])\s*$', line)
        if empty:
            j = i + 1
            while j < len(lines) and not lines[j].strip():
                j += 1
            upcoming = convert_wiki_images(lines[j].strip())[0] if j < len(lines) else ''
            if re.match(r'^!\[.*\]\(.*\)(?:\{.*\})?\s*$', upcoming):
                line = f'{empty[1]}{empty[2]} {upcoming}'
                fixes.append(f'第 {number} 行：将空编号与后续图片合并为列表项。')
                i = j
        list_item = re.match(r'^ {0,3}(?:[-+*] |\d+[.)] )', line)
        prev_item = out and re.match(r'^ {0,3}(?:[-+*] |\d+[.)] )', out[-1])
        if list_item and out and out[-1].strip() and not prev_item:
            before()
            fixes.append(f'第 {number} 行：列表前添加空行。')
        table_row = line.lstrip().startswith('|') and line.rstrip().endswith('|')
        if table_row and out and out[-1].strip() and not out[-1].lstrip().startswith('|'):
            before()
            fixes.append(f'第 {number} 行：表格前添加空行。')
        if out and out[-1].lstrip().startswith('|') and not table_row and line.strip():
            before()
        # Preserve hard breaks; avoid accumulating blank lines on repeated runs.
        if line.strip() or not out or out[-1].strip():
            out.append(line)
        i += 1
    if active:
        message = f'第 {active[2]} 行开始的代码围栏未闭合。'
        if not close_fence:
            raise ValueError(message + '请补齐；确认只缺末尾围栏时可使用 --close-unclosed-fence。')
        out.append(active[0] * active[1])
        fixes.append(message + '已在文档末尾补齐（显式启用）。')
    if front and front_end is None:
        warnings.append('开头存在 --- 但未发现 YAML 元数据结束标记，请检查。')
    while out and not out[-1].strip():
        out.pop()
    return '\n'.join(out) + '\n', fixes, warnings


def check_pipe_tables(text):
    """Reject clear ragged pipe tables rather than guessing missing cell contents."""
    def cells(line):
        parts, current, ticks, i = [], [], 0, 0
        while i < len(line):
            ch = line[i]
            if ch == chr(92) and i + 1 < len(line):
                current.extend(line[i:i + 2]); i += 2; continue
            if ch == '`':
                end = i + 1
                while end < len(line) and line[end] == '`':
                    end += 1
                count = end - i
                if not ticks: ticks = count
                elif ticks == count: ticks = 0
                current.append(line[i:end]); i = end; continue
            if ch == '|' and not ticks:
                parts.append(''.join(current).strip()); current = []
            else:
                current.append(ch)
            i += 1
        parts.append(''.join(current).strip())
        if line.lstrip().startswith('|'): parts = parts[1:]
        if line.rstrip().endswith('|'): parts = parts[:-1]
        return parts
    lines = text.splitlines()
    active = None
    expected = None
    for index, line in enumerate(lines):
        match = FENCE.match(line)
        if active:
            if match and match[2][0] == active[0] and len(match[2]) >= active[1] and not match[3].strip():
                active = None
            continue
        if match:
            active = (match[2][0],len(match[2])); expected = None; continue
        if line.startswith(('    ', chr(9))):
            expected = None; continue
        if '|' not in line or not line.strip():
            expected = None; continue
        columns = cells(line)
        if len(columns) > 1 and all(re.fullmatch(r':?-{3,}:?', c) for c in columns):
            expected = len(columns)
            if index and len(cells(lines[index-1])) != expected:
                raise ValueError(f'第 {index+1} 行表格分隔符列数与表头不一致，请人工修复。')
        elif expected and line.lstrip().startswith('|'):
            if len(columns) != expected:
                raise ValueError(f'第 {index+1} 行表格有 {len(columns)} 列，预期 {expected} 列；为避免丢失内容已停止。')


def run_pandoc(exe, args, cwd, data=None):
    result = subprocess.run([exe, *args], input=data, cwd=cwd, stdout=subprocess.PIPE,
                            stderr=subprocess.PIPE, creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
    stderr = result.stderr.decode('utf-8', errors='replace')
    if result.returncode:
        raise RuntimeError('Pandoc 转换失败：\n' + stderr)
    return result.stdout, stderr


def walk(node):
    if isinstance(node, dict):
        if 't' in node:
            yield node
        for value in node.values():
            yield from walk(value)
    elif isinstance(node, list):
        for value in node:
            yield from walk(value)


def ast_summary(ast):
    nodes = list(walk(ast.get('blocks', [])))
    counts = collections.Counter(n['t'] for n in nodes)
    return {k: counts[k] for k in ('Header', 'Table', 'CodeBlock', 'Link', 'Image', 'Note', 'OrderedList', 'BulletList')}


def preflight(ast, source, allow_missing):
    warnings, missing = [], []
    nodes = list(walk(ast.get('blocks', [])))
    class Images(HTMLParser):
        def handle_starttag(self, tag, attrs):
            if tag.lower() in ('img', 'image', 'source', 'embed', 'object', 'link', 'script'):
                for key, value in attrs:
                    if key.lower() in ('src', 'href', 'data', 'srcset') and value:
                        if re.search(r'(?:https?:)?//', value):
                            raise ValueError('离线模式不读取 HTML 中的远程资源：' + value)
                        if tag.lower() == 'img' and key.lower() == 'src':
                            raw_images.append(value)
    raw_images = []
    for node in nodes:
        if node['t'] in ('RawBlock', 'RawInline') and node['c'][0] == 'html':
            Images().feed(node['c'][1])
    image_nodes = [n for n in nodes if n['t'] == 'Image']
    image_nodes.extend({'t': 'Image', 'c': [None, None, [url, '']]} for url in raw_images)
    for node in image_nodes:
        if node['t'] == 'Image':
            url = node['c'][-1][0]
            parsed = urllib.parse.urlsplit(url)
            if parsed.scheme in ('http', 'https') or url.startswith('//'):
                raise ValueError('检测到远程图片。离线模式不下载远程资源，请先保存图片并改用本地路径：' + url)
            if parsed.scheme == 'data':
                continue
            if parsed.scheme == 'file':
                raw = urllib.parse.unquote(parsed.path).lstrip('/') if os.name == 'nt' else urllib.parse.unquote(parsed.path)
            else:
                raw = urllib.parse.unquote(url.split('#', 1)[0])
            path = Path(raw)
            path = path if path.is_absolute() else source.parent / path
            if not path.is_file():
                missing.append(str(path))
    if missing:
        message = '缺失本地图片：' + '; '.join(missing)
        if not allow_missing:
            raise FileNotFoundError(message)
        warnings.append(message)
    # Visible raw [^label] often means an undefined footnote.
    anchors = set()
    for node in nodes:
        c = node.get('c')
        if node['t'] == 'Header':
            anchors.add(c[1][0])
        elif node['t'] in ('Div', 'Span', 'CodeBlock'):
            anchors.add(c[0][0])
    for node in nodes:
        if node['t'] == 'Link':
            url = node['c'][-1][0]
            if url.startswith('#') and urllib.parse.unquote(url[1:]) not in anchors:
                warnings.append('未找到文档内部引用目标：' + url)
            elif re.search(r'\.md(?:#.*)?$', url, re.I):
                warnings.append('跨 Markdown 文件引用仍保留为原路径，请复核：' + url)
    for node in nodes:
        if node['t'] == 'Str' and re.search(r'\[\^[^\]]+\]', node['c']):
            warnings.append('发现未解析的脚注引用：' + node['c'])
    return warnings


def package_signature(path):
    """Formatting must not change text, links, images, bookmarks or table topology."""
    with zipfile.ZipFile(path) as archive:
        root = ET.fromstring(archive.read('word/document.xml'))
        rels = ET.fromstring(archive.read('word/_rels/document.xml.rels'))
        return {
            'texts': [n.text for n in root.iter(W + 't')],
            'notes': {name: archive.read(name) for name in ('word/footnotes.xml', 'word/endnotes.xml') if name in archive.namelist()},
            'links': [tuple(sorted(n.attrib.items())) for n in root.iter(W + 'hyperlink')],
            'bookmarks': [tuple(sorted(n.attrib.items())) for n in root.iter(W + 'bookmarkStart')],
            'tables': [[len(row.findall(W + 'tc')) for row in table.findall(W + 'tr')] for table in root.iter(W + 'tbl')],
            'relationships': sorted(tuple(sorted(n.attrib.items())) for n in rels),
            'media': sorted(n for n in archive.namelist() if n.startswith('word/media/')),
        }


def apply_template_table_style(path, style_name='Table'):
    """Use the reference template's actual named style, never a hard-coded style ID."""
    with zipfile.ZipFile(path) as archive:
        if not list(ET.fromstring(archive.read('word/document.xml')).iter(W + 'tbl')):
            return 0
    try:
        from docx import Document
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        from docx.enum.style import WD_STYLE_TYPE
    except ModuleNotFoundError as error:
        raise RuntimeError('应用模板表格样式需要 python-docx；执行 python -m pip install -r requirements.txt。') from error
    document = Document(path)
    try:
        style = document.styles[style_name]
    except KeyError as error:
        raise ValueError(f'参考模板缺少表格样式 {style_name!r}；请修改 config.json 的 table_style，或设为空字符串以禁用。') from error
    if style.type != WD_STYLE_TYPE.TABLE:
        raise ValueError(f'{style_name!r} 不是表格样式。')
    from docx.table import Table
    tables = [Table(element, document._body) for element in document.element.body.iter(qn('w:tbl'))]
    for table in tables:
        table.style = style
        properties = table._tbl.tblPr
        for name in ('tblBorders', 'shd'):
            for child in list(properties.findall(qn('w:' + name))):
                properties.remove(child)
        look = properties.find(qn('w:tblLook'))
        if look is None:
            look = OxmlElement('w:tblLook')
            properties.append(look)
        look.attrib.clear()
        for name, value in {'val': '0620', 'firstRow': '1', 'lastRow': '0', 'firstColumn': '0', 'lastColumn': '0', 'noHBand': '1', 'noVBand': '1'}.items():
            look.set(qn('w:' + name), value)
        for cell in table._tbl.iter(qn('w:tcPr')):
            for name in ('tcBorders', 'shd'):
                for child in list(cell.findall(qn('w:' + name))):
                    cell.remove(child)
    document.save(path)
    return len(tables)


def unique_path(path, overwrite):
    if overwrite or not path.exists():
        return path
    for index in range(1, 10000):
        candidate = path.with_name(f'{path.stem} ({index}){path.suffix}')
        if not candidate.exists():
            return candidate
    raise RuntimeError('同名文件太多，请更换输出路径。')


def convert(source, args, config):
    source = Path(source).expanduser().resolve(strict=True)
    if source.suffix.lower() not in ('.md', '.markdown'):
        raise ValueError('仅支持 .md / .markdown：' + str(source))
    exe = find_pandoc(args.pandoc or config.get('pandoc'))
    template = Path(args.template or config.get('template') or DEFAULT_TEMPLATE).expanduser()
    if not template.is_absolute():
        template = ROOT / template
    if not template.is_file():
        raise FileNotFoundError('模板不存在：' + str(template))
    text = source.read_text(encoding=args.encoding)
    fixed, fixes, warnings = (text, [], []) if args.no_fix else normalize(text, args.close_unclosed_fence)
    check_pipe_tables(fixed)
    ast_bytes, diagnostic = run_pandoc(exe, ['-f', 'markdown', '-t', 'json'], source.parent, fixed.encode('utf-8'))
    if diagnostic:
        warnings.append(diagnostic.strip())
    ast = json.loads(ast_bytes)
    summary = ast_summary(ast)
    warnings.extend(preflight(ast, source, args.allow_missing_images))
    if args.strict and warnings:
        raise ValueError('严格模式下检查未通过：\n' + '\n'.join(warnings))
    if args.check_only:
        return {'source': str(source), 'counts': summary, 'fixes': fixes, 'warnings': warnings, 'check_only': True}
    output_dir = Path(args.output_dir or config.get('output_dir') or default_output_dir()).expanduser().resolve()
    target = Path(args.output).expanduser().resolve() if args.output else output_dir / (source.stem + '.docx')
    if target.suffix.lower() != '.docx':
        raise ValueError('输出文件扩展名必须是 .docx。')
    if source == target:
        raise ValueError('不能覆盖源文件。')
    target.parent.mkdir(parents=True, exist_ok=True)
    target = unique_path(target, args.overwrite)
    highlight = args.syntax_highlighting if args.syntax_highlighting is not None else config.get('syntax_highlighting', False)
    highlight_options = [] if highlight else ['--no-highlight']
    pipeline = args.pipeline or config.get('pipeline', 'auto')
    # Direct conversion keeps native footnotes. Otherwise preserve existing skill's HTML pipeline.
    if pipeline == 'auto':
        pipeline = 'direct' if summary['Note'] else 'html'
    with tempfile.TemporaryDirectory(prefix='.mdword-', dir=target.parent) as temporary:
        stage = Path(temporary) / 'converted.docx'
        data = fixed.encode('utf-8')
        if pipeline == 'html':
            data, diagnostic = run_pandoc(exe, ['-f', 'markdown', '-t', 'html', *highlight_options], source.parent, data)
            if diagnostic:
                warnings.append(diagnostic.strip())
        options = ['-f', 'html' if pipeline == 'html' else 'markdown', '-o', str(stage),
                   '--reference-doc', str(template), '--resource-path', str(source.parent),
                   '--lua-filter', str(ROOT / 'markdown-to-docx.lua'), *highlight_options]
        _, diagnostic = run_pandoc(exe, options, source.parent, data)
        if diagnostic:
            warnings.append(diagnostic.strip())
        before = package_signature(stage)
        beautify = not args.no_table_beautify and config.get('beautify_tables', False)
        if beautify:
            try:
                from beautify_docx_tables import beautify_tables
            except ModuleNotFoundError as error:
                raise RuntimeError('缺少 python-docx；执行 python -m pip install -r requirements.txt，或使用 --no-table-beautify。') from error
            beautify_tables(stage)
        elif config.get('table_style', 'Table'):
            apply_template_table_style(stage, config.get('table_style', 'Table'))
        if before != package_signature(stage):
            raise RuntimeError('应用表格样式后内容/链接/结构校验不一致，已停止输出。')
        with zipfile.ZipFile(stage) as archive:
            xml = ET.fromstring(archive.read('word/document.xml'))
            parts = [xml] + [ET.fromstring(archive.read(name)) for name in ('word/footnotes.xml', 'word/endnotes.xml') if name in archive.namelist()]
            links = sum(len(list(part.iter(W + 'hyperlink'))) for part in parts)
            tables = sum(len(list(part.iter(W + 'tbl'))) for part in parts)
            if links < summary['Link']:
                warnings.append(f'链接数量需人工复核：源 {summary["Link"]}，Word（含脚注） {links}。')
            if tables != summary['Table']:
                warnings.append(f'表格数量需人工复核：源 {summary["Table"]}，Word {tables}。')
        if args.strict and warnings:
            raise ValueError('严格模式下停止交付：\n' + '\n'.join(warnings))
        os.replace(stage, target)
    report = {'source': str(source), 'output': str(target), 'template': str(template),
              'pipeline': pipeline, 'syntax_highlighting': highlight, 'table_beautify': beautify,
              'table_style': config.get('table_style', 'Table') if not beautify else None, 'counts': summary,
              'fixes': fixes, 'warnings': warnings,
              'validation': 'OOXML 内容/链接/表格结构校验；未进行 Word 页面渲染或网站可达性验证。'}
    report_path = target.with_suffix('.conversion-report.json')
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    if args.keep_fixed_md:
        target.with_suffix('.fixed.md').write_text(fixed, encoding='utf-8')
    return report


def parser():
    p = argparse.ArgumentParser(description='离线 Markdown 转 Word：保守格式修复 + 本机 Pandoc + 默认模板 + 表格美化')
    p.add_argument('inputs', nargs='*', help='Markdown 文件或目录；多个路径用空格分隔')
    p.add_argument('-o', '--output', help='单文件指定完整 DOCX 路径')
    p.add_argument('-d', '--output-dir', help='输出目录，默认当前用户的系统下载目录，可在 config.json 修改')
    p.add_argument('--template', help='自定义 Word 参考模板')
    p.add_argument('--pandoc', help='pandoc.exe 路径')
    p.add_argument('--encoding', default='utf-8-sig', help='源编码，默认 UTF-8；可指定 gb18030')
    p.add_argument('--recursive', action='store_true', help='递归转换目录中的 Markdown')
    p.add_argument('--overwrite', action='store_true', help='允许覆盖同名 DOCX（默认添加序号）')
    p.add_argument('--no-fix', action='store_true', help='不修复 Markdown 格式')
    p.add_argument('--no-table-beautify', action='store_true', help='关闭表格二次美化（默认已关闭）')
    highlight_group = p.add_mutually_exclusive_group()
    highlight_group.add_argument('--syntax-highlighting', dest='syntax_highlighting', action='store_true', help='开启代码语法高亮（会产生 NormalTok/StringTok 等字符样式）')
    highlight_group.add_argument('--no-syntax-highlighting', dest='syntax_highlighting', action='store_false', help='关闭代码语法高亮，使用模板 Source Code 段落样式')
    p.set_defaults(syntax_highlighting=None)
    p.add_argument('--keep-fixed-md', action='store_true', help='另外输出修复版 Markdown；原文件不变')
    p.add_argument('--close-unclosed-fence', action='store_true', help='显式允许在末尾补齐未闭合代码围栏')
    p.add_argument('--allow-missing-images', action='store_true', help='允许缺失本地图片并写入报告')
    p.add_argument('--strict', action='store_true', help='有警告时停止交付')
    p.add_argument('--check-only', action='store_true', help='只检查和预览修复建议，不输出 DOCX')
    p.add_argument('--pipeline', choices=['auto', 'html', 'direct'], help='默认 auto：有脚注时直转，其余使用 Skill 原 HTML 流程')
    p.add_argument('--gui', action='store_true', help='打开文件选择窗口；无参数时也会打开')
    return p


def gui(args, config):
    import tkinter as tk
    from tkinter import filedialog, messagebox
    root = tk.Tk()
    root.withdraw()
    files = filedialog.askopenfilenames(title='选择 Markdown 文件（可多选）', filetypes=[('Markdown', '*.md *.markdown'), ('所有文件', '*.*')])
    if not files:
        root.destroy()
        return 0
    default_dir = args.output_dir or config.get('output_dir') or str(default_output_dir())
    selected = filedialog.askdirectory(title='选择输出目录（取消则使用默认目录）', initialdir=default_dir if Path(default_dir).is_dir() else str(Path.home()))
    args.output_dir = selected or default_dir
    success, errors = [], []
    for path in files:
        try:
            report = convert(path, args, config)
            success.append(report.get('output', path) + (f'（{len(report["warnings"])} 条警告，详见报告）' if report['warnings'] else ''))
        except Exception as error:
            errors.append(f'{path}\n{error}')
    text = '\n'.join(success) + ('\n\n失败：\n' + '\n'.join(errors) if errors else '')
    (messagebox.showwarning if errors else messagebox.showinfo)('转换结果', text or '未处理文件')
    root.destroy()
    return 1 if errors else 0


def main(argv=None):
    p = parser()
    args = p.parse_args(argv)
    config = load_config()
    if args.gui or not args.inputs:
        return gui(args, config)
    files = []
    for item in args.inputs:
        path = Path(item).expanduser()
        if path.is_dir():
            files.extend(sorted(f for f in (path.rglob('*') if args.recursive else path.iterdir()) if f.is_file() and f.suffix.lower() in ('.md', '.markdown') and not f.name.endswith('.fixed.md')))
        else:
            files.append(path)
    files = list(dict.fromkeys(f.resolve() for f in files))
    if not files:
        p.error('没有找到 Markdown 文件。')
    if args.output and len(files) != 1:
        p.error('-o 仅适用于单个文件；批量转换请用 -d。')
    failures = 0
    for source in files:
        try:
            report = convert(source, args, config)
            print(json.dumps(report, ensure_ascii=False, indent=2))
        except Exception as error:
            failures += 1
            print(f'[失败] {source}: {error}', file=sys.stderr)
    return 1 if failures else 0


if __name__ == '__main__':
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        raise SystemExit(130)
