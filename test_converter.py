import argparse
from pathlib import Path
import tempfile
import unittest
import sys
import json
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import md_to_word as app
from beautify_docx_tables import beautify_tables


class FormattingTests(unittest.TestCase):
    def test_code_and_hardbreak_preserved(self):
        src = '正文  \n下一行\n##标题\n正文\n```text\n### 文件名\n#comment\n1.代码\n```\n'
        fixed, changes, warnings = app.normalize(src)
        self.assertIn('正文  \n下一行', fixed)
        self.assertIn('## 标题', fixed)
        self.assertIn('```text\n### 文件名\n#comment\n1.代码\n```', fixed)
        self.assertTrue(changes)

    def test_idempotent(self):
        fixed, _, _ = app.normalize('##标题\n正文\n1.第一项\n2.第二项\n')
        self.assertEqual(fixed, app.normalize(fixed)[0])

    def test_tilde_and_nested_backticks(self):
        src = '~~~~text\n```\n###comment\n~~~~\n'
        self.assertEqual(src, app.normalize(src)[0])

    def test_unclosed_requires_optin(self):
        with self.assertRaises(ValueError):
            app.normalize('```text\n## comment\n')
        self.assertTrue(app.normalize('```text\nx\n', True)[0].endswith('```\n'))

    def test_frontmatter_indent_and_numbers(self):
        src = '---\ntitle: 标题\n---\n\n    ###code\n1.2\n**bold**\n-1\n'
        self.assertEqual(src, app.normalize(src)[0])

    def test_empty_number_image(self):
        fixed, _, _ = app.normalize('1.\n\n![](assets/a.png)\n')
        self.assertIn('1. ![](assets/a.png)', fixed)

    def test_tables_preserve_links_and_structure(self):
        with tempfile.TemporaryDirectory(dir=Path(__file__).parent) as directory:
            path = Path(directory)/'table.docx'
            d = Document()
            table = d.add_table(rows=3, cols=3)
            for i,row in enumerate(table.rows):
                for j,cell in enumerate(row.cells): cell.text = f'{i},{j}'
            hyperlink = OxmlElement('w:hyperlink'); hyperlink.set(qn('w:anchor'), 'example')
            run = OxmlElement('w:r'); text=OxmlElement('w:t'); text.text='linked text';run.append(text);hyperlink.append(run)
            table.cell(0,0).paragraphs[0]._p.append(hyperlink)
            d.save(path)
            signature=app.package_signature(path)
            beautify_tables(path)
            self.assertEqual(signature, app.package_signature(path))
            first=app.package_signature(path)
            beautify_tables(path)
            self.assertEqual(first,app.package_signature(path))

    def test_many_columns_positive_width(self):
        from beautify_docx_tables import _column_widths
        table=Document().add_table(rows=2,cols=35)
        widths=_column_widths(table,9000)
        self.assertEqual(sum(widths),9000)
        self.assertTrue(all(w>0 for w in widths))

    def test_wiki_images_with_spaces_and_size(self):
        result, count = app.convert_wiki_images('![[asset/图 1.png|300x200]]')
        self.assertEqual(count, 1)
        self.assertIn('width=300px height=200px',result)
        self.assertIn('%20',result)
        self.assertIn('![',result)

    def test_wiki_images_preserve_code_and_escapes(self):
        src='`![[x.png]]` and \\![[y.png]] and ![[z.png]]'
        result,count=app.convert_wiki_images(src)
        self.assertEqual(count,1)
        self.assertIn('`![[x.png]]`',result)
        self.assertIn(r'\![[y.png]]',result)
        code='```text\n![[code.png]]\n```\n'
        self.assertEqual(app.normalize(code)[0],code)

    def test_wiki_images_normalize_idempotent(self):
        result,fixes,_=app.normalize('![[asset/图 1.png|说明]]\n')
        self.assertTrue(fixes)
        self.assertNotIn('![[',result)
        self.assertEqual(result,app.normalize(result)[0])

    def test_ragged_table(self):
        with self.assertRaises(ValueError):
            app.check_pipe_tables('| A | B |\n| --- | --- |\n| only one |\n')
        app.check_pipe_tables('| A | B |\n| --- | --- |\n| `a|b` | escaped\\|pipe |\n')

    def test_remote_resource_rejected(self):
        ast={'blocks':[{'t':'Image','c':[[],[],['https://example.com/a.png','']]}]}
        with self.assertRaises(ValueError): app.preflight(ast,Path('a.md'),False)

    def test_missing_image_and_anchor(self):
        ast={'blocks':[{'t':'Image','c':[[],[],['assets/not-exists.png','']]}]}
        with self.assertRaises(FileNotFoundError): app.preflight(ast,Path('a.md'),False)
        ast={'blocks':[{'t':'Link','c':[[],[],['#absent','']]}]}
        self.assertTrue(app.preflight(ast,Path('a.md'),False))



class IntegrationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        try: app.find_pandoc()
        except RuntimeError: raise unittest.SkipTest('Pandoc not installed')

    def test_complete_pipeline_and_no_overwrite(self):
        import base64
        import zipfile
        with tempfile.TemporaryDirectory(dir=Path(__file__).parent) as directory:
            root=Path(directory)
            (root/'assets').mkdir()
            (root/'assets'/'图 1.png').write_bytes(base64.b64decode('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jGz0AAAAASUVORK5CYII='))
            source=root/'中文报告.md'
            text='# 测试报告\n\n##概览 {#overview}\n正文  \n硬换行\n\n[引用](https://example.com) [内部](#overview)\n\n| 编号 | 内容 |\n| --- | --- |\n| 1 | [引用](https://example.com) |\n\n```text\n### code comment\n1. not list\n```\n\n![[assets/图 1.png]]\n'
            source.write_text(text,encoding='utf-8')
            args=app.parser().parse_args([str(source),'-d',str(root/'out'),'--keep-fixed-md'])
            report=app.convert(source,args,{})
            self.assertEqual(source.read_text(encoding='utf-8'),text)
            self.assertFalse(report['warnings'])
            self.assertEqual(report['counts']['Table'],1)
            self.assertEqual(report['counts']['Link'],3)
            self.assertEqual(report['counts']['Image'],1)
            output=Path(report['output'])
            self.assertTrue(output.is_file())
            with zipfile.ZipFile(output) as archive:
                self.assertTrue(any(name.startswith('word/media/') for name in archive.namelist()))
            self.assertTrue(output.with_suffix('.conversion-report.json').is_file())
            self.assertTrue(output.with_suffix('.fixed.md').is_file())
            again=app.convert(source,args,{})
            self.assertNotEqual(report['output'],again['output'])
            self.assertEqual(Path(again['output']).stem,'中文报告 (1)')

    def test_native_footnotes(self):
        import zipfile
        with tempfile.TemporaryDirectory(dir=Path(__file__).parent) as directory:
            root=Path(directory);source=root/'notes.md'
            source.write_text('# Report\n\nText[^a].\n\n[^a]: Footnote [link](https://example.com).\n',encoding='utf-8')
            args=app.parser().parse_args([str(source),'-d',str(root/'out')])
            report=app.convert(source,args,{})
            self.assertEqual(report['pipeline'],'direct')
            with zipfile.ZipFile(report['output']) as archive:
                self.assertIn('word/footnotes.xml',archive.namelist())
                self.assertIn(b'Footnote',archive.read('word/footnotes.xml'))

if __name__ == '__main__':
    unittest.main()
