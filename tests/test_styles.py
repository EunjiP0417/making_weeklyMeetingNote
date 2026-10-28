"""지정 스타일의 문단 여백과 실제 내어쓰기 보존을 확인한다."""
import copy
import importlib.util
from pathlib import Path
import sys
import tempfile
import unittest

from lxml import etree as E

ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location('weekly_style_test', ROOT/'업무보고_자동취합.py')
report = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = report
SPEC.loader.exec_module(report)


def layout(shape):
    return (shape.get('tabPrIDRef'), tuple((report.local(x), tuple(sorted(x.attrib.items())))
            for x in shape.iter() if report.local(x) in ('intent', 'left', 'right')))


def paragraph_shape(shape):
    cloned = copy.deepcopy(shape)
    cloned.attrib.pop('id', None)
    return E.tostring(cloned)


class StyleReviewTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.pack = report.Package(report.TEMPLATE_PATH)
        self.styles = {s.get('name'): s for s in self.pack.header.iter() if report.local(s) == 'style'}

    def add_shape(self, intent=0):
        shape = copy.deepcopy(self.pack.shapes[self.styles['-'].get('paraPrIDRef')])
        ident = str(self.pack.next_shape)
        self.pack.next_shape += 1
        shape.set('id', ident)
        for x in shape.iter():
            if report.local(x) in ('left', 'right', 'intent'):
                x.set('value', str({'left': 7000, 'right': 4321, 'intent': intent}[report.local(x)]))
            elif report.local(x) == 'lineSpacing':
                x.set('value', '80')
        parent = self.pack.shapes[self.styles['-'].get('paraPrIDRef')].getparent()
        parent.append(shape)
        parent.set('itemCnt', str(len(parent)))
        self.pack.shapes[ident] = shape
        return ident

    def add_para(self, value, shape, style='바탕글'):
        p = self.pack.simple(self.pack.paras[0], value)
        p.set('paraPrIDRef', shape)
        p.set('styleIDRef', self.styles[style].get('id'))
        self.pack.root.append(p)
        return p

    def review(self):
        self.pack.data[self.pack.section_name] = E.tostring(self.pack.root, encoding='UTF-8', xml_declaration=True)
        self.pack.data['Contents/header.xml'] = E.tostring(self.pack.header, encoding='UTF-8', xml_declaration=True)
        source = self.folder/'source.hwpx'
        target = self.folder/'reviewed.hwpx'
        report.write_package(self.pack.data, source)
        original = source.read_bytes()
        report.review_styles(source, target, lambda _: None)
        self.assertEqual(source.read_bytes(), original)
        return report.Package(target)

    def test_regular_bullets_use_style_shape_and_remove_leading_spaces_across_runs(self):
        wrong = self.add_shape()
        self.add_para('     - 일정 : 본문 안  공백 유지', wrong)
        dot = self.add_para('  ', wrong)
        run = E.SubElement(dot, f'{{{report.HP}}}run', charPrIDRef='0')
        E.SubElement(run, f'{{{report.HP}}}t').text = '\u3000 · 세부 내용'
        self.add_para('    이어지는 본문', wrong)
        tabbed = self.add_para('', wrong)
        E.SubElement(tabbed.find('hp:run/hp:t', report.NS), f'{{{report.HP}}}tab', width='500').tail = '  - 수동 탭 문단'
        result = self.review()
        styles = {s.get('name'): s for s in result.header.iter() if report.local(s) == 'style'}
        for value, name in (('- 일정 : 본문 안  공백 유지', '-'), ('· 세부 내용', '·'),
                            ('이어지는 본문', '·'), ('- 수동 탭 문단', '-')):
            p = next(p for p in result.paras if report.text(p) == value)
            self.assertEqual(p.get('styleIDRef'), styles[name].get('id'))
            self.assertEqual(paragraph_shape(result.shapes[p.get('paraPrIDRef')]),
                             paragraph_shape(result.shapes[styles[name].get('paraPrIDRef')]))
        self.assertFalse(list(result.paras[-1].find('hp:run/hp:t', report.NS)))

    def test_hanging_layout_and_inline_tabs_linebreaks_are_preserved(self):
        hanging = self.add_shape(intent=-1234)
        p = self.add_para('    - 내용 : 첫 줄', hanging)
        t = p.find('hp:run/hp:t', report.NS)
        E.SubElement(t, f'{{{report.HP}}}tab', width='1000').tail = '탭 뒤 내용'
        E.SubElement(t, f'{{{report.HP}}}lineBreak').tail = '다음 줄'
        self.add_para('  기호 없는 이어지는 문단', hanging)
        before = layout(self.pack.shapes[hanging])
        result = self.review()
        styles = {s.get('name'): s for s in result.header.iter() if report.local(s) == 'style'}
        for p in result.paras[-2:]:
            self.assertEqual(p.get('styleIDRef'), styles['-'].get('id'))
            shape = result.shapes[p.get('paraPrIDRef')]
            self.assertEqual(layout(shape), before)
            # 원본의 잘못된 줄 간격은 보존 대상이 아니며 지정 스타일로 돌아간다.
            self.assertTrue(all(x.get('value') == '160' for x in shape.iter() if report.local(x) == 'lineSpacing'))
            self.assertFalse(report.text(p).startswith(' '))
        inline = result.paras[-2].find('hp:run/hp:t', report.NS)
        self.assertEqual([report.local(x) for x in inline], ['tab', 'lineBreak'])
        self.assertEqual([x.tail for x in inline], ['탭 뒤 내용', '다음 줄'])

    def test_title_date_section_and_numbered_titles_use_requested_style_indentation(self):
        wrong = self.add_shape()
        self.add_para('   8.  업무 제목', wrong)
        result = self.review()
        styles = {s.get('name'): s for s in result.header.iter() if report.local(s) == 'style'}
        targets = [('생방기술국 주요 업무 보고', '바탕글'), ('2026.09.23.(수) 생방기술2', '날짜'),
                   ('<추진실적>', '추진'), ('<추진계획>', '추진'), ('8. 업무 제목', '1,2,3')]
        for value, name in targets:
            with self.subTest(value=value):
                p = next(p for p in result.root.iter(f'{{{report.HP}}}p') if report.text(p) == value)
                self.assertEqual(p.get('styleIDRef'), styles[name].get('id'))
                self.assertEqual(paragraph_shape(result.shapes[p.get('paraPrIDRef')]),
                                 paragraph_shape(result.shapes[styles[name].get('paraPrIDRef')]))

    def test_known_markers_inside_tables_are_checked_without_changing_table_structure(self):
        wrong = self.add_shape()
        p = self.add_para('', wrong)
        table = E.SubElement(p.find('hp:run', report.NS), f'{{{report.HP}}}tbl', id='123')
        cell = E.SubElement(table, f'{{{report.HP}}}tc')
        for value in ('   - 표 안 업무', '   · 표 안 상세'):
            q = self.add_para(value, wrong)
            self.pack.root.remove(q)
            cell.append(q)
        result = self.review()
        table = next(x for x in result.root.iter() if report.local(x) == 'tbl' and x.get('id') == '123')
        self.assertEqual(len(table), 1)
        styles = {s.get('name'): s for s in result.header.iter() if report.local(s) == 'style'}
        for p, name in zip(table.iter(f'{{{report.HP}}}p'), ('-', '·')):
            self.assertEqual(p.get('styleIDRef'), styles[name].get('id'))
            self.assertFalse(report.text(p).startswith(' '))
            self.assertEqual(layout(result.shapes[p.get('paraPrIDRef')]),
                             layout(result.shapes[styles[name].get('paraPrIDRef')]))

    def test_plain_table_cells_keep_their_existing_layout(self):
        original_shape = self.add_shape(intent=-800)
        host = self.add_para('', original_shape)
        table = E.SubElement(host.find('hp:run', report.NS), f'{{{report.HP}}}tbl', id='124')
        cell = E.SubElement(table, f'{{{report.HP}}}tc')
        p = self.add_para('프로그램명', original_shape)
        self.pack.root.remove(p)
        cell.append(p)
        before = E.tostring(p)
        result = self.review()
        actual = next(q for q in result.root.iter(f'{{{report.HP}}}p') if report.text(q) == '프로그램명')
        self.assertEqual(E.tostring(actual), before)


if __name__ == '__main__':
    unittest.main()
