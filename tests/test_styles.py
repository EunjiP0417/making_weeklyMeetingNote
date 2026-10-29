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

    def test_template_title_date_and_surrounding_layout_are_unchanged(self):
        original = report.Package(report.TEMPLATE_PATH)
        stop = next(i for i, p in enumerate(original.paras) if report.heading(p))
        before = [E.tostring(p) for p in original.paras[:stop]]
        result = self.review()
        self.assertEqual([E.tostring(p) for p in result.paras[:stop]], before)
        # 도형 안의 제목까지 본문 바탕글 스타일로 덮어쓰지 않는다.
        title = next(p for p in result.root.iter(f'{{{report.HP}}}p')
                     if report.text(p) == '생방기술국 주요 업무 보고')
        original_title = next(p for p in original.root.iter(f'{{{report.HP}}}p')
                              if report.text(p) == report.text(title))
        self.assertEqual(title.attrib, original_title.attrib)
        self.assertEqual(paragraph_shape(result.shapes[title.get('paraPrIDRef')]),
                         paragraph_shape(original.shapes[original_title.get('paraPrIDRef')]))

    def test_section_and_numbered_titles_use_requested_style_indentation(self):
        wrong = self.add_shape()
        self.add_para('   8.  업무 제목', wrong)
        result = self.review()
        styles = {s.get('name'): s for s in result.header.iter() if report.local(s) == 'style'}
        targets = [('<추진실적>', '추진'), ('<추진계획>', '추진'), ('8. 업무 제목', '1,2,3')]
        for value, name in targets:
            with self.subTest(value=value):
                p = next(p for p in result.root.iter(f'{{{report.HP}}}p') if report.text(p) == value)
                self.assertEqual(p.get('styleIDRef'), styles[name].get('id'))
                self.assertEqual(paragraph_shape(result.shapes[p.get('paraPrIDRef')]),
                                 paragraph_shape(result.shapes[styles[name].get('paraPrIDRef')]))

    def test_all_table_content_and_host_paragraph_keep_original_formatting(self):
        wrong = self.add_shape()
        hanging = self.add_shape(intent=-800)
        p = self.add_para('   - 표가 놓인 줄', wrong)
        table = E.SubElement(p.find('hp:run', report.NS), f'{{{report.HP}}}tbl', id='123')
        cell = E.SubElement(table, f'{{{report.HP}}}tc')
        for value, shape, style in (('   - 표 안 업무', wrong, '-'),
                                    ('   · 표 안 상세', hanging, '\uf0a0'),
                                    ('   7. 표 안 번호', wrong, '1,2,3'),
                                    ('<추진계획>', wrong, '추진'),
                                    ('  기호 없는 셀 내용', wrong, '-'),
                                    ('10.26.(월)', wrong, '-'),
                                    ('가수, 총 2화 녹음', wrong, '-'),
                                    ('  ', hanging, '바탕글')):
            q = self.add_para(value, shape, style)
            self.pack.root.remove(q)
            cell.append(q)
        inline = cell[0].find('hp:run/hp:t', report.NS)
        E.SubElement(inline, f'{{{report.HP}}}tab', width='500').tail = ' 탭 뒤 공백'
        E.SubElement(inline, f'{{{report.HP}}}lineBreak').tail = '  둘째 줄'
        nested = E.SubElement(cell[-1].find('hp:run', report.NS), f'{{{report.HP}}}tbl', id='nested')
        nested_cell = E.SubElement(nested, f'{{{report.HP}}}tc')
        q = self.add_para('  - 중첩 표 본문', wrong, '-')
        self.pack.root.remove(q)
        nested_cell.append(q)
        before = E.tostring(p, method='c14n')
        result = self.review()
        actual = next(q for q in result.paras if report.text(q) == '   - 표가 놓인 줄')
        self.assertEqual(E.tostring(actual, method='c14n'), before)
        for ident in (wrong, hanging):
            self.assertEqual(paragraph_shape(result.shapes[ident]), paragraph_shape(self.pack.shapes[ident]))

    def test_picture_host_and_caption_preserved_but_adjacent_body_is_normalized(self):
        wrong = self.add_shape()
        host = self.add_para('   - 사진 설명', wrong, '-')
        run = host.find('hp:run', report.NS)
        picture = E.SubElement(run, f'{{{report.HP}}}pic', id='picture1')
        E.SubElement(picture, f'{{{report.HP}}}pos', treatAsChar='1', horzOffset='1234')
        E.SubElement(picture, f'{{{report.HP}}}sz', width='10000', height='8000')
        E.SubElement(picture, f'{{{report.HP}}}img', binaryItemIDRef='image1')
        caption = E.SubElement(picture, f'{{{report.HP}}}caption')
        p = self.add_para('   · 사진 캡션', wrong, '\uf0a0')
        self.pack.root.remove(p)
        caption.append(p)
        self.pack.data['BinData/image1.png'] = b'original-image-payload'
        before = E.tostring(host, method='c14n')
        self.add_para('   - 일반 본문', wrong)
        result = self.review()
        self.assertEqual(E.tostring(result.paras[-2], method='c14n'), before)
        self.assertEqual(result.data['BinData/image1.png'], self.pack.data['BinData/image1.png'])
        body = result.paras[-1]
        self.assertEqual(report.text(body), '- 일반 본문')
        self.assertEqual(body.get('styleIDRef'), self.styles['-'].get('id'))
        self.assertNotEqual(layout(result.shapes[body.get('paraPrIDRef')]), layout(self.pack.shapes[wrong]))

    def test_grouped_image_textbox_and_empty_object_host_are_preserved(self):
        wrong = self.add_shape()
        for value in ('', '   5. 이미지가 붙은 번호 줄'):
            host = self.add_para(value, wrong)
            container = E.SubElement(host.find('hp:run', report.NS), f'{{{report.HP}}}container')
            E.SubElement(container, f'{{{report.HP}}}pic', id='grouped_picture')
            rect = E.SubElement(container, f'{{{report.HP}}}rect')
            textbox = E.SubElement(rect, f'{{{report.HP}}}drawText')
            p = self.add_para('   - 그림 안 글상자', wrong, '-')
            self.pack.root.remove(p)
            textbox.append(p)
        before = [E.tostring(p, method='c14n') for p in list(self.pack.root)[-2:]]
        result = self.review()
        self.assertEqual([E.tostring(p, method='c14n') for p in result.paras[-2:]], before)

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
