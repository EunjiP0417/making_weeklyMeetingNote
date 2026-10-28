"""실제 회의자료 없이 부서 경계와 본문·표 보존을 검사한다."""
import importlib.util
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import Mock

from lxml import etree as E


ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location('weekly_department_test', ROOT/'업무보고_자동취합.py')
report = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = report
SPEC.loader.exec_module(report)


class DepartmentExtractionTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.source = self.folder/'previous.hwpx'

    def fixture(self, marker=True, automatic=False, multiple=False, unnumbered=False):
        header = E.Element('header')
        for ident, kind in (('0', 'NONE'), ('1', 'NUMBER'), ('2', 'NONE')):
            shape = E.SubElement(header, 'paraPr', id=ident)
            E.SubElement(shape, 'heading', type=kind, level='0')
        root = E.Element('section', nsmap={'hp': report.HP})

        def para(value, title=False, table=None):
            p = E.SubElement(root, f'{{{report.HP}}}p',
                             paraPrIDRef=('1' if automatic else '2') if title else '0', styleIDRef='0')
            run = E.SubElement(p, f'{{{report.HP}}}run', charPrIDRef='1' if title else '0')
            E.SubElement(run, f'{{{report.HP}}}t').text = value
            if table:
                tbl = E.SubElement(run, f'{{{report.HP}}}tbl')
                inner = E.SubElement(tbl, f'{{{report.HP}}}p')
                inner_run = E.SubElement(inner, f'{{{report.HP}}}run')
                E.SubElement(inner_run, f'{{{report.HP}}}t').text = table

        para('<추진실적>')
        para('1. 견학 및 기타 제작사항', title=True)  # 지난주 실적의 동명 제목은 무시
        para('지난주 실적 본문')
        para('<추진계획>')
        titles = ['첫 업무', '견학 및 기타 제작사항' if marker else '마지막 1부서 업무',
                  '첫 2부서 업무', '견학 및 기타 제작사항' if multiple else '마지막 업무']
        for n, title in enumerate(titles):
            para(title if automatic or unnumbered else f'{n+7}. {title}', title=True)
            para(f'본문 {n}')
            para('', table=f'표 내용 {n} / 99. 표 안 가짜 업무')
        report.write_package({'Contents/section0.xml': E.tostring(root),
                              'Contents/header.xml': E.tostring(header)}, self.source)
        return titles

    def extract(self, department, selector=None):
        path = report.extract_previous_plans(self.source, self.folder, lambda s: None,
                                             department, selector)
        pack = report.Package(path)
        return pack, [item.title for item in pack.items(self.source)['추진계획']]

    def test_automatic_boundary_includes_marker_body_and_table_only_in_department1(self):
        for automatic in (False, True):
            with self.subTest(automatic=automatic):
                titles = self.fixture(automatic=automatic)
                selector = Mock(side_effect=AssertionError('자동 경계에서는 팝업 불필요'))
                for department, expected in zip(report.DEPARTMENTS, (titles[:2], titles[2:])):
                    pack, actual = self.extract(department, selector)
                    self.assertEqual(actual, expected)
                    content = ''.join(pack.root.itertext())
                    self.assertNotIn('지난주 실적 본문', content)
                    for n in range(4):
                        included = n < 2 if department == '생방기술1' else n >= 2
                        self.assertEqual(f'본문 {n}' in content, included)
                        self.assertEqual(f'표 내용 {n}' in content, included)
                selector.assert_not_called()

    def test_missing_marker_lists_only_plan_titles_and_selected_item_is_inclusive(self):
        titles = self.fixture(marker=False)
        for department, index, expected in (('생방기술1', 1, titles[:2]), ('생방기술2', 2, titles[2:])):
            with self.subTest(department=department):
                selector = Mock(return_value=index)
                _, actual = self.extract(department, selector)
                self.assertEqual(actual, expected)
                selector.assert_called_once_with(department, titles)

    def test_first_last_and_duplicate_title_selection_uses_position(self):
        titles = self.fixture(marker=False)
        for department in report.DEPARTMENTS:
            for index in (0, 3):
                with self.subTest(department=department, index=index):
                    _, actual = self.extract(department, lambda *_: index)
                    self.assertEqual(actual, titles[:index+1] if department == '생방기술1' else titles[index:])
        data = self.source.read_bytes()
        pack = report.Package(self.source)
        matching = [p for p in pack.paras if report.text(p) == '10. 마지막 업무'][0]
        matching.find('hp:run/hp:t', report.NS).text = '10. 첫 2부서 업무'
        pack.data[pack.section_name] = E.tostring(pack.root)
        report.write_package(pack.data, self.source)
        _, actual = self.extract('생방기술2', lambda *_: 3)
        self.assertEqual(actual, ['첫 2부서 업무'])
        self.assertNotEqual(self.source.read_bytes(), data)

    def test_unnumbered_titles_keep_original_order(self):
        titles = self.fixture(marker=False, unnumbered=True)
        selector = Mock(return_value=2)
        _, actual = self.extract('생방기술2', selector)
        self.assertEqual(actual, titles[2:])
        selector.assert_called_once_with('생방기술2', titles)

    def test_cancel_does_not_write_extracted_document(self):
        self.fixture(marker=False)
        with self.assertRaises(report.LoadCancelled):
            self.extract('생방기술1', lambda *_: None)
        self.assertFalse((self.folder/'previous_department1_plans.hwpx').exists())

    def test_marker_as_last_item_keeps_department1_and_rejects_empty_department2(self):
        titles = self.fixture()
        pack = report.Package(self.source)
        first_removed = next(i for i, p in enumerate(pack.paras)
                             if report.text(p) == '9. 첫 2부서 업무')
        for p in pack.paras[first_removed:]:
            pack.root.remove(p)
        pack.data[pack.section_name] = E.tostring(pack.root)
        report.write_package(pack.data, self.source)
        _, actual = self.extract('생방기술1')
        self.assertEqual(actual, titles[:2])
        with self.assertRaisesRegex(ValueError, '생방기술2 업무 제목'):
            self.extract('생방기술2')

    def test_missing_selector_invalid_selection_and_ambiguous_marker_stop(self):
        self.fixture(marker=False)
        with self.assertRaisesRegex(ValueError, '마지막 추진실적'):
            self.extract('생방기술1')
        for index in (-1, 4, '1', True):
            with self.subTest(index=index), self.assertRaises(ValueError):
                self.extract('생방기술2', lambda *_: index)
        self.fixture(multiple=True)
        with self.assertRaisesRegex(ValueError, '여러 개'):
            self.extract('생방기술1')

    def test_template_department_and_date_are_updated_without_changing_original(self):
        original = report.TEMPLATE_PATH.read_bytes()
        for department in report.DEPARTMENTS:
            output = self.folder/f'{department}.hwpx'
            report.template_skeleton(report.TEMPLATE_PATH, output, '2026.10.28.', '수', department)
            content = ''.join(report.Package(output).root.itertext())
            self.assertIn(department, content)
            self.assertNotIn('생방기술2' if department == '생방기술1' else '생방기술1', content)
            self.assertIn('2026.10.28.', content)
        self.assertEqual(report.TEMPLATE_PATH.read_bytes(), original)


if __name__ == '__main__':
    unittest.main()
