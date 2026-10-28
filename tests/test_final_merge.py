"""부서 최종본의 항목별 연결·연속 번호·저장 이름을 검사한다."""
import importlib.util
from pathlib import Path
import sys
import tempfile
import unittest

from lxml import etree as E

ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location('weekly_final_test', ROOT/'업무보고_자동취합.py')
report = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = report
SPEC.loader.exec_module(report)


class FinalMergeTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)

    def fixture(self, department, kinds=report.KINDS):
        pack = report.Package(report.TEMPLATE_PATH)
        prototype = pack.paras[0]
        styles = {s.get('name'): s for s in pack.header.iter() if report.local(s) == 'style'}
        for child in list(pack.root):
            pack.root.remove(child)
        for kind in kinds:
            pack.root.append(pack.simple(prototype, f'<{kind}>'))
            for n in (1, 2):
                p = pack.simple(prototype, f'{n}. {department} {kind} 업무 {n}')
                p.set('paraPrIDRef', styles['1,2,3'].get('paraPrIDRef'))
                pack.root.append(p)
                p = pack.simple(prototype, '- 업무 본문')
                p.set('paraPrIDRef', styles['-'].get('paraPrIDRef'))
                pack.root.append(p)
        pack.data[pack.section_name] = E.tostring(pack.root, encoding='UTF-8', xml_declaration=True)
        pack.data['Contents/header.xml'] = E.tostring(pack.header, encoding='UTF-8', xml_declaration=True)
        path = self.folder/f'{department}.hwpx'
        report.write_package(pack.data, path)
        return path

    def test_each_kind_is_department1_then_department2_and_renumbered_independently(self):
        paths = [self.fixture(d) for d in report.DEPARTMENTS]
        ordered = report.collect_final_items([(p, p) for p in paths], lambda _: None)
        for kind in report.KINDS:
            self.assertEqual([i.source for i in ordered[kind]], [paths[0]]*2+[paths[1]]*2)
            self.assertEqual([i.title for i in ordered[kind]],
                             [f'{d} {kind} 업무 {n}' for d in report.DEPARTMENTS for n in (1, 2)])
        pack = report.Package(report.TEMPLATE_PATH)
        prototype = pack.paras[0]
        for child in list(pack.root):
            pack.root.remove(child)
        for kind in report.KINDS:
            pack.root.append(pack.simple(prototype, f'<{kind}>'))
            for item in ordered[kind]:
                pack.root.append(pack.simple(prototype, f'REPORT_ITEM_{item.key} 2. {item.title}'))
        pack.data[pack.section_name] = E.tostring(pack.root, encoding='UTF-8', xml_declaration=True)
        pack.data['Contents/header.xml'] = E.tostring(pack.header, encoding='UTF-8', xml_declaration=True)
        source = self.folder/'combined.hwpx'
        report.write_package(pack.data, source)
        target = self.folder/'numbered.hwpx'
        report.body_spacing(source, target, {k: 0 for k in report.KINDS}, ordered)
        actual = report.Package(target)
        lines = [report.text(p) for p in actual.paras if report.text(p)]
        self.assertEqual(lines, [value for kind in report.KINDS for value in
                         [f'<{kind}>', *[f'{n}. {item.title}' for n, item in enumerate(ordered[kind], 1)]]])

    def test_missing_kind_and_missing_department_are_rejected(self):
        first = self.fixture('생방기술1')
        second = self.fixture('생방기술2', ('추진계획',))
        with self.assertRaisesRegex(ValueError, '생방기술2'):
            report.collect_final_items([(first, first), (second, second)], lambda _: None)
        with self.assertRaises(ValueError):
            report.collect_final_items([(first, first)], lambda _: None)
        with self.assertRaises(ValueError):
            report.prepare_final([first, first], self.folder, lambda _: None)

    def test_filenames_and_national_template_header(self):
        for department in report.REPORT_DEPARTMENTS:
            self.assertEqual(report.output_name('2026.10.24.', department=department),
                             f'{department}_회의자료_261024.hwp')
        output = self.folder/'national.hwpx'
        report.template_skeleton(report.TEMPLATE_PATH, output, '2026.10.24.', '토', '생방기술국')
        text = ''.join(report.Package(output).root.itertext())
        self.assertIn('생방기술국', text)
        self.assertNotIn('생방기술2', text)
        self.assertIn('2026.10.24.', text)


if __name__ == '__main__':
    unittest.main()
