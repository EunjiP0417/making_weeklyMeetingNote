"""실제 파일을 사용해 결과 교체와 실패 시 기존 파일 보존을 검사한다."""
import importlib.util
from pathlib import Path
import sys
import tempfile
import unittest

SPEC = importlib.util.spec_from_file_location(
    'weekly_save_test', Path(__file__).resolve().parents[1]/'업무보고_자동취합.py')
report = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = report
SPEC.loader.exec_module(report)


class FakeHwp:
    def __init__(self, save_ok=True, open_ok=True, on_validate=None):
        self.save_ok = save_ok
        self.open_ok = open_ok
        self.on_validate = on_validate
        self.opened = []

    def save_as(self, path, format):
        Path(path).write_bytes(b'new report')
        return self.save_ok

    def open(self, path):
        self.opened.append(Path(path))
        if len(self.opened) == 1 and self.on_validate:
            self.on_validate()
        return self.open_ok

    def clear(self, option):
        pass


class ResultSavingTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.output = self.folder/'회의자료.hwp'

    def test_overwrite_replaces_only_after_new_report_validates(self):
        self.output.write_bytes(b'old report')
        hwp = FakeHwp(on_validate=lambda: self.assertEqual(self.output.read_bytes(), b'old report'))
        report.save_result(hwp, self.output, overwrite=True)
        self.assertEqual(self.output.read_bytes(), b'new report')
        self.assertEqual(hwp.opened[-1], self.output)
        self.assertEqual(list(self.folder.iterdir()), [self.output])

    def test_save_or_validation_failure_keeps_existing_result(self):
        for save_ok, open_ok in ((False, True), (True, False)):
            self.output.write_bytes(b'old report')
            with self.assertRaises(RuntimeError):
                report.save_result(FakeHwp(save_ok, open_ok), self.output, overwrite=True)
            self.assertEqual(self.output.read_bytes(), b'old report')
            self.assertEqual(list(self.folder.iterdir()), [self.output])

    def test_new_filename_keeps_other_results(self):
        other = self.folder/'기존결과.hwp'
        other.write_bytes(b'old report')
        report.save_result(FakeHwp(), self.output)
        self.assertEqual(other.read_bytes(), b'old report')
        self.assertEqual(self.output.read_bytes(), b'new report')

    @unittest.skipUnless(sys.platform == 'win32', 'Windows의 기존 대상 rename 보호 검사')
    def test_unapproved_file_created_during_merge_is_not_overwritten(self):
        hwp = FakeHwp(on_validate=lambda: self.output.write_bytes(b'created meanwhile'))
        with self.assertRaisesRegex(RuntimeError, '같은 이름'):
            report.save_result(hwp, self.output)
        self.assertEqual(self.output.read_bytes(), b'created meanwhile')

    @unittest.skipUnless(sys.platform == 'win32', 'Windows 파일 잠금 검사')
    def test_locked_result_reports_how_to_retry_and_keeps_existing_file(self):
        import win32con
        import win32file
        self.output.write_bytes(b'old report')
        handle = win32file.CreateFile(str(self.output), win32con.GENERIC_READ,
                                     win32con.FILE_SHARE_READ, None, win32con.OPEN_EXISTING, 0, None)
        try:
            with self.assertRaisesRegex(RuntimeError, '한/글에서 닫고'):
                report.save_result(FakeHwp(), self.output, overwrite=True)
            self.assertEqual(self.output.read_bytes(), b'old report')
        finally:
            handle.Close()


if __name__ == '__main__':
    unittest.main()
