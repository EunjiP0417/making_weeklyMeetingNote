"""실제 Windows 창으로 부트로더 숨김 처리의 대상 범위를 검사한다."""
import importlib.util
import os
from pathlib import Path
import sys
import time
import unittest
from unittest.mock import patch


@unittest.skipUnless(sys.platform == 'win32', 'Windows 창 검사입니다.')
class BootloaderWindowTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        import pywintypes
        import win32gui
        import win32process
        cls.gui = win32gui
        spec = importlib.util.spec_from_file_location(
            'weekly_bootloader_test', Path(__file__).resolve().parents[1]/'업무보고_자동취합.py')
        cls.module = importlib.util.module_from_spec(spec)
        sys.modules[spec.name] = cls.module
        spec.loader.exec_module(cls.module)
        window_class = win32gui.WNDCLASS()
        window_class.lpszClassName = 'PyInstallerOnefileHiddenWindow'
        window_class.lpfnWndProc = win32gui.DefWindowProc
        cls.atom = win32gui.RegisterClass(window_class)

    @classmethod
    def tearDownClass(cls):
        cls.gui.UnregisterClass(cls.atom, None)

    def setUp(self):
        self.windows = []

    def tearDown(self):
        for hwnd in self.windows:
            self.gui.DestroyWindow(hwnd)

    def window(self, title='PyInstaller Onefile Hidden Window', cls=None):
        hwnd = self.gui.CreateWindowEx(
            0x08000080, cls or self.atom, title, 0x10000000,
            -30000, -30000, 20, 20, 0, 0, 0, None)
        self.windows.append(hwnd)
        self.assertTrue(self.gui.IsWindowVisible(hwnd))
        return hwnd

    def wait_for_hide(self, hwnd):
        deadline = time.monotonic()+2
        while self.gui.IsWindowVisible(hwnd) and time.monotonic() < deadline:
            self.gui.PumpWaitingMessages()
            time.sleep(.01)
        self.assertFalse(self.gui.IsWindowVisible(hwnd))
        self.assertTrue(self.gui.IsWindow(hwnd), '보조 창의 종료 처리를 유지해야 합니다.')

    def test_only_matching_parent_helper_is_hidden(self):
        helper = self.window()
        main = self.window('주간 회의자료 취합', 'STATIC')
        same_title = self.window(cls='STATIC')
        different_title = self.window('다른 보조 창')
        with patch.object(sys, 'frozen', True, create=True), \
                patch.object(self.module.os, 'getppid', return_value=os.getpid()):
            self.module.hide_onefile_helper_window()
            self.wait_for_hide(helper)
            # 실행 환경이 다시 표시해도 같은 창을 다시 숨길 수 있다.
            self.gui.ShowWindow(helper, 8)
            self.assertTrue(self.gui.IsWindowVisible(helper))
            self.module.hide_onefile_helper_window()
            self.wait_for_hide(helper)
        for hwnd in (main, same_title, different_title):
            self.assertTrue(self.gui.IsWindowVisible(hwnd))

    def test_other_process_helper_is_untouched(self):
        helper = self.window()
        with patch.object(sys, 'frozen', True, create=True):
            self.module.hide_onefile_helper_window()
        self.gui.PumpWaitingMessages()
        self.assertTrue(self.gui.IsWindowVisible(helper))

    def test_source_execution_does_not_hide_windows(self):
        helper = self.window()
        with patch.object(sys, 'frozen', False, create=True), \
                patch.object(self.module.os, 'getppid', return_value=os.getpid()):
            self.module.hide_onefile_helper_window()
        self.assertTrue(self.gui.IsWindowVisible(helper))


if __name__ == '__main__':
    unittest.main()
