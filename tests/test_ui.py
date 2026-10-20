"""Windows GUI 회귀 검사: py -m unittest discover -s tests -v"""
from pathlib import Path
from types import SimpleNamespace
import datetime as dt
import importlib.util
import sys
import tempfile
import unittest


@unittest.skipUnless(sys.platform == 'win32', 'Windows GUI에서 실행하는 검사입니다.')
class ReportUiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        root = Path(__file__).resolve().parents[1]
        spec = importlib.util.spec_from_file_location('weekly_merger_ui_test', root/'업무보고_자동취합.py')
        cls.module = importlib.util.module_from_spec(spec)
        sys.modules[spec.name] = cls.module
        spec.loader.exec_module(cls.module)
        cls.app = cls.module.App()
        scale = cls.app.ui_scale
        # 일반 사용 중인 프로그램 창을 가리지 않고 위젯 배치를 검사한다.
        cls.app.geometry(f'{round(1180*scale)}x{round(850*scale)}+{cls.app.winfo_screenwidth()+100}+0')
        cls.app.update()
        cls.errors = []
        cls.app.report_callback_exception = lambda *args: cls.errors.append(args)

    @classmethod
    def tearDownClass(cls):
        cls.app.destroy()

    def setUp(self):
        self.errors.clear()
        self.app.busy = False
        self.app.drag_start.clear()
        self.app.close_calendar()
        for kind in self.module.KINDS:
            self.app.items[kind] = [self.module.WorkItem(
                f'{kind}_{i}', kind, f'업무 제목 {i}', Path(f'원본자료_{i}.hwp'), Path('cache.hwpx'), i, i+1
            ) for i in range(3)]
            self.app.refresh_items(kind)
            self.app.item_lists[kind].xview_moveto(0)
            self.app.item_lists[kind].yview_moveto(0)
        self.app.tabs.select(0)
        self.app.date.set('2026.11.01.')
        self.app.output.set('')
        self.app.update()

    def tearDown(self):
        self.app.close_calendar()
        self.assertEqual(self.errors, [])

    def view(self, kind):
        self.app.tabs.select(self.module.KINDS.index(kind))
        self.app.update()
        return self.app.item_lists[kind]

    def pointer(self, listing, index, column='title'):
        row = listing.get_children()[index]
        box = listing.bbox(row, column)
        self.assertTrue(box)
        x, y, width, height = box
        return SimpleNamespace(x=min(x+8, listing.winfo_width()-8), y=y+height//2)

    def keys(self, kind):
        return [item.key for item in self.app.items[kind]]

    def test_title_and_filename_are_separate_columns(self):
        for kind in self.module.KINDS:
            with self.subTest(kind=kind):
                listing = self.view(kind)
                self.assertEqual(tuple(listing['columns']), ('title', 'filename'))
                self.assertEqual(listing.heading('title', 'text'), '업무 제목')
                self.assertEqual(listing.heading('filename', 'text'), '파일명')
                for n, row in enumerate(listing.get_children()):
                    self.assertEqual(listing.item(row, 'values'), (f'{n+1}. 업무 제목 {n}', f'원본자료_{n}.hwp'))

    def test_move_and_keyboard_shortcuts_in_both_tabs(self):
        for kind in self.module.KINDS:
            with self.subTest(kind=kind):
                listing = self.view(kind)
                self.app.refresh_items(kind, 0)
                self.app.move_item(kind, 1)
                self.assertEqual(self.keys(kind), [f'{kind}_{i}' for i in (1, 0, 2)])
                listing.focus_set()
                listing.event_generate('<Alt-Up>')
                self.app.update()
                self.assertEqual(self.keys(kind), [f'{kind}_{i}' for i in (0, 1, 2)])
                self.app.move_item(kind, -1)
                self.assertEqual(self.keys(kind), [f'{kind}_{i}' for i in (0, 1, 2)])

    def test_drag_from_either_column_in_both_tabs(self):
        for kind in self.module.KINDS:
            for column in ('title', 'filename'):
                with self.subTest(kind=kind, column=column):
                    listing = self.view(kind)
                    before = self.keys(kind)
                    self.app.drag_begin(kind, self.pointer(listing, 0, column))
                    target = self.pointer(listing, 2, column)
                    self.app.drag_motion(kind, target)
                    self.app.drag_end(kind, target)
                    self.assertEqual(self.keys(kind), before[1:]+before[:1])
                    first = listing.get_children()[0]
                    self.assertTrue(listing.item(first, 'values')[0].startswith('1. '))
                    self.assertEqual(listing.selection(), (before[0],))

    def test_headers_blank_space_and_outside_drop_do_not_start_or_corrupt_drag(self):
        kind = self.module.KINDS[0]
        listing = self.view(kind)
        before = self.keys(kind)
        self.app.drag_begin(kind, SimpleNamespace(x=10, y=2))
        self.assertNotIn(kind, self.app.drag_start)
        blank = SimpleNamespace(x=10, y=listing.winfo_height()-10)
        self.app.drag_begin(kind, blank)
        self.assertNotIn(kind, self.app.drag_start)
        self.app.drag_begin(kind, self.pointer(listing, 0))
        self.app.drag_end(kind, SimpleNamespace(x=-5, y=100))
        self.assertEqual(self.keys(kind), before)
        self.app.drag_begin(kind, self.pointer(listing, 0))
        self.app.drag_end(kind, blank)
        self.assertEqual(self.keys(kind), before[1:]+before[:1])

    def test_column_width_survives_reorder_and_busy_blocks_changes(self):
        kind = self.module.KINDS[0]
        listing = self.view(kind)
        width = listing.column('filename', 'width')+30
        listing.column('filename', width=width)
        self.app.refresh_items(kind, 0)
        self.app.move_item(kind, 1)
        self.assertEqual(listing.column('filename', 'width'), width)
        before = self.keys(kind)
        self.app.busy = True
        self.app.move_item(kind, 1)
        self.app.drag_begin(kind, self.pointer(listing, 0))
        self.assertEqual(self.keys(kind), before)
        self.assertNotIn(kind, self.app.drag_start)
        self.app.busy = False

    def test_calendar_starts_on_sunday_and_entry_click_opens_it(self):
        self.assertFalse(hasattr(self.app, 'calendar_button'))
        self.app.date_entry.event_generate('<Button-1>', x=5, y=5)
        self.app.update()
        picker = self.app.calendar_popup
        self.assertTrue(picker.winfo_exists())
        labels = sorted((int(w.grid_info()['column']), w.cget('text'))
                        for w in picker.grid_days.winfo_children() if int(w.grid_info()['row']) == 0)
        self.assertEqual([text for _, text in labels], list('일월화수목금토'))
        self.assertEqual(int(picker.days[1].grid_info()['column']), 0)  # 2026-11-01 is Sunday
        self.assertEqual(self.app.weekday.get(), '일요일')
        self.app.preview_label.event_generate('<Button-1>', x=2, y=2)
        self.app.update()
        self.assertIsNone(self.app.calendar_popup)

    def test_leap_date_manual_typing_weekday_and_filename(self):
        self.app.date.set('2028.02.01.')
        self.app.output.set(str(Path(tempfile.gettempdir())/'생방기술국_회의자료_280201.hwp'))
        self.app.open_calendar()
        picker = self.app.calendar_popup
        self.assertIn(29, picker.days)
        self.assertNotIn(30, picker.days)
        self.assertEqual(int(picker.days[29].grid_info()['column']), 2)  # Tuesday
        picker.days[29].invoke()
        self.app.update()
        self.assertEqual(self.app.date.get(), '2028.02.29.')
        self.assertEqual(self.app.weekday.get(), '화요일')
        self.assertTrue(self.app.output.get().endswith('생방기술국_회의자료_280229.hwp'))
        self.app.date_entry.delete(0, 'end')
        self.app.date_entry.insert(0, '2026-12-31')
        self.assertEqual(self.app.weekday.get(), '목요일')
        self.app.open_calendar()
        self.app.calendar_popup.change_month(1)
        self.assertEqual((self.app.calendar_popup.year, self.app.calendar_popup.month), (2027, 1))

    def test_outer_input_scrollbar_is_removed_but_wheel_remains_bound(self):
        shell = self.app.left_canvas.master
        self.assertEqual(shell.grid_slaves(column=1), [])
        self.assertTrue(self.app.bind('<MouseWheel>'))

    def test_scale_round_trip_updates_existing_fonts_without_padding_growth(self):
        self.app.after_cancel(self.app._display_job)
        original = self.module.display_metrics(self.app)
        font = self.app.body_font
        before = {kind: self.keys(kind) for kind in self.module.KINDS}
        baseline = {}
        try:
            for dpi in (96, 168, 96, 168):
                info = dict(original, dpi=dpi)
                self.app.apply_display_scale(info)
                self.app.update_idletasks()
                values = (font.metrics('linespace'), font.measure('업무 제목'),
                          str(self.app.content_box.cget('padding')),
                          int(self.module.ttk.Style(self.app).lookup('Work.Treeview', 'rowheight')))
                self.assertIs(self.app.body_font, font)
                self.assertEqual(self.app.ui_scale, dpi/96)
                self.assertEqual(values[-1], round(28*dpi/96))
                self.assertEqual({kind: self.keys(kind) for kind in self.module.KINDS}, before)
                if dpi in baseline:
                    self.assertEqual(values, baseline[dpi])
                baseline[dpi] = values
            self.assertGreater(baseline[168][0], baseline[96][0])
        finally:
            self.app.apply_display_scale(original)
            self.app._display_job = self.app.after(200, self.app.watch_display)


if __name__ == '__main__':
    unittest.main()
