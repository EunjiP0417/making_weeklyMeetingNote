"""Windows GUI 회귀 검사: py -m unittest discover -s tests -v"""
from pathlib import Path
from types import SimpleNamespace
import datetime as dt
import importlib.util
import sys
import tempfile
import time
import unittest
from unittest.mock import patch


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
        self.app.select_department('생방기술2')
        self.app.set_busy(False)
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

    def test_department_buttons_order_highlight_invalidation_and_busy_guard(self):
        buttons = self.app.department_buttons
        self.assertEqual(list(buttons), ['생방기술1', '생방기술2', '생방기술국'])
        self.assertLess(buttons['생방기술1'].winfo_x(), buttons['생방기술2'].winfo_x())
        self.assertEqual(buttons['생방기술1'].winfo_width(), buttons['생방기술2'].winfo_width())
        self.assertTrue(buttons['생방기술2'].instate(['selected']))
        style = self.module.ttk.Style(self.app)
        self.assertEqual(style.lookup('Department.TButton', 'background', ('selected',)), '#171717')
        self.assertEqual(style.lookup('Department.TButton', 'foreground', ('selected',)), 'white')
        padding = tuple(map(int, self.app.tk.splitlist(style.lookup('Department.TButton', 'padding'))))
        self.assertEqual(padding, (round(8*self.app.ui_scale), round(6*self.app.ui_scale)))
        self.assertEqual(self.app.start_button.cget('text'), '생방기술2 최종 취합  →')
        self.app.loaded_department = '생방기술2'
        buttons['생방기술1'].invoke()
        self.assertEqual(self.app.department, '생방기술1')
        self.assertTrue(buttons['생방기술1'].instate(['selected']))
        self.assertTrue(buttons['생방기술2'].instate(['!selected']))
        self.assertIsNone(self.app.loaded_department)
        self.assertTrue(all(not rows for rows in self.app.items.values()))
        self.app.set_busy(True)
        self.assertTrue(all(b.instate(['disabled']) for b in buttons.values()))
        buttons['생방기술2'].invoke()
        self.app.select_department('생방기술2')
        self.assertEqual(self.app.department, '생방기술1')
        self.app.set_busy(False)
        self.assertTrue(all(b.instate(['!disabled']) for b in buttons.values()))
        self.assertTrue(buttons['생방기술1'].instate(['selected']))

    def test_boundary_dialog_prompts_titles_confirmation_and_cancellation(self):
        for department, prompt in (
                ('생방기술1', '생방기술1의 마지막 추진실적을 선택하시오.'),
                ('생방기술2', '생방기술2의 첫 번째 추진실적을 선택하시오.')):
            with self.subTest(department=department):
                titles = ['같은 제목', '긴 업무 제목 '*12, '같은 제목']
                dialog = self.module.BoundaryDialog(self.app, department, titles)
                self.app.update()
                self.assertEqual(dialog.prompt.cget('text'), prompt)
                self.assertEqual(list(dialog.listbox.get(0, 'end')), titles)
                self.assertEqual(dialog.listbox.curselection(), ())
                self.assertTrue(dialog.confirm_button.instate(['disabled']))
                dialog.confirm()
                self.assertTrue(dialog.winfo_exists())
                dialog.listbox.selection_set(2)
                dialog.listbox.event_generate('<<ListboxSelect>>')
                self.assertTrue(dialog.confirm_button.instate(['!disabled']))
                dialog.confirm_button.invoke()
                self.assertEqual(dialog.result, 2)
                self.assertFalse(dialog.winfo_exists())
                dialog = self.module.BoundaryDialog(self.app, department, titles)
                dialog.destroy()
                self.assertIsNone(dialog.result)

    def test_background_load_boundary_response_and_cancel_restore_buttons(self):
        self.app.cache_dir = None
        for department, cancel in ((d, c) for d in self.module.DEPARTMENTS for c in (False, True)):
            with self.subTest(department=department, cancel=cancel), tempfile.TemporaryDirectory() as folder:
                self.app.select_department(department)
                previous = Path(folder)/'previous.hwpx'
                previous.touch()
                self.app.previous_path = previous
                self.app.paths = [Path(folder)/'current.hwpx']
                original_items = {k: list(v) for k, v in self.app.items.items()}
                if not all(original_items.values()):
                    original_items = {k: [self.module.WorkItem(k, k, '업무', previous,
                        previous, 1, 2)] for k in self.module.KINDS}
                calls = []

                def fake_prepare(paths, cache, log, previous_path, department, select_boundary):
                    calls.append((paths, previous_path, department))
                    chosen = select_boundary(department, ['첫 업무', '경계 업무', '끝 업무'])
                    if chosen is None:
                        raise self.module.LoadCancelled()
                    self.assertEqual(chosen, 1)
                    return original_items

                def answer_dialog():
                    for child in self.app.winfo_children():
                        if isinstance(child, self.module.BoundaryDialog):
                            if cancel:
                                child.destroy()
                            else:
                                child.listbox.selection_set(1)
                                child.confirm()
                            return
                    self.app.after(10, answer_dialog)

                with patch.object(self.module, 'prepare', side_effect=fake_prepare):
                    self.app.after(10, answer_dialog)
                    self.app.load_titles()
                    self.assertTrue(self.app.busy)
                    deadline = time.monotonic()+5
                    while self.app.busy and time.monotonic() < deadline:
                        self.app.update()
                        time.sleep(.01)
                    self.assertFalse(self.app.busy)
                self.assertEqual(calls, [(self.app.paths, previous, self.app.department)])
                self.assertTrue(all(b.instate(['!disabled']) for b in self.app.department_buttons.values()))
                if cancel:
                    self.assertIn('취소', self.app.preview_state.get())
                    self.assertIsNone(self.app.loaded_department)
                    self.assertTrue(all(not items for items in self.app.items.values()))
                else:
                    self.assertEqual(self.app.loaded_department, self.app.department)
                    self.assertEqual(self.app.items, original_items)
                if self.app.cache_dir is not None:
                    self.app.cache_dir.cleanup()
                    self.app.cache_dir = None
        self.app.paths = []
        self.app.previous_path = None

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
        self.app.output.set(str(Path(tempfile.gettempdir())/'생방기술2_회의자료_280201.hwp'))
        self.app.open_calendar()
        picker = self.app.calendar_popup
        self.assertIn(29, picker.days)
        self.assertNotIn(30, picker.days)
        self.assertEqual(int(picker.days[29].grid_info()['column']), 2)  # Tuesday
        picker.days[29].invoke()
        self.app.update()
        self.assertEqual(self.app.date.get(), '2028.02.29.')
        self.assertEqual(self.app.weekday.get(), '화요일')
        self.assertTrue(self.app.output.get().endswith('생방기술2_회의자료_280229.hwp'))

    def test_final_mode_is_separate_page_and_filenames_follow_mode_and_date(self):
        self.app.output.set(str(Path(tempfile.gettempdir())/'old.hwp'))
        for mode, department in (('생방기술1', '생방기술1'), ('생방기술2', '생방기술2'),
                                 (self.module.COMBINED_MODE, '생방기술국'), ('생방기술1', '생방기술1')):
            with self.subTest(mode=mode):
                self.app.select_department(mode)
                self.app.update()
                self.assertEqual(self.app.final_inputs.winfo_ismapped(), mode == self.module.COMBINED_MODE)
                self.assertEqual(self.app.department_inputs.winfo_ismapped(), mode != self.module.COMBINED_MODE)
                self.assertTrue(self.app.output.get().endswith(f'{department}_회의자료_261101.hwp'))
                self.assertIsNone(self.app.loaded_department)
                self.assertTrue(all(b.instate(['!disabled']) for b in self.app.order_buttons))
                self.assertEqual(self.app.start_button.cget('text'), f'{department} 최종 취합  →')

    def test_final_mode_rejects_invalid_files_and_saves_user_reordered_items(self):
        self.app.select_department(self.module.COMBINED_MODE)
        self.app.final_paths = {d: None for d in self.module.DEPARTMENTS}
        with patch.object(self.module.messagebox, 'showerror') as error:
            self.app.load_titles()
            error.assert_called_once()
        with tempfile.TemporaryDirectory() as folder:
            paths = [Path(folder)/f'department_{n}.hwpx' for n in (1, 2)]
            for path in paths:
                path.touch()
            self.app.final_paths = dict(zip(self.module.DEPARTMENTS, [paths[0], paths[0]]))
            with patch.object(self.module.messagebox, 'showerror') as error:
                self.app.load_titles()
                error.assert_called_once()
            self.app.final_paths = dict(zip(self.module.DEPARTMENTS, paths))
            expected = {k: [self.module.WorkItem(f'{k}_{n}', k, f'생방{n} 업무', paths[n-1],
                        paths[n-1], 0, 1) for n in (1, 2)] for k in self.module.KINDS}
            with patch.object(self.module, 'prepare_final', return_value=expected) as prepare_final, \
                    patch.object(self.module, 'prepare', side_effect=AssertionError('최종 모드에서 이월 금지')):
                self.app.load_titles()
                deadline = time.monotonic()+5
                while self.app.busy and time.monotonic() < deadline:
                    self.app.update()
                    time.sleep(.01)
                self.assertFalse(self.app.busy)
                self.assertEqual(prepare_final.call_args.args[0], paths)
            self.assertEqual(self.app.items, expected)
            self.assertEqual(self.app.loaded_paths, paths)
            before = {k: self.keys(k) for k in self.module.KINDS}
            for kind in self.module.KINDS:
                listing = self.view(kind)
                self.app.refresh_items(kind, 0)
                self.app.move_item(kind, 1)
                self.assertEqual(self.keys(kind), before[kind][::-1])
                self.app.drag_begin(kind, self.pointer(listing, 0))
                self.app.drag_end(kind, self.pointer(listing, 1))
                self.assertEqual(self.keys(kind), before[kind])
                self.app.refresh_items(kind, 0)
                listing.focus_set()
                listing.event_generate('<Alt-Down>')
                self.app.update()
                self.assertEqual(self.keys(kind), before[kind][::-1])
                self.assertNotIn(kind, self.app.drag_start)
            wanted = {k: list(self.app.items[k]) for k in self.module.KINDS}
            self.app.output.set(str(Path(folder)/'생방기술국_회의자료_261101.hwp'))
            with patch.object(self.module, 'merge') as merge, \
                    patch.object(self.module, 'CompletionDialog', return_value=None):
                self.app.start()
                deadline = time.monotonic()+5
                while self.app.busy and time.monotonic() < deadline:
                    self.app.update()
                    time.sleep(.01)
                self.assertFalse(self.app.busy)
                merge.assert_called_once()
                self.assertEqual(merge.call_args.args[0], wanted)
                self.assertEqual(merge.call_args.args[1].name, '생방기술국_회의자료_261101.hwp')
                self.assertEqual(merge.call_args.kwargs['department'], '생방기술국')
            self.app.cache_dir.cleanup()
            self.app.cache_dir = None
        self.app.final_paths = {d: None for d in self.module.DEPARTMENTS}
        self.app.date_entry.delete(0, 'end')
        self.app.date_entry.insert(0, '2026-12-31')
        self.assertEqual(self.app.weekday.get(), '목요일')
        self.app.open_calendar()
        self.app.calendar_popup.change_month(1)
        self.assertEqual((self.app.calendar_popup.year, self.app.calendar_popup.month), (2027, 1))

    def test_output_conflict_dialog_buttons_cancel_and_default_rename(self):
        for action in ('overwrite', 'rename', 'cancel', 'enter'):
            dialog = self.module.OutputConflictDialog(self.app, Path('검토 폴더/회의자료.hwp'))
            self.app.update()
            self.assertIn('회의자료.hwp', dialog.path_label.cget('text'))
            if action == 'enter':
                dialog.event_generate('<Return>')
            else:
                getattr(dialog, action+'_button').invoke()
            self.assertFalse(dialog.winfo_exists())
            self.assertEqual(dialog.result, {'cancel':None, 'enter':'rename'}.get(action, action))
        dialog = self.module.OutputConflictDialog(self.app, Path('회의자료.hwp'))
        self.app.update()
        dialog.event_generate('<Escape>')
        self.assertIsNone(dialog.result)
        self.assertFalse(dialog.winfo_exists())

    def test_output_resolution_overwrite_rename_recheck_and_cancel(self):
        with tempfile.TemporaryDirectory() as folder:
            original = Path(folder)/'회의자료.hwp'
            original.write_bytes(b'previous')
            renamed = Path(folder)/'회의자료_수정.hwp'
            existing = Path(folder)/'다른결과.hwp'
            existing.write_bytes(b'other')
            with patch.object(self.app, 'wait_window'), \
                    patch.object(self.module, 'OutputConflictDialog', return_value=SimpleNamespace(result='overwrite')):
                self.assertEqual(self.app.resolve_output(original, []), (original.resolve(), True))
            with patch.object(self.app, 'wait_window'), \
                    patch.object(self.module, 'OutputConflictDialog', side_effect=[SimpleNamespace(result='rename'), SimpleNamespace(result='overwrite')]) as dialog, \
                    patch.object(self.module.filedialog, 'asksaveasfilename', return_value=str(existing)):
                self.assertEqual(self.app.resolve_output(original, []), (existing.resolve(), True))
                self.assertEqual(dialog.call_count, 2)
            with patch.object(self.app, 'wait_window'), \
                    patch.object(self.module, 'OutputConflictDialog', return_value=SimpleNamespace(result='rename')), \
                    patch.object(self.module.filedialog, 'asksaveasfilename', return_value=str(renamed)):
                self.assertEqual(self.app.resolve_output(original, []), (renamed.resolve(), False))
            for action, name in ((None, None), ('rename', '')):
                with patch.object(self.app, 'wait_window'), \
                        patch.object(self.module, 'OutputConflictDialog', return_value=SimpleNamespace(result=action)), \
                        patch.object(self.module.filedialog, 'asksaveasfilename', return_value=name):
                    self.assertIsNone(self.app.resolve_output(original, []))
            self.assertEqual(original.read_bytes(), b'previous')
            self.assertEqual(existing.read_bytes(), b'other')

    def test_output_resolution_protects_sources_and_checks_renamed_paths(self):
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder)/'원본.hwp'
            source.touch()
            existing = Path(folder)/'결과.hwp'
            existing.touch()
            with patch.object(self.module, 'OutputConflictDialog') as dialog:
                with self.assertRaisesRegex(ValueError, '원본'):
                    self.app.resolve_output(source, [source])
                dialog.assert_not_called()
            for target, error in ((source, '원본'), (Path(folder)/'잘못.txt', '확장자')):
                with patch.object(self.app, 'wait_window'), \
                        patch.object(self.module, 'OutputConflictDialog', return_value=SimpleNamespace(result='rename')), \
                        patch.object(self.module.filedialog, 'asksaveasfilename', return_value=str(target)):
                    with self.assertRaisesRegex(ValueError, error):
                        self.app.resolve_output(existing, [source])

    def test_start_keeps_custom_filename_and_passes_overwrite_for_all_modes(self):
        with tempfile.TemporaryDirectory() as folder:
            for mode in self.module.MODES:
                self.app.select_department(mode)
                paths = [Path(folder)/f'input_{n}.hwp' for n in (1, 2)]
                self.app.paths = paths
                self.app.final_paths = dict(zip(self.module.DEPARTMENTS, paths))
                self.app.previous_path = Path(folder)/'previous.hwp'
                inputs, previous = self.app.current_inputs()
                self.app.loaded_paths = inputs
                self.app.loaded_previous = previous
                self.app.loaded_department = mode
                self.app.cache_dir = SimpleNamespace(cleanup=lambda: None)
                self.app.items = {k: [self.module.WorkItem(k, k, '업무', paths[0], paths[0], 0, 1)]
                                  for k in self.module.KINDS}
                target = Path(folder)/f'{mode}_검토용.hwp'
                self.app.output.set(str(target))
                with patch.object(self.app, 'resolve_output', return_value=None), \
                        patch.object(self.module, 'merge') as merge:
                    self.app.start()
                    merge.assert_not_called()
                    self.assertFalse(self.app.busy)
                with patch.object(self.app, 'resolve_output', return_value=(target, True)), \
                        patch.object(self.module, 'merge') as merge, \
                        patch.object(self.module, 'CompletionDialog', return_value=None):
                    self.app.start()
                    deadline = time.monotonic()+5
                    while self.app.busy and time.monotonic() < deadline:
                        self.app.update()
                        time.sleep(.01)
                    self.assertFalse(self.app.busy)
                    self.assertEqual(merge.call_args.args[1], target)
                    self.assertTrue(merge.call_args.kwargs['overwrite'])
                    self.assertEqual(self.app.output.get(), str(target))
            self.app.cache_dir = None

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
