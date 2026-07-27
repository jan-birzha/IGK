# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Any

from config import (
    DATA_START_ROW, DEFAULT_FILENAME_COL, DEFAULT_TOTAL_COL, EXCEL_EXTENSIONS,
    L6_L7_FIELDS, L_COLUMN, MANUAL_ROW15_COLS, MANUAL_ROW15_FIELDS, R_COLUMN,
    UI, XL_CALCULATION_MANUAL, XL_SHIFT_DOWN, XL_TEXT_FORMAT
)
from utils import (
    bind_entry_paste_shortcuts, collect_files_with_extensions_recursive,
    dialog_parent, prepare_modal_dialog, show_modal_dialog
)

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    import pythoncom
    import win32com.client as win32
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False


def _read_range_row_values(block) -> tuple:
    if block is None:
        return ()
    if isinstance(block, (tuple, list)) and block and isinstance(block[0], (tuple, list)):
        return tuple(block[0])
    if isinstance(block, (tuple, list)):
        return tuple(block)
    return (block,)


def _write_column_values(ws, start_row: int, col: int, values: list) -> None:
    if not values:
        return
    if len(values) == 1:
        ws.Cells(start_row, col).Value = values[0]
        return
    top = ws.Cells(start_row, col)
    bottom = ws.Cells(start_row + len(values) - 1, col)
    ws.Range(top, bottom).Value = tuple((v,) for v in values)


def _configure_excel_fast(excel) -> None:
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    excel.EnableEvents = False
    try:
        excel.Calculation = XL_CALCULATION_MANUAL
    except Exception:
        pass
    try:
        excel.Interactive = False
    except Exception:
        pass


def _restore_excel_defaults(excel) -> None:
    try:
        excel.ScreenUpdating = True
        excel.DisplayAlerts = True
        excel.EnableEvents = True
        excel.CutCopyMode = False
        excel.Interactive = True
    except Exception:
        pass


def name_select_folder(parent: tk.Misc | None = None) -> str | None:
    folder_path = filedialog.askdirectory(parent=dialog_parent(parent), title="Выберите папку")
    return folder_path or None


def name_list_filenames_recursive(folder_path: str) -> list[str]:
    file_names: list[str] = []
    for _current_dir, _sub_dirs, files in os.walk(folder_path):
        for file_name in files:
            file_names.append(file_name)
    return file_names


def name_select_target_file(parent: tk.Misc | None = None) -> str | None:
    file_path = filedialog.askopenfilename(
        parent=dialog_parent(parent),
        title="Выберите файл Excel для вставки данных",
        filetypes=[("Excel Files", "*.xlsx *.xlsm *.xls")],
    )
    return file_path or None


def _cell_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _find_filename_column(ws) -> int:
    block = ws.Range(ws.Cells(12, 1), ws.Cells(14, 39)).Value
    if block is None:
        return DEFAULT_FILENAME_COL
    rows = block if isinstance(block[0], (tuple, list)) else (block,)
    for header_row in rows:
        for col_idx, val in enumerate(header_row, start=1):
            if val and "название файла" in _cell_text(val).lower():
                return col_idx
    return DEFAULT_FILENAME_COL


def _find_total_column(ws, search_row: int) -> int:
    block = ws.Range(ws.Cells(search_row, 1), ws.Cells(search_row, 39)).Value
    row_values = _read_range_row_values(block)
    for col_idx, val in enumerate(row_values, start=1):
        if _cell_text(val).lower() == "итого":
            cell = ws.Cells(search_row, col_idx)
            if cell.MergeCells:
                return cell.MergeArea.Column
            return col_idx
    return DEFAULT_TOTAL_COL


def _unmerge_row_cells(ws, row: int, columns: tuple[str, ...]) -> None:
    for col_letter in columns:
        cell = ws.Range(f"{col_letter}{row}")
        if cell.MergeCells:
            cell.MergeArea.UnMerge()


def _open_workbook(excel, file_path: str):
    abs_path = os.path.normpath(os.path.abspath(file_path))
    if not os.path.isfile(abs_path):
        raise FileNotFoundError(f"Файл не найден: {abs_path}")

    normalized = abs_path.lower()
    for workbook in excel.Workbooks:
        try:
            if os.path.normpath(workbook.FullName).lower() == normalized:
                return workbook
        except Exception:
            continue

    try:
        return excel.Workbooks.Open(abs_path)
    except Exception:
        return excel.Workbooks.Open(
            Filename=abs_path,
            UpdateLinks=0,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True,
            AddToMru=False,
        )


def _is_numeric_value(value: object) -> tuple[bool, float]:
    if value is None:
        return False, 0.0
    if isinstance(value, (int, float)):
        return True, float(value)
    text = _cell_text(value).replace(",", ".")
    if not text:
        return False, 0.0
    try:
        return True, float(text)
    except ValueError:
        return False, 0.0


def _cell_target_range(cell):
    if cell.MergeCells:
        return cell.MergeArea
    return cell


def _set_text_format(ws, row: int, col: int) -> None:
    target = _cell_target_range(ws.Cells(row, col))
    target.NumberFormat = XL_TEXT_FORMAT
    try:
        target.NumberFormatLocal = XL_TEXT_FORMAT
    except Exception:
        pass


def _ensure_manual_row_text_format(ws, row: int = DATA_START_ROW) -> None:
    for col in MANUAL_ROW15_COLS:
        _set_text_format(ws, row, col)


def _set_cell_display_value(ws, row: int, col: int, value: str) -> None:
    if not value:
        return
    cell = ws.Cells(row, col)
    target = _cell_target_range(cell)
    target.NumberFormat = XL_TEXT_FORMAT
    try:
        target.NumberFormatLocal = XL_TEXT_FORMAT
    except Exception:
        pass
    target.Cells(1, 1).Value = value


def _read_manual_row_values(ws, row: int = DATA_START_ROW) -> dict[int, str]:
    block = ws.Range(ws.Cells(row, 2), ws.Cells(row, 6)).Value
    cells = _read_range_row_values(block)
    return {
        2: _cell_text(cells[0]) if len(cells) > 0 else "",
        3: _cell_text(cells[1]) if len(cells) > 1 else "",
        5: _cell_text(cells[3]) if len(cells) > 3 else "",
    }


def _read_manual_row15_openpyxl(path: str) -> dict[int, str] | None:
    if not HAS_OPENPYXL:
        return None
    ext = os.path.splitext(path)[1].lower()
    if ext not in (".xlsx", ".xlsm"):
        return None
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active

        def val(coord: str) -> str:
            return _cell_text(ws[coord].value)

        data = {2: val("B15"), 3: val("C15"), 5: val("E15")}
        wb.close()
        return data
    except Exception:
        return None


def _apply_manual_row_entries(ws, row: int, entries: dict[int, str]) -> None:
    for col, value in entries.items():
        if value:
            _set_cell_display_value(ws, row, col, value)


def _apply_l6_l7(ws, l6_value: str, l7_value: str) -> None:
    if l6_value:
        _set_cell_display_value(ws, 6, L_COLUMN, l6_value)
    if l7_value:
        _set_cell_display_value(ws, 7, L_COLUMN, l7_value)


def name_ask_manual_combined(
        parent: tk.Misc | None,
        title: str,
        fields: tuple[tuple[int, str, str], ...],
        include_l6l7: bool,
) -> dict:
    result: dict[Any, str] = {col: "" for col, _, _ in fields}
    if include_l6l7:
        result["L6"] = ""
        result["L7"] = ""

    if not fields and not include_l6l7:
        return result

    dialog = prepare_modal_dialog(parent, title)

    ttk.Label(
        dialog,
        text="Заполните нужные поля (пустое поле — оставить без изменений):",
        background=UI["bg_main"],
        foreground=UI["text"],
        font=UI["font_card_desc"],
    ).pack(padx=20, pady=(16, 12))

    entries: dict[Any, tk.Entry] = {}

    for col, cell_label, prompt in fields:
        frame = tk.Frame(dialog, bg=UI["bg_main"])
        frame.pack(fill=tk.X, padx=20, pady=4)
        tk.Label(
            frame,
            text=f"{cell_label}:",
            background=UI["bg_main"],
            foreground=UI["text"],
            font=("Segoe UI", 9, "bold"),
        ).pack(anchor="w")
        tk.Label(
            frame,
            text=prompt,
            background=UI["bg_main"],
            foreground=UI["muted"],
            font=UI["font_footer"],
            wraplength=400,
            justify=tk.LEFT
        ).pack(anchor="w")
        entry = tk.Entry(frame, width=50, exportselection=False, font=("Segoe UI", 10), bg=UI["card"], fg=UI["text"],
                         insertbackground=UI["text"], highlightthickness=1, highlightcolor=UI["accent_primary"],
                         highlightbackground=UI["border"], relief=tk.FLAT)
        entry.pack(fill=tk.X, pady=(4, 0))
        bind_entry_paste_shortcuts(dialog, entry)
        entries[col] = entry

    if include_l6l7:
        if fields:
            sep = tk.Frame(dialog, height=1, bg=UI["border"])
            sep.pack(fill=tk.X, padx=20, pady=(8, 4))
        for key, cell_label, prompt in L6_L7_FIELDS:
            frame = tk.Frame(dialog, bg=UI["bg_main"])
            frame.pack(fill=tk.X, padx=20, pady=4)
            tk.Label(
                frame,
                text=f"{cell_label}:",
                background=UI["bg_main"],
                foreground=UI["text"],
                font=("Segoe UI", 9, "bold"),
            ).pack(anchor="w")
            tk.Label(
                frame,
                text=prompt,
                background=UI["bg_main"],
                foreground=UI["muted"],
                font=UI["font_footer"],
                wraplength=400,
                justify=tk.LEFT
            ).pack(anchor="w")
            entry = tk.Entry(frame, width=50, exportselection=False, font=("Segoe UI", 10), bg=UI["card"],
                             fg=UI["text"], insertbackground=UI["text"], highlightthickness=1,
                             highlightcolor=UI["accent_primary"], highlightbackground=UI["border"], relief=tk.FLAT)
            entry.pack(fill=tk.X, pady=(4, 0))
            bind_entry_paste_shortcuts(dialog, entry)
            entries[key] = entry

    if entries:
        next(iter(entries.values())).focus_set()

    btn_frame = tk.Frame(dialog, bg=UI["bg_main"])
    btn_frame.pack(pady=(16, 16))

    def on_ok() -> None:
        for key, entry in entries.items():
            result[key] = entry.get().strip()
        dialog.destroy()

    ok_btn = tk.Button(btn_frame, text="OK", command=on_ok, bg=UI["accent_primary"], fg="#ffffff", relief=tk.FLAT,
                       width=10, font=("Segoe UI", 9, "bold"), activebackground=UI["accent_hover"],
                       activeforeground="#ffffff")
    ok_btn.pack(side=tk.LEFT, padx=6)

    dialog.bind("<Return>", lambda _e: on_ok())
    dialog.bind("<Escape>", lambda _e: on_cancel())

    show_modal_dialog(dialog)
    dialog.wait_window()
    return result


def name_process_workbook(target_path: str, file_names: list[str], parent: tk.Misc | None = None) -> None:
    if not HAS_WIN32:
        messagebox.showerror(
            "Ошибка",
            "Для работы с Excel требуется Windows, Microsoft Excel и pywin32.\nУстановите: pip install pywin32",
            parent=dialog_parent(parent),
        )
        return

    data_count = len(file_names)
    proceed = messagebox.askyesno(
        "Подтверждение",
        f"Будет вставлено {data_count} имён файлов в выбранный файл.\nПродолжить?",
        parent=dialog_parent(parent),
    )
    if not proceed:
        return

    start_total = time.perf_counter()
    timings: list[str] = []

    t0 = time.perf_counter()
    pre_read = _read_manual_row15_openpyxl(target_path)
    manual_entries: dict[int, str] = {}
    l6_value = ""
    l7_value = ""
    if pre_read is not None:
        empty_fields = tuple(
            (col, label, prompt) for col, label, prompt in MANUAL_ROW15_FIELDS if not pre_read.get(col)
        )
        entered = name_ask_manual_combined(parent, "Ручной ввод", empty_fields, True)
        manual_entries = {c: v for c, v in entered.items() if c in (2, 3, 5)}
        l6_value = entered.get("L6", "")
        l7_value = entered.get("L7", "")
        timings.append(f"Ручной ввод: {time.perf_counter() - t0:.2f}s")

    excel = None
    workbook = None
    pythoncom.CoInitialize()

    try:
        t0 = time.perf_counter()
        excel = win32.DispatchEx("Excel.Application")
        _configure_excel_fast(excel)
        timings.append(f"Запуск Excel: {time.perf_counter() - t0:.2f}s")

        t0 = time.perf_counter()
        workbook = _open_workbook(excel, target_path)
        ws = workbook.Worksheets(1)
        timings.append(f"Открытие файла: {time.perf_counter() - t0:.2f}s")

        if pre_read is not None:
            _apply_manual_row_entries(ws, DATA_START_ROW, manual_entries)
            _apply_l6_l7(ws, l6_value, l7_value)
        else:
            t0 = time.perf_counter()
            current = _read_manual_row_values(ws, DATA_START_ROW)
            empty_fields = tuple(
                (col, label, prompt) for col, label, prompt in MANUAL_ROW15_FIELDS if not current.get(col)
            )
            entered = name_ask_manual_combined(parent, "Ручной ввод", empty_fields, True)
            manual_entries = {c: v for c, v in entered.items() if c in (2, 3, 5)}
            _apply_manual_row_entries(ws, DATA_START_ROW, manual_entries)
            _apply_l6_l7(ws, entered.get("L6", ""), entered.get("L7", ""))
            timings.append(f"Ручной ввод: {time.perf_counter() - t0:.2f}s")

        _ensure_manual_row_text_format(ws, DATA_START_ROW)

        t0 = time.perf_counter()
        filename_col = _find_filename_column(ws)
        total_search_row = DATA_START_ROW + 1
        total_col = _find_total_column(ws, total_search_row)
        total_row = DATA_START_ROW + data_count

        cell_a15_value = ws.Cells(DATA_START_ROW, 1).Value
        is_numeric, base_value = _is_numeric_value(cell_a15_value)

        _unmerge_row_cells(ws, 18, ("A",))
        _unmerge_row_cells(ws, 19, ("A",))

        if data_count >= 2:
            insert_count = data_count - 1
            source_row = ws.Rows(DATA_START_ROW)
            target_rows = ws.Rows(f"{DATA_START_ROW + 1}:{DATA_START_ROW + insert_count}")
            source_row.Copy()
            target_rows.Insert(Shift=XL_SHIFT_DOWN)
            excel.CutCopyMode = False

        _write_column_values(ws, DATA_START_ROW, filename_col, file_names)

        if 1 <= data_count <= 5:
            total_cell = ws.Cells(total_row, total_col)
            if total_cell.MergeCells:
                total_cell.MergeArea.Cells(1, 1).Value = "Итого"
            else:
                total_cell.Value = "Итого"

        if data_count >= 2 and is_numeric:
            col_a_values = [base_value + i for i in range(1, data_count)]
            _write_column_values(ws, DATA_START_ROW + 1, 1, col_a_values)

        timings.append(f"Обработка данных: {time.perf_counter() - t0:.2f}s")

        t0 = time.perf_counter()
        workbook.Save()
        timings.append(f"Сохранение: {time.perf_counter() - t0:.2f}s")

        elapsed_total = time.perf_counter() - start_total
        timing_report = "\n".join(f"  • {line}" for line in timings)
        messagebox.showinfo(
            "Готово",
            f"Данные успешно перенесены!\n"
            f"Файл: {os.path.basename(target_path)}\n"
            f"Строк перенесено: {data_count}\n\n"
            f"Время выполнения: {elapsed_total:.2f}s\n"
            f"{timing_report}",
            parent=dialog_parent(parent),
        )
    except Exception as exc:
        messagebox.showerror("Ошибка", f"Произошла ошибка:\n{exc}", parent=dialog_parent(parent))
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                _restore_excel_defaults(excel)
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def _apply_manual_entries_to_workbook_array(
        ws,
        manual_entries: dict[int, str],
        l6_value: str,
        l7_value: str,
) -> None:
    row = DATA_START_ROW
    while True:
        r_value = ws.Cells(row, R_COLUMN).Value
        if _cell_text(r_value) == "":
            break

        current = _read_manual_row_values(ws, row)
        to_apply = {
            col: value
            for col, value in manual_entries.items()
            if value and not current.get(col)
        }
        if to_apply:
            _apply_manual_row_entries(ws, row, to_apply)
            _ensure_manual_row_text_format(ws, row)

        row += 1
    _apply_l6_l7(ws, l6_value, l7_value)


def run_name_array(parent: tk.Misc | None = None) -> None:
    if not HAS_WIN32:
        messagebox.showerror(
            "Ошибка",
            "Для работы с Excel требуется Windows, Microsoft Excel и pywin32.\nУстановите: pip install pywin32",
            parent=dialog_parent(parent),
        )
        return

    folder_path = name_select_folder(parent)
    if not folder_path:
        return

    excel_files = collect_files_with_extensions_recursive(folder_path, EXCEL_EXTENSIONS)
    if not excel_files:
        messagebox.showwarning("Внимание", "В выбранной папке нет файлов Excel.", parent=dialog_parent(parent))
        return

    proceed = messagebox.askyesno(
        "Подтверждение",
        f"Найдено файлов Excel: {len(excel_files)}.\nПродолжить обработку?",
        parent=dialog_parent(parent),
    )
    if not proceed:
        return

    entered = name_ask_manual_combined(parent, "Ручной ввод — массив РДО", MANUAL_ROW15_FIELDS, True)
    manual_entries = {c: v for c, v in entered.items() if c in (2, 3, 5) and v}
    l6_value = entered.get("L6", "")
    l7_value = entered.get("L7", "")

    start_total = time.perf_counter()
    pythoncom.CoInitialize()
    excel = None
    processed = 0
    errors: list[str] = []

    try:
        excel = win32.DispatchEx("Excel.Application")
        _configure_excel_fast(excel)

        for path in excel_files:
            workbook = None
            try:
                workbook = _open_workbook(excel, path)
                ws = workbook.Worksheets(1)
                _apply_manual_entries_to_workbook_array(ws, manual_entries, l6_value, l7_value)
                workbook.Save()
                processed += 1
            except Exception as exc:
                errors.append(f"{os.path.basename(path)}: {exc}")
            finally:
                if workbook is not None:
                    try:
                        workbook.Close(SaveChanges=False)
                    except Exception:
                        pass
    finally:
        if excel is not None:
            try:
                _restore_excel_defaults(excel)
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()

    elapsed_total = time.perf_counter() - start_total
    summary = f"\n\nОбщее время обработки: {elapsed_total:.2f}s"

    if errors:
        messagebox.showwarning(
            "Готово с предупреждениями",
            f"Обработано успешно: {processed} из {len(excel_files)}.\n\nОшибки:\n" + "\n".join(errors) + summary,
            parent=dialog_parent(parent),
        )
    else:
        messagebox.showinfo("Готово", f"Обработано файлов Excel: {processed} из {len(excel_files)}." + summary, parent=dialog_parent(parent))


def run_name_pick_folder_and_list(parent: tk.Misc | None, listbox: tk.Listbox) -> None:
    folder_path = name_select_folder(parent)
    if not folder_path:
        return

    file_names = name_list_filenames_recursive(folder_path)
    if not file_names:
        messagebox.showwarning("Внимание", "В выбранной папке нет файлов.", parent=dialog_parent(parent))
        return

    listbox.delete(0, tk.END)
    for name in file_names:
        listbox.insert(tk.END, name)


def run_insert_to_rdo(parent: tk.Misc | None, listbox: tk.Listbox) -> None:
    if not HAS_WIN32:
        messagebox.showerror(
            "Ошибка",
            "Для работы с Excel требуется Windows, Microsoft Excel и pywin32.\nУстановите: pip install pywin32",
            parent=dialog_parent(parent),
        )
        return

    file_names = list(listbox.get(0, tk.END))
    if not file_names:
        messagebox.showwarning(
            "Внимание",
            "Список в разделе «Название файла» пуст. Сначала нажмите «Наименование файла в РДО».",
            parent=dialog_parent(parent),
        )
        return

    target_path = name_select_target_file(parent)
    if not target_path:
        return

    name_process_workbook(target_path, file_names, parent)