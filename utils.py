# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import platform
import threading
import time
import tkinter as tk
from functools import wraps
from typing import Any, Callable

from config import ENABLE_PROFILING, UI

_profiling_data: dict[str, list[float]] = {}
# Защищает _profiling_data от гонки при параллельной записи из нескольких
# потоков ThreadPoolExecutor (CWE-362: Race Condition).
_profiling_lock = threading.Lock()


def profile_time(func_name: str) -> Callable:
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        def wrapper(*args: Any, **kwargs: Any) -> Any:
            if not ENABLE_PROFILING:
                return func(*args, **kwargs)
            start = time.perf_counter()
            try:
                return func(*args, **kwargs)
            finally:
                elapsed = time.perf_counter() - start
                with _profiling_lock:
                    _profiling_data.setdefault(func_name, []).append(elapsed)

        return wrapper

    return decorator


def clear_profiling_data() -> None:
    with _profiling_lock:
        _profiling_data.clear()


def get_profiling_report() -> str:
    with _profiling_lock:
        if not _profiling_data or not ENABLE_PROFILING:
            return ""
        # Копируем под локом, дальнейшее форматирование делаем уже вне лока
        snapshot = {name: list(times) for name, times in _profiling_data.items()}

    lines = ["\n=== ПРОФИЛИРОВАНИЕ ==="]
    total_time = 0.0
    for func_name in sorted(snapshot.keys()):
        times = snapshot[func_name]
        count = len(times)
        total = sum(times)
        avg = total / count if count else 0.0
        total_time += total
        lines.append(f"{func_name}: {count} вызовов, {total:.3f}s (avg: {avg:.3f}s)")
    lines.append(f"{'─' * 50}\nОБЩЕЕ ВРЕМЯ: {total_time:.3f}s")
    return "\n".join(lines)


def dialog_parent(parent: tk.Misc | None) -> tk.Misc | None:
    if parent is None:
        return None
    try:
        return parent if parent.winfo_viewable() else None
    except tk.TclError:
        return None


def prepare_modal_dialog(parent: tk.Misc | None, title: str) -> tk.Toplevel:
    owner = dialog_parent(parent)
    dialog = tk.Toplevel(owner)
    dialog.withdraw()
    dialog.title(title)
    dialog.resizable(False, False)
    dialog.configure(bg=UI["bg_main"])
    if owner is not None:
        dialog.transient(owner)
    dialog.grab_set()
    return dialog


def show_modal_dialog(dialog: tk.Toplevel) -> None:
    dialog.update_idletasks()
    w = dialog.winfo_reqwidth()
    h = dialog.winfo_reqheight()
    sw = dialog.winfo_screenwidth()
    sh = dialog.winfo_screenheight()
    dialog.geometry(f"+{max(0, (sw - w) // 2)}+{max(0, (sh - h) // 2)}")
    dialog.deiconify()
    dialog.lift()
    dialog.attributes("-topmost", True)
    dialog.focus_force()
    dialog.update_idletasks()
    dialog.attributes("-topmost", False)


def bind_entry_paste_shortcuts(dialog: tk.Misc, entry: tk.Entry) -> None:
    def paste_from_clipboard(_event: tk.Event | None = None) -> str:
        entry.focus_set()
        try:
            clip = dialog.clipboard_get()
        except tk.TclError:
            return "break"
        try:
            if entry.selection_present():
                entry.delete("sel.first", "sel.last")
        except tk.TclError:
            pass
        entry.insert(tk.INSERT, clip)
        return "break"

    def show_context_menu(event: tk.Event) -> None:
        menu = tk.Menu(dialog, tearoff=0, bg=UI["card"], fg=UI["text"], activebackground=UI["accent_primary"])
        menu.add_command(label="Вставить", command=lambda: paste_from_clipboard(None))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def on_control_keypress_paste(event: tk.Event):
        if not (event.state & 0x0004):
            return None
        ks = event.keysym
        if ks in ("v", "V", "Cyrillic_em", "Cyrillic_EM"):
            return paste_from_clipboard(event)
        if platform.system() == "Windows" and getattr(event, "keycode", 0) == 86:
            return paste_from_clipboard(event)
        return None

    entry.bind("<Control-KeyPress>", on_control_keypress_paste)
    entry.bind("<Button-3>", show_context_menu)


def collect_files_with_extensions_recursive(folder_path: str, extensions: set[str]) -> list[str]:
    matched: list[str] = []
    for current_dir, _sub_dirs, files in os.walk(folder_path):
        for file_name in files:
            if os.path.splitext(file_name)[1].lower() in extensions:
                matched.append(os.path.join(current_dir, file_name))
    return matched