# -*- coding: utf-8 -*-
from __future__ import annotations

import platform
import tkinter as tk
from tkinter import ttk
from typing import Callable

from config import UI
from excel_module import (
    run_convert_excel_to_xml,
    run_dzo_insert_to_rdo,
    run_dzo_select_folder_and_process,
    run_insert_to_rdo,
    run_name_array,
    run_name_pick_folder_and_list,
)

from igk_module import run_igk, run_igk_array
from ui.widgets import (ModernNeonCardButton, build_dzo_rdo_table_section, build_filename_section)


class IGKNameApp:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("ДОСПГФ")
        self.root.resizable(False, False)
        self.root.configure(bg=UI["bg_main"])

        if platform.system() == "Windows":
            try:
                import ctypes
                hwnd = ctypes.windll.user32.GetParent(self.root.winfo_id())
                ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 20, ctypes.byref(ctypes.c_int(1)), 4)
            except Exception:
                pass

        self._setup_styles()
        self._build_ui()
        self._center_window()

    def _setup_styles(self) -> None:
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(".", background=UI["bg_main"], foreground=UI["text"])
        style.configure("TFrame", background=UI["bg_main"])
        style.configure("Vertical.TScrollbar", background=UI["bg_main"], troughcolor=UI["card"], arrowcolor=UI["muted"])

    def _center_window(self) -> None:
        self.root.update_idletasks()
        w, h = 1000, 850
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = max(0, (sw - w) // 2)
        y = max(0, (sh - h) // 2)
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    def _build_ui(self) -> None:
        # Sidebar
        sidebar = tk.Frame(self.root, bg=UI["bg_sidebar"], width=240)
        sidebar.pack(side=tk.LEFT, fill=tk.Y)
        sidebar.pack_propagate(False)

        logo_frame = tk.Frame(sidebar, bg=UI["bg_sidebar"])
        logo_frame.pack(pady=(32, 24))

        logo_icon = tk.Label(logo_frame, text="⚡", font=("icomoon", 24), fg=UI["accent_primary"], bg=UI["bg_sidebar"])
        logo_icon.pack()

        logo_text = tk.Label(logo_frame, text="Ростелеком", font=("Segoe UI", 16, "bold"), bg=UI["bg_sidebar"], fg=UI["text"])
        logo_text.pack(pady=4)

        self.menu_buttons = []
        self._create_sidebar_btn(sidebar, "Работа с одним пакетом", 0, pady=(10, 4))
        self._create_sidebar_btn(sidebar, "Работа с массивом пакетов", 1, pady=4)
        self._create_sidebar_btn(sidebar, "Работа с одним пакетом ДЗО", 2, pady=4)

        footer_lbl = tk.Label(sidebar, text="Версия 3.1 (Modern UI)", font=UI["font_footer"], bg=UI["bg_sidebar"], fg=UI["muted"])
        footer_lbl.pack(side=tk.BOTTOM, pady=20)

        # Main Content
        self.main_content = tk.Frame(self.root, bg=UI["bg_main"])
        self.main_content.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=24, pady=24)

        self.pages = []
        self.page_single = tk.Frame(self.main_content, bg=UI["bg_main"])
        self.page_array = tk.Frame(self.main_content, bg=UI["bg_main"])
        self.page_dzo = tk.Frame(self.main_content, bg=UI["bg_main"])

        self.pages.append(self.page_single)
        self.pages.append(self.page_array)
        self.pages.append(self.page_dzo)

        self._build_single_page(self.page_single)
        self._build_array_page(self.page_array)
        self._build_dzo_page(self.page_dzo)

        self._show_page(0)

    def _create_sidebar_btn(self, parent, text: str, index: int, **kwargs) -> None:
        btn = tk.Button(
            parent,
            text=text,
            font=("Segoe UI", 9, "bold"),
            bg=UI["bg_sidebar"],
            fg=UI["muted"],
            activebackground=UI["bg_sidebar"],
            activeforeground=UI["text"],
            relief=tk.FLAT,
            anchor="w",
            padx=20,
            pady=12,
            command=lambda: self._show_page(index)
        )
        btn.pack(fill=tk.X, **kwargs)
        self.menu_buttons.append(btn)

    def _show_page(self, index: int) -> None:
        for i, page in enumerate(self.pages):
            if i == index:
                page.pack(fill=tk.BOTH, expand=True)
                self.menu_buttons[i].config(bg=UI["card"], fg=UI["accent_primary"])
            else:
                page.pack_forget()
                self.menu_buttons[i].config(bg=UI["bg_sidebar"], fg=UI["muted"])

    def _build_header_section(self, parent: tk.Widget, title: str, subtitle: str) -> None:
        header = tk.Frame(parent, bg=UI["bg_main"])
        header.pack(fill=tk.X, pady=(0, 10))
        tk.Label(header, text=title, font=UI["font_title"], bg=UI["bg_main"], fg=UI["text"]).pack(anchor="w")
        tk.Label(header, text=subtitle, font=UI["font_subtitle"], bg=UI["bg_main"], fg=UI["muted"]).pack(anchor="w", pady=(2, 0))

    def _build_single_page(self, parent: tk.Widget) -> None:
        self._build_header_section(parent, "Формирование ПУД (ДОСПГФ)", "Выберите инструмент для запуска")

        cards_frame = tk.Frame(parent, bg=UI["bg_main"])
        cards_frame.pack(fill=tk.X, pady=(8, 2))

        ModernNeonCardButton(
            cards_frame,
            text="Вставка ИГК в файлы",
            subtitle="Добавить номер ИГК в PDF, JPEG, PNG, GIF, TIFF",
            icon="📄",
            command=lambda: self._launch(run_igk),
        ).pack(fill=tk.X, pady=(0, 6))

        ModernNeonCardButton(
            cards_frame,
            text="Выбор папки для РДО",
            subtitle="Загрузить список всех файлов из директории в таблицу ниже",
            icon="📁",
            command=lambda: self._launch(
                lambda p: run_name_pick_folder_and_list(p, self.filename_listbox)
            ),
        ).pack(fill=tk.X, pady=(0, 6))

        self.filename_listbox = build_filename_section(parent)

        insert_frame = tk.Frame(parent, bg=UI["bg_main"])
        insert_frame.pack(fill=tk.X, pady=(4, 0))
        ModernNeonCardButton(
            insert_frame,
            text="Перенести список в РДО (Excel)",
            subtitle="Экспортировать отсортированный список файлов в целевой шаблон",
            icon="📊",
            command=lambda: self._launch(
                lambda p: run_insert_to_rdo(p, self.filename_listbox)
            ),
        ).pack(fill=tk.X)

    def _build_array_page(self, parent: tk.Widget) -> None:
        self._build_header_section(parent, "Пакетная обработка", "Автоматическая работа с массивами директорий")

        cards_frame = tk.Frame(parent, bg=UI["bg_main"])
        cards_frame.pack(fill=tk.X, pady=(8, 4))

        ModernNeonCardButton(
            cards_frame,
            text="Пакетная вставка ИГК",
            subtitle="Массовая маркировка ПУД во всех папках и подпапках выбранного каталога",
            icon="🗂",
            command=lambda: self._launch(run_igk_array),
        ).pack(fill=tk.X, pady=(0, 8))

        ModernNeonCardButton(
            cards_frame,
            text="Автозаполнение массива РДО",
            subtitle="Обновление данных во всех Excel-шаблонах каталога",
            icon="🗃",
            command=lambda: self._launch(run_name_array),
        ).pack(fill=tk.X, pady=(0, 8))

    def _build_dzo_page(self, parent: tk.Widget) -> None:
        self._build_header_section(parent, "Работа с одним пакетом ДЗО", "Инструменты для обработки пакетов ДЗО")

        top_cards = tk.Frame(parent, bg=UI["bg_main"])
        top_cards.pack(fill=tk.X, pady=(4, 2))

        # Кнопка 1: Вставка ИГК в файлы
        ModernNeonCardButton(
            top_cards,
            text="Вставка ИГК в файлы",
            subtitle="Добавить номер ИГК в выбранные файлы (PDF, JPEG, PNG и др.)",
            icon="📄",
            command=lambda: self._launch(run_igk),
        ).pack(fill=tk.X, pady=(0, 4))

        # Кнопка 2: Выбор папки для РДО
        ModernNeonCardButton(
            top_cards,
            text="Выбор папки для РДО",
            subtitle="Выбрать папку с файлами и Excel-файл",
            icon="📁",
            command=lambda: self._launch(
                lambda p: run_dzo_select_folder_and_process(p, self.dzo_tree)
            ),
        ).pack(fill=tk.X, pady=(0, 4))

        # Раздел с названием "РДО" (Таблица)
        self.dzo_tree = build_dzo_rdo_table_section(parent)

        bottom_cards = tk.Frame(parent, bg=UI["bg_main"])
        bottom_cards.pack(fill=tk.X, pady=(4, 0))

        # Кнопка 3: Перенести список в РДО (Excel)
        ModernNeonCardButton(
            bottom_cards,
            text="Перенести список в РДО (Excel)",
            subtitle="Экспортировать сопоставленные данные в Excel",
            icon="📊",
            command=lambda: self._launch(
                lambda p: run_dzo_insert_to_rdo(p, self.dzo_tree)
            ),
        ).pack(fill=tk.X, pady=(0, 4))

        # Кнопка 4: Конвертация excel в xml
        ModernNeonCardButton(
            bottom_cards,
            text="Конвертация excel в xml",
            subtitle="Преобразовать выбранный Excel файл в формат XML",
            icon="⚙",
            command=lambda: self._launch(run_convert_excel_to_xml),
        ).pack(fill=tk.X)

    def _launch(self, handler: Callable[[tk.Misc | None], None]) -> None:
        self.root.withdraw()
        try:
            handler(self.root)
        finally:
            self.root.deiconify()
            self.root.lift()
            self.root.focus_force()

    def run(self) -> None:
        self.root.mainloop()