# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk
from typing import Callable

from config import UI


def _bind_listbox_drag_reorder(listbox: tk.Listbox) -> None:
    state = {"drag_index": None}

    def on_press(event: tk.Event) -> None:
        index = listbox.nearest(event.y)
        if index >= 0:
            state["drag_index"] = index

    def on_motion(event: tk.Event) -> None:
        if state["drag_index"] is None:
            return
        new_index = listbox.nearest(event.y)
        old_index = state["drag_index"]
        if new_index != old_index and 0 <= new_index < listbox.size():
            value = listbox.get(old_index)
            was_selected = old_index in listbox.curselection()
            listbox.delete(old_index)
            listbox.insert(new_index, value)
            if was_selected:
                listbox.selection_set(new_index)
            state["drag_index"] = new_index

    def on_release(_event: tk.Event) -> None:
        state["drag_index"] = None

    listbox.bind("<Button-1>", on_press, add="+")
    listbox.bind("<B1-Motion>", on_motion, add="+")
    listbox.bind("<ButtonRelease-1>", on_release, add="+")


def _delete_selected_listbox_items(listbox: tk.Listbox) -> None:
    selected = list(listbox.curselection())
    if not selected:
        return
    for index in reversed(selected):
        listbox.delete(index)


def build_filename_section(parent_frame: tk.Widget) -> tk.Listbox:
    section = tk.Frame(parent_frame, bg=UI["bg_main"])
    section.pack(fill=tk.BOTH, expand=True, pady=(8, 4))

    header_row = tk.Frame(section, bg=UI["bg_main"])
    header_row.pack(fill=tk.X)
    tk.Label(
        header_row,
        text="Название файла",
        font=UI["font_section"],
        bg=UI["bg_main"],
        fg=UI["text"],
    ).pack(side=tk.LEFT)

    tk.Label(
        header_row,
        text="Перетаскивайте строки мышью, чтобы изменить порядок. Delete — удалить выбранные.",
        font=UI["font_footer"],
        bg=UI["bg_main"],
        fg=UI["muted"],
        wraplength=350,
        justify=tk.LEFT
    ).pack(side=tk.LEFT, padx=(10, 0))

    list_frame = tk.Frame(section, bg=UI["card"], highlightbackground=UI["border"], highlightthickness=1)
    list_frame.pack(fill=tk.BOTH, expand=True, pady=(6, 6))

    style = ttk.Style()
    style.configure("Dark.Vertical.TScrollbar", troughcolor=UI["card"], background=UI["border"], bordercolor=UI["card"],
                    arrowcolor=UI["text"])

    scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, style="Dark.Vertical.TScrollbar")
    listbox = tk.Listbox(
        list_frame,
        selectmode=tk.EXTENDED,
        activestyle="none",
        font=("Segoe UI", 9),
        bg=UI["card"],
        fg=UI["text"],
        selectbackground=UI["accent_primary"],
        selectforeground="#ffffff",
        relief=tk.FLAT,
        highlightthickness=0,
        yscrollcommand=scrollbar.set,
    )
    scrollbar.config(command=listbox.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=2, pady=2)

    _bind_listbox_drag_reorder(listbox)
    listbox.bind("<Delete>", lambda _e: _delete_selected_listbox_items(listbox))

    action_row = tk.Frame(section, bg=UI["bg_main"])
    action_row.pack(fill=tk.X, pady=(0, 2))

    del_btn = tk.Button(
        action_row,
        text="Удалить выбранное",
        command=lambda: _delete_selected_listbox_items(listbox),
        bg="#ef4444",
        fg="#ffffff",
        relief=tk.FLAT,
        font=("Segoe UI", 8, "bold"),
        activebackground="#dc2626",
        activeforeground="#ffffff",
        padx=10,
        pady=2
    )
    del_btn.pack(side=tk.LEFT)

    return listbox


class ModernNeonCardButton(tk.Canvas):
    def __init__(
            self,
            master,
            text: str,
            subtitle: str,
            icon: str,
            command: Callable[[], None],
            **kwargs,
    ):
        super().__init__(
            master,
            width=550,
            height=76,
            highlightthickness=0,
            bg=UI["bg_main"],
            cursor="hand2",
            **kwargs,
        )
        self._command = command

        self._border = self._create_rounded_rect(2, 2, 548, 74, radius=8, fill=UI["card"], outline=UI["border"], width=1)
        self._icon_box = self._create_rounded_rect(16, 13, 62, 59, radius=6, fill=UI["bg_main"], outline="", width=0)

        self.create_text(39, 36, text=icon, anchor="center", font=("Segoe UI", 18), fill=UI["accent_primary"])
        self._title_text = self.create_text(80, 24, text=text, anchor="w", font=UI["font_card_title"], fill=UI["text"])
        self.create_text(80, 48, text=subtitle, anchor="w", font=UI["font_card_desc"], fill=UI["muted"])

        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<Button-1>", self._on_click)

    def _create_rounded_rect(self, x1, y1, x2, y2, radius=6, **kwargs):
        coords = [
            x1 + radius, y1, x2 - radius, y1, x2, y1, x2, y1 + radius,
            x2, y2 - radius, x2, y2, x2 - radius, y2, x1 + radius, y2,
            x1, y2, x1, y2 - radius, x1, y1 + radius, x1, y1,
        ]
        return self.create_polygon(coords, smooth=True, **kwargs)

    def _on_enter(self, _event=None) -> None:
        self.itemconfig(self._border, fill=UI["card_hover"], outline=UI["accent_primary"], width=1)

    def _on_leave(self, _event=None) -> None:
        self.itemconfig(self._border, fill=UI["card"], outline=UI["border"], width=1)

    def _on_click(self, _event=None) -> None:
        self._command()


# ── Вспомогательные функции для Treeview в ДЗО ─────────────────────────────

def _bind_treeview_drag_reorder(tree: ttk.Treeview) -> None:
    """Обеспечивает перетаскивание строк в Treeview с помощью мыши."""
    state = {"drag_item": None}

    def on_press(event: tk.Event) -> None:
        item = tree.identify_row(event.y)
        if item:
            state["drag_item"] = item

    def on_motion(event: tk.Event) -> None:
        if not state["drag_item"]:
            return
        target_item = tree.identify_row(event.y)
        if target_item and target_item != state["drag_item"]:
            # Определяем, перемещаем выше или ниже
            target_index = tree.index(target_item)
            tree.move(state["drag_item"], "", target_index)

    def on_release(_event: tk.Event) -> None:
        state["drag_item"] = None

    tree.bind("<Button-1>", on_press, add="+")
    tree.bind("<B1-Motion>", on_motion, add="+")
    tree.bind("<ButtonRelease-1>", on_release, add="+")


def _delete_selected_treeview_items(tree: ttk.Treeview) -> None:
    """Удаляет выделенные строки из Treeview."""
    selected = tree.selection()
    if not selected:
        return
    for item in selected:
        tree.delete(item)


# ── Обновленный раздел РДО с тёмной темой и управлением ──────────────────────

def build_dzo_rdo_table_section(parent_frame: tk.Widget) -> ttk.Treeview:
    """
    Создает раздел 'РДО' с тёмной стилизацией таблицы, перетаскиванием строк,
    кнопкой и клавишей Delete для удаления элементов.
    """
    section = tk.Frame(parent_frame, bg=UI["bg_main"])
    section.pack(fill=tk.BOTH, expand=True, pady=(6, 4))

    # Заголовок раздела
    header_row = tk.Frame(section, bg=UI["bg_main"])
    header_row.pack(fill=tk.X, pady=(0, 4))

    tk.Label(
        header_row,
        text="РДО",
        font=UI["font_section"],
        bg=UI["bg_main"],
        fg=UI["text"],
    ).pack(side=tk.LEFT)

    tk.Label(
        header_row,
        text="Перетаскивайте строки мышью, чтобы изменить порядок. Delete — удалить выбранные.",
        font=UI["font_footer"],
        bg=UI["bg_main"],
        fg=UI["muted"],
        wraplength=350,
        justify=tk.LEFT
    ).pack(side=tk.LEFT, padx=(10, 0))

    # Стилизация Treeview под тёмную тему (исправление белого фона!)
    style = ttk.Style()
    style.theme_use("default")

    style.configure(
        "DZO.Treeview",
        background=UI["card"],
        foreground=UI["text"],
        fieldbackground=UI["card"],
        borderwidth=0,
        font=("Segoe UI", 9),
        rowheight=24
    )
    style.map(
        "DZO.Treeview",
        background=[("selected", UI["accent_primary"])],
        foreground=[("selected", "#ffffff")]
    )
    style.configure(
        "DZO.Treeview.Heading",
        background=UI["bg_main"],
        foreground=UI["text"],
        relief="flat",
        font=("Segoe UI", 9, "bold")
    )
    style.map("DZO.Treeview.Heading", background=[("active", UI["card_hover"])])

    table_frame = tk.Frame(section, bg=UI["card"], highlightbackground=UI["border"], highlightthickness=1)
    table_frame.pack(fill=tk.BOTH, expand=True)

    columns = ("col_a", "col_b", "col_d", "filename")

    # Включаем selectmode="extended" для выделения нескольких строк
    tree = ttk.Treeview(
        table_frame,
        columns=columns,
        show="headings",
        selectmode="extended",
        style="DZO.Treeview"
    )

    tree.heading("col_a", text="Дата")
    tree.heading("col_b", text="Номер")
    tree.heading("col_d", text="Сумма")
    tree.heading("filename", text="Наименование файла")

    tree.column("col_a", width=110, anchor="center")
    tree.column("col_b", width=100, anchor="center")
    tree.column("col_d", width=110, anchor="e")
    tree.column("filename", width=320, anchor="w")

    scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, style="Dark.Vertical.TScrollbar", command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)

    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=2, pady=2)

    # Привязка перетаскивания и клавиши Delete
    _bind_treeview_drag_reorder(tree)
    tree.bind("<Delete>", lambda _e: _delete_selected_treeview_items(tree))

    # Нижняя панель с кнопкой удаления
    action_row = tk.Frame(section, bg=UI["bg_main"])
    action_row.pack(fill=tk.X, pady=(4, 0))

    del_btn = tk.Button(
        action_row,
        text="Удалить выбранное",
        command=lambda: _delete_selected_treeview_items(tree),
        bg="#ef4444",
        fg="#ffffff",
        relief=tk.FLAT,
        font=("Segoe UI", 8, "bold"),
        activebackground="#dc2626",
        activeforeground="#ffffff",
        padx=10,
        pady=2
    )
    del_btn.pack(side=tk.LEFT)

    return tree