# -*- coding: utf-8 -*-
"""
Программа вставки текста с префиксом «ИГК» в файлы PDF и растровые изображения.
"""

from __future__ import annotations

import os
import platform
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Операция: подключаем библиотеку для работы с изображениями (JPEG, PNG, TIFF, GIF).
from PIL import Image, ImageDraw, ImageFont

# Операция: подключаем библиотеку для вставки текста на страницы PDF.
import fitz

# Операция: отключаем вывод предупреждений MuPDF в stderr (например «No default Layer config» при слоях OCG).
try:
    fitz.TOOLS.mupdf_display_errors(False)
except AttributeError:
    pass


# Операция: задаём расширения файлов, с которыми умеет работать программа.
SUPPORTED_EXTENSIONS = {".pdf", ".jpg", ".jpeg", ".png", ".tif", ".tiff", ".gif"}

# Операция: фиксированный префикс перед пользовательским текстом (кириллица, верхний регистр).
PREFIX_IGK = "ИГК "

# Операция: имя и размер шрифта по заданию.
FONT_NAME = "Arial"
FONT_SIZE = 9

# Операция: отступ от правого и верхнего края в миллиметрах (одинаковый для PDF и растра по смыслу «6 мм»).
MARGIN_MM = 4.0


def mm_to_pdf_points(mm: float) -> float:
    # Операция: переводим миллиметры в PDF-пункты (1 pt = 1/72 дюйма).
    return mm * 72.0 / 25.4


def mm_to_pixels(mm: float, dpi: float) -> int:
    # Операция: переводим миллиметры в пиксели растра по фактическому DPI изображения (или запасному значению).
    return max(1, int(round(mm * dpi / 25.4)))


def resolve_arial_font_path() -> str | None:
    # Операция: определяем путь к файлу шрифта Arial в типичных расположениях ОС.
    system = platform.system()
    if system == "Windows":
        windir = os.environ.get("WINDIR", r"C:\Windows")
        candidate = os.path.join(windir, "Fonts", "arial.ttf")
        if os.path.isfile(candidate):
            return candidate
    elif system == "Darwin":
        candidate = "/Library/Fonts/Arial.ttf"
        if os.path.isfile(candidate):
            return candidate
        candidate = os.path.expanduser("~/Library/Fonts/Arial.ttf")
        if os.path.isfile(candidate):
            return candidate
    else:
        for candidate in (
            "/usr/share/fonts/truetype/msttcorefonts/Arial.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        ):
            if os.path.isfile(candidate):
                return candidate
    # Операция: если Arial не найден, возвращаем None — дальше используем запасной шрифт.
    return None


def load_image_font() -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    # Операция: загружаем шрифт размером 10 pt для отрисовки на изображениях.
    path = resolve_arial_font_path()
    if path:
        return ImageFont.truetype(path, FONT_SIZE)
    # Операция: запасной вариант — встроенный шрифт PIL (внешний вид отличается от Arial).
    return ImageFont.load_default()


def build_stamp_text(user_text: str) -> str:
    # Операция: формируем полную строку для вставки: префикс «ИГК » + текст пользователя.
    return f"{PREFIX_IGK}{user_text}"


def measure_text_width_pil(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.FreeTypeFont | ImageFont.ImageFont) -> int:
    # Операция: измеряем ширину строки в пикселях для выравнивания по правому краю.
    bbox = draw.textbbox((0, 0), text, font=font)
    return bbox[2] - bbox[0]


def stamp_raster_image(path: str, stamp: str, font: ImageFont.FreeTypeFont | ImageFont.ImageFont) -> None:
    # Операция: открываем файл изображения (может быть многокадровый GIF — берём первый кадр).
    img = Image.open(path)
    if getattr(img, "n_frames", 1) > 1:
        img.seek(0)
    # Операция: определяем DPI из метаданных изображения для перевода 6 мм в пиксели (без DPI — 72).
    dpi_tuple = img.info.get("dpi") or (72.0, 72.0)
    dpi_x = float(dpi_tuple[0]) if dpi_tuple[0] else 72.0
    margin_px = mm_to_pixels(MARGIN_MM, dpi_x)
    # Операция: переводим в RGBA, чтобы корректно рисовать поверх любых режимов (в т.ч. с прозрачностью).
    work = img.convert("RGBA")
    overlay = Image.new("RGBA", work.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)
    # Операция: вычисляем координаты текста в правом верхнем углу (6 мм сверху и 6 мм справа от краёв).
    tw = measure_text_width_pil(draw, stamp, font)
    x = work.size[0] - margin_px - tw
    y = margin_px
    # Операция: рисуем чёрный текст с лёгкой белой обводкой для читаемости на любом фоне.
    for dx, dy in ((-1, 0), (1, 0), (0, -1), (0, 1)):
        draw.text((x + dx, y + dy), stamp, font=font, fill=(255, 255, 255, 255))
    draw.text((x, y), stamp, font=font, fill=(0, 0, 0, 255))
    # Операция: накладываем слой с текстом на изображение.
    combined = Image.alpha_composite(work, overlay)
    # Операция: сохраняем результат в тот же формат и путь (перезапись исходного файла).
    save_kwargs: dict = {}
    ext = os.path.splitext(path)[1].lower()
    if ext in (".jpg", ".jpeg"):
        combined = combined.convert("RGB")
        save_kwargs["quality"] = 95
    elif ext == ".gif":
        combined = combined.convert("P", palette=Image.ADAPTIVE)
    elif ext in (".tif", ".tiff"):
        save_kwargs["compression"] = "tiff_lzw"
    combined.save(path, **save_kwargs)
    # Операция: закрываем дескриптор изображения после записи на диск.
    img.close()


def stamp_pdf(path: str, stamp: str) -> None:
    # Операция: открываем PDF-документ для изменения.
    doc = fitz.open(path)
    arial_path = resolve_arial_font_path()
    # Операция: проверяем, что в документе есть хотя бы одна страница (штамп только на первой).
    if doc.page_count == 0:
        doc.close()
        raise ValueError("PDF не содержит страниц.")
    page = doc[0]
    # Операция: отступ 6 мм сверху и 6 мм справа в координатах страницы (пункты); полоса по высоте под одну строку.
    m_pt = mm_to_pdf_points(MARGIN_MM)
    band_h = FONT_SIZE * 1.35 + 2
    rect = fitz.Rect(0, m_pt, page.rect.width - m_pt, m_pt + band_h)
    # Операция: вставляем текст только на первой странице; fontname — всегда строка (имя шрифта в PDF).
    # Для Arial передаём путь к TTF в fontfile и строковое имя "arial", иначе insert_textbox
    # ошибочно получает не строку и падает с "'Font' object has no attribute 'startswith'".
    if arial_path:
        page.insert_textbox(
            rect,
            stamp,
            fontfile=arial_path,
            fontname="arial",
            fontsize=FONT_SIZE,
            align=fitz.TEXT_ALIGN_RIGHT,
            color=(0, 0, 0),
        )
    else:
        page.insert_textbox(
            rect,
            stamp,
            fontname="helv",
            fontsize=FONT_SIZE,
            align=fitz.TEXT_ALIGN_RIGHT,
            color=(0, 0, 0),
        )
    # Операция: сохраняем PDF; clean=True помогает убрать лишнюю структуру без смены логики вставки текста.
    doc.save(path, incremental=True, encryption=fitz.PDF_ENCRYPT_KEEP, clean=True)
    # Операция: закрываем документ после записи.
    doc.close()


def pick_files() -> list[str]:
    # Операция: показываем диалог выбора одного или нескольких файлов с фильтром по поддерживаемым типам.
    root = tk.Tk()
    root.withdraw()
    paths = filedialog.askopenfilenames(
        title="Выберите файлы (PDF, JPEG, PNG, TIFF, GIF)",
        filetypes=[
            ("Поддерживаемые", "*.pdf;*.jpg;*.jpeg;*.png;*.tif;*.tiff;*.gif"),
            ("PDF", "*.pdf"),
            ("Изображения", "*.jpg;*.jpeg;*.png;*.tif;*.tiff;*.gif"),
            ("Все файлы", "*.*"),
        ],
    )
    root.destroy()
    return list(paths)


def ask_user_text() -> str | None:
    # Операция: показываем модальное окно для ввода произвольного текста и кнопки «OK» / «Отмена».
    result: dict[str, str | None] = {"value": None}

    dialog = tk.Tk()
    dialog.title("Текст для вставки")
    dialog.resizable(False, False)

    # Операция: подпись над полем ввода поясняет пользователю назначение окна.
    ttk.Label(dialog, text="Введите текст для вставки в файлы (будет добавлен префикс «ИГК»):").pack(
        padx=12, pady=(12, 6)
    )
    # Операция: классический tk.Entry и явная вставка из буфера — на Windows вызов <<Paste>> из обработчика
    # клавиш часто не вставляет текст, зато чтение clipboard_get() и insert() работает так же, как пункт «Вставить».
    entry = tk.Entry(dialog, width=50, exportselection=False)
    entry.pack(padx=12, pady=6, fill=tk.X)
    entry.focus_set()

    def paste_from_clipboard(_event: tk.Event | None = None) -> str:
        # Операция: вставка из буфера в позицию курсора (Ctrl+V / Ctrl+М, пункт «Вставить» в меню).
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
        # Операция: контекстное меню по правой кнопке мыши — «Вставить» вызывает ту же функцию, что Ctrl+V / Ctrl+М.
        menu = tk.Menu(dialog, tearoff=0)
        menu.add_command(label="Вставить", command=lambda: paste_from_clipboard(None))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def on_control_keypress_paste(event: tk.Event) -> str | None:
        # Операция: вставка по Ctrl+V (EN) или Ctrl+М (RU) — в строке bind() нельзя указать кириллицу «м», поэтому
        # ловим Control+KeyPress и проверяем keysym (лат. v / Cyrillic_em) и на Windows — сканкод клавиши V (86).
        if not (event.state & 0x0004):
            return None
        ks = event.keysym
        if ks in ("v", "V", "Cyrillic_em", "Cyrillic_EM"):
            return paste_from_clipboard(event)
        if platform.system() == "Windows" and getattr(event, "keycode", 0) == 86:
            return paste_from_clipboard(event)
        return None

    # Операция: одна привязка на все варианты вставки — без <Control-м> в строке (TclError: bad keysym «м»).
    entry.bind("<Control-KeyPress>", on_control_keypress_paste)
    entry.bind("<Button-3>", show_context_menu)

    btn_frame = ttk.Frame(dialog)
    btn_frame.pack(pady=(6, 12))

    def on_ok() -> None:
        # Операция: при нажатии «OK» сохраняем введённую строку и закрываем окно.
        result["value"] = entry.get()
        dialog.destroy()

    def on_cancel() -> None:
        # Операция: при «Отмена» оставляем значение None и закрываем окно.
        result["value"] = None
        dialog.destroy()

    ttk.Button(btn_frame, text="OK", command=on_ok).pack(side=tk.LEFT, padx=4)
    ttk.Button(btn_frame, text="Отмена", command=on_cancel).pack(side=tk.LEFT, padx=4)

    dialog.bind("<Return>", lambda _e: on_ok())
    dialog.bind("<Escape>", lambda _e: on_cancel())

    # Операция: центрируем окно на экране после расчёта требуемых размеров содержимого.
    dialog.update_idletasks()
    req_w = dialog.winfo_reqwidth()
    req_h = dialog.winfo_reqheight()
    scr_w = dialog.winfo_screenwidth()
    scr_h = dialog.winfo_screenheight()
    pos_x = max(0, (scr_w - req_w) // 2)
    pos_y = max(0, (scr_h - req_h) // 2)
    dialog.geometry(f"+{pos_x}+{pos_y}")

    dialog.mainloop()
    return result["value"]


def validate_paths(paths: list[str]) -> list[str]:
    # Операция: отфильтровываем файлы с неподдерживаемым расширением.
    ok: list[str] = []
    bad: list[str] = []
    for p in paths:
        ext = os.path.splitext(p)[1].lower()
        if ext in SUPPORTED_EXTENSIONS:
            ok.append(p)
        else:
            bad.append(p)
    return ok


def main() -> None:
    # Операция: запрашиваем у пользователя список файлов для обработки.
    paths = pick_files()
    if not paths:
        # Операция: если файлы не выбраны — завершаем работу без сообщений (пользователь отменил выбор).
        sys.exit(0)

    # Операция: оставляем только поддерживаемые расширения.
    paths = validate_paths(paths)
    if not paths:
        messagebox.showerror("Ошибка", "Ни один выбранный файл не подходит по формату.")
        sys.exit(1)

    # Операция: запрашиваем текст для вставки во втором окне.
    user_text = ask_user_text()
    if user_text is None:
        # Операция: пользователь закрыл окно или нажал «Отмена» — выходим без изменений файлов.
        sys.exit(0)

    stamp = build_stamp_text(user_text)
    # Операция: один раз загружаем шрифт для всех растровых файлов.
    pil_font = load_image_font()

    errors: list[str] = []
    for path in paths:
        ext = os.path.splitext(path)[1].lower()
        try:
            if ext == ".pdf":
                # Операция: для PDF вызываем вставку текста только на первую страницу.
                stamp_pdf(path, stamp)
            else:
                # Операция: для JPEG/PNG/TIFF/GIF рисуем текст поверх изображения.
                stamp_raster_image(path, stamp, pil_font)
        except Exception as exc:  # noqa: BLE001 — показываем пользователю любую ошибку записи/формата.
            errors.append(f"{path}: {exc}")

    # Операция: показываем итог: успех или список файлов с ошибками.
    if errors:
        messagebox.showwarning(
            "Готово с предупреждениями",
            "Часть файлов обработана с ошибками:\n\n" + "\n".join(errors),
        )
    else:
        messagebox.showinfo("Готово", "Текст успешно вставлен во все выбранные файлы.")


if __name__ == "__main__":
    main()
