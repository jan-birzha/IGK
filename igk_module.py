# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import platform
import time
import tkinter as tk
from concurrent.futures import ThreadPoolExecutor, TimeoutError
from tkinter import filedialog, messagebox, ttk
from typing import Any

from PIL import Image, ImageDraw, ImageFont
import pymupdf
from pypdf import PdfReader, PdfWriter

from config import (
    FONT_SIZE, MARGIN_MM, PARALLEL_MIN_FILES, PARALLEL_WORKERS,
    PREFIX_IGK, SUPPORTED_EXTENSIONS, UI
)
from utils import (
    bind_entry_paste_shortcuts, clear_profiling_data,
    collect_files_with_extensions_recursive, dialog_parent,
    get_profiling_report, prepare_modal_dialog, profile_time,
    show_modal_dialog
)

try:
    pymupdf.TOOLS.mupdf_display_errors(False)
except AttributeError:
    pass


def mm_to_pdf_points(mm: float) -> float:
    return mm * 72.0 / 25.4


def mm_to_pixels(mm: float, dpi: float) -> int:
    return max(1, int(round(mm * dpi / 25.4)))


def resolve_arial_font_path() -> str | None:
    system = platform.system()
    if system == "Windows":
        windir = os.environ.get("WINDIR", r"C:\Windows")
        candidate = os.path.join(windir, "Fonts", "arial.ttf")
        if os.path.isfile(candidate):
            return candidate
    elif system == "Darwin":
        for candidate in ("/Library/Fonts/Arial.ttf", os.path.expanduser("~/Library/Fonts/Arial.ttf")):
            if os.path.isfile(candidate):
                return candidate
    else:
        for candidate in (
                "/usr/share/fonts/truetype/msttcorefonts/Arial.ttf",
                "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        ):
            if os.path.isfile(candidate):
                return candidate
    return None


def load_image_font():
    path = resolve_arial_font_path()
    if path:
        return ImageFont.truetype(path, FONT_SIZE)
    return ImageFont.load_default()


def build_stamp_text(user_text: str) -> str:
    return f"{PREFIX_IGK}{user_text}"


def measure_text_width_pil(draw: ImageDraw.ImageDraw, text: str, font) -> int:
    bbox = draw.textbbox((0, 0), text, font=font)
    return bbox[2] - bbox[0]


@profile_time("resave_pdf_with_pypdf")
def resave_pdf_with_pypdf(path: str) -> None:
    reader = PdfReader(path)
    writer = PdfWriter()
    writer.append(reader)
    with open(path, "wb") as f:
        writer.write(f)


def has_digital_signature(doc: pymupdf.Document) -> bool:
    try:
        for xref in range(1, doc.xref_length()):
            obj = doc.xref_object(xref)
            if "/Sig" in obj or "/Signature" in obj:
                return True
    except Exception:
        pass
    return False


@profile_time("stamp_pdf")
def stamp_pdf(path: str, stamp: str) -> None:
    doc = pymupdf.open(path)
    if has_digital_signature(doc):
        doc.close()
        raise RuntimeError("Файл содержит цифровую подпись. Вставка текста в подписанный PDF невозможна.")
    if doc.page_count == 0:
        doc.close()
        raise ValueError("PDF не содержит страниц.")

    page = doc[0]
    rotation = page.rotation

    m_pt = mm_to_pdf_points(MARGIN_MM)
    band_h = FONT_SIZE * 1.35 + 2
    stamp_w = 200

    w = page.rect.width
    h = page.rect.height

    if rotation == 90:
        x0, y0, x1, y1 = m_pt, m_pt, m_pt + band_h, m_pt + stamp_w
    elif rotation == 180:
        x0, y0, x1, y1 = m_pt, h - m_pt - band_h, m_pt + stamp_w, h - m_pt
    elif rotation == 270:
        x0, y0, x1, y1 = w - m_pt - band_h, h - m_pt - stamp_w, w - m_pt, h - m_pt
    else:
        x0, y0, x1, y1 = w - m_pt - stamp_w, m_pt, w - m_pt, m_pt + band_h

    rect = pymupdf.Rect(x0, y0, x1, y1)
    page.draw_rect(rect, color=(1, 1, 1), fill=(1, 1, 1), overlay=True)

    arial_path = resolve_arial_font_path()
    align_mode = pymupdf.TEXT_ALIGN_LEFT if rotation in (90, 270) else pymupdf.TEXT_ALIGN_RIGHT

    if arial_path:
        inserted = page.insert_textbox(
            rect, stamp, fontfile=arial_path, fontname="arial",
            fontsize=FONT_SIZE, align=align_mode, color=(0, 0, 0), rotate=rotation
        )
    else:
        inserted = page.insert_textbox(
            rect, stamp, fontname="helv",
            fontsize=FONT_SIZE, align=align_mode, color=(0, 0, 0), rotate=rotation
        )

    if inserted <= 0:
        doc.close()
        raise RuntimeError(f"Не удалось вставить текст в PDF. Код возврата: {inserted}")

    doc.save(path, incremental=True, encryption=pymupdf.PDF_ENCRYPT_KEEP, clean=False)
    doc.close()


@profile_time("stamp_raster_image")
def stamp_raster_image(path: str, stamp: str, font) -> None:
    img = Image.open(path)
    if getattr(img, "n_frames", 1) > 1:
        img.seek(0)

    dpi_tuple = img.info.get("dpi") or (72.0, 72.0)
    dpi_x = float(dpi_tuple[0]) if dpi_tuple[0] else 72.0
    margin_px = mm_to_pixels(MARGIN_MM, dpi_x)

    work = img.convert("RGBA")
    overlay = Image.new("RGBA", work.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)

    tw = measure_text_width_pil(draw, stamp, font)
    x = work.size[0] - margin_px - tw
    y = margin_px

    for dx, dy in ((-1, 0), (1, 0), (0, -1), (0, 1)):
        draw.text((x + dx, y + dy), stamp, font=font, fill=(255, 255, 255, 255))
    draw.text((x, y), stamp, font=font, fill=(0, 0, 0, 255))

    combined = Image.alpha_composite(work, overlay)
    save_kwargs: dict[str, Any] = {}
    ext = os.path.splitext(path)[1].lower()

    if ext in (".jpg", ".jpeg"):
        combined = combined.convert("RGB")
        save_kwargs["quality"] = 95
    elif ext == ".gif":
        combined = combined.convert("P", palette=Image.ADAPTIVE)
    elif ext in (".tif", ".tiff"):
        save_kwargs["compression"] = "tiff_lzw"

    combined.save(path, **save_kwargs)
    img.close()


def process_single_file(args: tuple) -> tuple[str, bool, str]:
    path, stamp, font_path = args
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext == ".pdf":
            resave_pdf_with_pypdf(path)
            stamp_pdf(path, stamp)
        else:
            font = ImageFont.truetype(font_path, FONT_SIZE) if font_path else ImageFont.load_default()
            stamp_raster_image(path, stamp, font)
        return (path, True, "")
    except Exception as exc:
        return (path, False, str(exc))


def igk_pick_files(parent: tk.Misc | None = None) -> list[str]:
    paths = filedialog.askopenfilenames(
        parent=dialog_parent(parent),
        title="Выберите файлы (PDF, JPEG, PNG, TIFF, GIF)",
        filetypes=[
            ("Поддерживаемые", "*.pdf;*.jpg;*.jpeg;*.png;*.tif;*.tiff;*.gif"),
            ("PDF", "*.pdf"),
            ("Изображения", "*.jpg;*.jpeg;*.png;*.tif;*.tiff;*.gif"),
            ("Все файлы", "*.*"),
        ],
    )
    return list(paths)


def igk_pick_folder(parent: tk.Misc | None = None) -> str | None:
    folder_path = filedialog.askdirectory(
        parent=dialog_parent(parent),
        title="Выберите папку с файлами (PDF, JPEG, PNG, TIFF, GIF)",
    )
    return folder_path or None


def igk_ask_user_text(parent: tk.Misc | None = None) -> str | None:
    result: dict[str, str | None] = {"value": None}
    dialog = prepare_modal_dialog(parent, "Текст для вставки")

    ttk.Label(
        dialog,
        text="Введите текст для вставки в файлы (будет добавлен префикс «ИГК»):",
        background=UI["bg_main"],
        foreground=UI["text"],
        font=UI["font_card_desc"],
    ).pack(padx=20, pady=(16, 8))

    entry = tk.Entry(dialog, width=50, exportselection=False, font=("Segoe UI", 10), bg=UI["card"], fg=UI["text"],
                     insertbackground=UI["text"], highlightthickness=1, highlightcolor=UI["accent_primary"],
                     highlightbackground=UI["border"], relief=tk.FLAT)
    entry.pack(padx=20, pady=6, fill=tk.X)
    entry.focus_set()
    bind_entry_paste_shortcuts(dialog, entry)

    btn_frame = ttk.Frame(dialog)
    btn_frame.pack(pady=(10, 16))

    def on_ok() -> None:
        result["value"] = entry.get()
        dialog.destroy()

    def on_cancel() -> None:
        result["value"] = None
        dialog.destroy()

    ok_btn = tk.Button(btn_frame, text="OK", command=on_ok, bg=UI["accent_primary"], fg="#ffffff", relief=tk.FLAT,
                       width=10, font=("Segoe UI", 9, "bold"), activebackground=UI["accent_hover"],
                       activeforeground="#ffffff")
    ok_btn.pack(side=tk.LEFT, padx=6)
    cancel_btn = tk.Button(btn_frame, text="Отмена", command=on_cancel, bg=UI["card"], fg=UI["muted"], relief=tk.FLAT,
                           width=10, font=("Segoe UI", 9), highlightthickness=1, highlightbackground=UI["border"],
                           activebackground=UI["card_hover"])
    cancel_btn.pack(side=tk.LEFT, padx=6)

    dialog.bind("<Return>", lambda _e: on_ok())
    dialog.bind("<Escape>", lambda _e: on_cancel())

    show_modal_dialog(dialog)
    dialog.wait_window()
    return result["value"]


def igk_validate_paths(paths: list[str]) -> list[str]:
    return [p for p in paths if os.path.splitext(p)[1].lower() in SUPPORTED_EXTENSIONS]


def _process_igk_paths(paths: list[str], parent: tk.Misc | None = None) -> None:
    clear_profiling_data()
    start_total = time.perf_counter()

    paths = igk_validate_paths(paths)
    if not paths:
        messagebox.showerror("Ошибка", "Ни один выбранный файл не подходит по формату.", parent=dialog_parent(parent))
        return

    user_text = igk_ask_user_text(parent)
    if user_text is None:
        return

    stamp = build_stamp_text(user_text)
    font_path = resolve_arial_font_path()
    total = len(paths)

    # Тайм-аут на обработку одного файла: специально сконструированный
    # "файл-бомба" (PDF/TIFF/GIF с экстремальной структурой) не должен
    # бесконечно занимать поток и блокировать весь пакет (CWE-400, DoS).
    PER_FILE_TIMEOUT_SECONDS = 60

    if total >= PARALLEL_MIN_FILES:
        args_list = [(path, stamp, font_path) for path in paths]
        results = []
        with ThreadPoolExecutor(max_workers=PARALLEL_WORKERS) as executor:
            futures = {executor.submit(process_single_file, args): args[0] for args in args_list}
            for future in futures:
                path = futures[future]
                try:
                    results.append(future.result(timeout=PER_FILE_TIMEOUT_SECONDS))
                except TimeoutError:
                    results.append((path, False, f"Превышено время обработки файла ({PER_FILE_TIMEOUT_SECONDS}s)"))
                except Exception as exc:
                    results.append((path, False, str(exc)))
    else:
        pil_font = load_image_font()
        results = []
        for path in paths:
            ext = os.path.splitext(path)[1].lower()
            try:
                if ext == ".pdf":
                    resave_pdf_with_pypdf(path)
                    stamp_pdf(path, stamp)
                else:
                    stamp_raster_image(path, stamp, pil_font)
                results.append((path, True, ""))
            except Exception as exc:
                results.append((path, False, str(exc)))

    errors = []
    ok_count = sum(1 for _, success, _ in results if success)
    for path, success, error in results:
        if not success:
            errors.append(f"{os.path.basename(path)}: {error}")

    elapsed_total = time.perf_counter() - start_total
    report = get_profiling_report()
    summary = f"\n\nОбщее время обработки: {elapsed_total:.2f}s\nПотоков: {PARALLEL_WORKERS}"

    if errors:
        messagebox.showwarning(
            "Готово с предупреждениями",
            f"Обработано успешно: {ok_count} из {total}.\n\nОшибки:\n" + "\n".join(errors) + report + summary,
            parent=dialog_parent(parent),
        )
    else:
        messagebox.showinfo(
            "Готово",
            f"Текст успешно вставлен во все {total} файл(ов)." + report + summary,
            parent=dialog_parent(parent),
        )


def run_igk(parent: tk.Misc | None = None) -> None:
    paths = igk_pick_files(parent)
    if not paths:
        return
    _process_igk_paths(paths, parent)


def run_igk_array(parent: tk.Misc | None = None) -> None:
    folder_path = igk_pick_folder(parent)
    if not folder_path:
        return

    paths = collect_files_with_extensions_recursive(folder_path, SUPPORTED_EXTENSIONS)
    if not paths:
        messagebox.showwarning(
            "Внимание",
            "В выбранной папке (и подпапках) нет файлов поддерживаемого формата.",
            parent=dialog_parent(parent),
        )
        return

    _process_igk_paths(paths, parent)