# -*- coding: utf-8 -*-
from multiprocessing import cpu_count

# ── Настройки интерфейса (Dark & Neon Purple) ───────────────────
UI = {
    "bg_sidebar": "#0f111a",
    "bg_main": "#151824",
    "card": "#1e2235",
    "card_hover": "#262b44",
    "accent_primary": "#a855f7",
    "accent_hover": "#c084fc",
    "text": "#f8fafc",
    "muted": "#94a3b8",
    "border": "#2e344f",
    "font_title": ("Segoe UI", 18, "bold"),
    "font_subtitle": ("Segoe UI", 9),
    "font_card_title": ("Segoe UI", 12, "bold"),
    "font_card_desc": ("Segoe UI", 9),
    "font_btn": ("Segoe UI", 10, "bold"),
    "font_footer": ("Segoe UI", 8),
    "font_section": ("Segoe UI", 11, "bold"),
}

# ── Константы ИГК ──────────────────────────────────────────────
SUPPORTED_EXTENSIONS = {".pdf", ".jpg", ".jpeg", ".png", ".tif", ".tiff", ".gif"}
PREFIX_IGK = "ИГК "
FONT_NAME = "Arial"
FONT_SIZE = 10
MARGIN_MM = 4.0
ENABLE_PROFILING = False
PARALLEL_WORKERS = max(1, cpu_count() - 1)
PARALLEL_MIN_FILES = 3

# ── Константы Excel ────────────────────────────────────────────
DATA_START_ROW = 15
HEADER_SEARCH_ROWS = (12, 13, 14)
DEFAULT_FILENAME_COL = 19
DEFAULT_TOTAL_COL = 11
XL_SHIFT_DOWN = -4121
XL_CALCULATION_MANUAL = -4135
XL_TEXT_FORMAT = "@"
R_COLUMN = 18
L_COLUMN = 12
MANUAL_ROW15_COLS = (2, 3, 5)
MANUAL_ROW15_FIELDS = (
    (2, "Дата ДД", "Введите дату формата ДД.ММ.ГГГГ"),
    (3, "Номер ДД", "Введите номер ГК"),
    (5, "ИГК", "Введите ИГК"),
)
L6_L7_FIELDS = (
    ("L6", "Номер лицевого счета", "Введите номер лицевого счета. Оставьте пустым, чтобы не менять."),
    ("L7", "Аналитический код раздела на лицевом счете",
     "Введите аналитический код раздела на лицевом счете. Оставьте пустым, чтобы не менять."),
)
EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}