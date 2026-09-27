"""
Recent Office Finder
A modern Windows desktop app for quickly indexing and finding recently modified
folders, Word documents, and Excel spreadsheets.
Uses only the Python standard library; no pip packages required.
"""

import json
import os
import shutil
import subprocess
import sys
import threading
import zipfile
from dataclasses import dataclass
from datetime import datetime, timedelta
from difflib import SequenceMatcher
from pathlib import Path
from queue import Empty, Queue
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Enable Windows High-DPI awareness for crisp typography and UI elements
if sys.platform.startswith("win"):
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            windll.user32.SetProcessDPIAware()
        except Exception:
            pass

# Default Configuration
# Keep published source machine-neutral. When no local config exists,
# load_config() falls back to the current user's Documents/Desktop folders.
DEFAULT_SCAN_ROOTS = []
DEFAULT_SEARCH_DAYS = 0  # 0 means all time (no cutoff), 30/90/180/365 limit by days
DEFAULT_FILE_SCOPE = "common"  # "common", "office", "all"

# Categorized Extensions
EXT_WORD = {".doc", ".docx"}
EXT_EXCEL = {".xls", ".xlsx", ".xlsm", ".csv", ".tsv"}
EXT_PDF_PPT = {".pdf", ".ppt", ".pptx"}
EXT_TEXT = {".txt", ".md", ".log", ".json", ".xml", ".yaml", ".yml", ".ini", ".cfg", ".conf"}
EXT_MEDIA = {".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp", ".svg", ".ico", ".mp4", ".mp3", ".wav", ".m4a", ".mkv", ".avi"}
EXT_CODE = {".py", ".bat", ".ps1", ".sh", ".js", ".ts", ".html", ".css", ".sql", ".rs", ".go", ".java", ".c", ".cpp", ".h"}

# Pre-compiled extension groups
COMMON_EXTENSIONS = EXT_WORD | EXT_EXCEL | EXT_PDF_PPT | EXT_TEXT | EXT_MEDIA | EXT_CODE
OFFICE_EXTENSIONS = EXT_WORD | EXT_EXCEL | EXT_PDF_PPT | {".txt", ".md"}

ALLOWED_EXTENSIONS = COMMON_EXTENSIONS
IGNORE_DIR_NAMES = {
    ".git", ".idea", ".vscode", "__pycache__", "node_modules", "target",
    "build", "dist", ".venv", "venv", "$RECYCLE.BIN", "System Volume Information",
    ".pytest_tmp", ".pytest_cache"
}
MAX_RESULTS = 2000
CONFIG_FILE = Path(__file__).resolve().parent / "finder_config.json"

FONT_FAMILY = "Microsoft JhengHei UI" if sys.platform.startswith("win") else "Segoe UI"
FONT_MONO = "Consolas" if sys.platform.startswith("win") else "Courier"

# Modern Dark Slate Theme Palette
PALETTE = {
    "bg_main": "#0b121e",          # Deep slate-navy base background
    "bg_sidebar": "#101a2b",       # Left rail panel background
    "bg_card": "#152236",          # Cards and tables background
    "bg_card_alt": "#1a2a42",      # Hover / secondary card background
    "bg_input": "#0d1624",         # Recessed input field
    "border_subtle": "#22354f",    # Subtle borders and dividers
    "border_focus": "#3b82f6",     # Accent focus border
    "text_main": "#f1f5f9",        # High-contrast primary text
    "text_muted": "#94a3b8",       # Secondary muted text
    "text_dim": "#576b85",         # Dim hint / shortcut text
    "accent_blue": "#3b82f6",      # Primary blue
    "accent_blue_hover": "#60a5fa",
    "accent_blue_active": "#2563eb",
    "accent_blue_wash": "#193256",  # Selection wash
    "accent_green": "#10b981",     # Excel green
    "accent_green_hover": "#34d399",
    "accent_purple": "#a78bfa",    # Word purple
    "accent_amber": "#f59e0b",     # Folder amber
    "accent_red": "#ef4444",       # Warning/Error red
    "tree_row_even": "#142033",
    "tree_row_odd": "#18263c",
    "tree_selected": "#1d4ed8",
    "tree_selected_fg": "#ffffff",
}


def load_config():
    """Load scan roots and preferences from configuration file or defaults."""
    roots = list(DEFAULT_SCAN_ROOTS)
    search_days = DEFAULT_SEARCH_DAYS
    file_scope = DEFAULT_FILE_SCOPE
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data.get("scan_roots"), list) and data["scan_roots"]:
                    roots = data["scan_roots"]
                if isinstance(data.get("search_days"), int) and data["search_days"] >= 0:
                    search_days = data["search_days"]
                if isinstance(data.get("file_scope"), str) and data["file_scope"] in {"common", "office", "all"}:
                    file_scope = data["file_scope"]
        except Exception:
            pass
    else:
        # If default roots don't exist on this machine, add common user folders as helpful defaults
        any_exists = any(Path(r).exists() for r in roots)
        if not any_exists:
            user_home = Path.home()
            common_candidates = [user_home / "Documents", user_home / "Desktop"]
            fallback = [str(p) for p in common_candidates if p.exists()]
            if fallback:
                roots = fallback
                save_config(roots, search_days, file_scope)
    return roots, search_days, file_scope


def save_config(roots, search_days, file_scope=DEFAULT_FILE_SCOPE):
    """Save scan roots and preferences to configuration file."""
    try:
        data = {
            "scan_roots": roots,
            "search_days": search_days,
            "file_scope": file_scope,
            "updated_at": datetime.now().isoformat()
        }
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


@dataclass(frozen=True)
class FinderItem:
    path: Path
    is_dir: bool
    modified_time: float
    size_bytes: int = 0

    @property
    def name(self):
        return self.path.name

    @property
    def type_text(self):
        if self.is_dir:
            return "Folder"
        ext = self.path.suffix.lower()
        if ext in EXT_WORD:
            return "Word"
        elif ext in EXT_EXCEL:
            return "Excel"
        elif ext in EXT_PDF_PPT:
            return "PDF" if ext == ".pdf" else "PPT"
        elif ext in EXT_TEXT:
            return "Text"
        elif ext in EXT_MEDIA:
            return "Media"
        elif ext in EXT_CODE:
            return "Code"
        return "File"

    @property
    def type_badge(self):
        if self.is_dir:
            return "📁"
        ext = self.path.suffix.lower()
        if ext in {".xls", ".xlsx", ".xlsm"}:
            return "📊"
        elif ext in {".csv", ".tsv"}:
            return "📊"
        elif ext in {".doc", ".docx"}:
            return "📄"
        elif ext in {".ppt", ".pptx"}:
            return "📽️"
        elif ext == ".pdf":
            return "📑"
        elif ext in {".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp", ".svg", ".ico"}:
            return "🖼️"
        elif ext in {".mp4", ".mkv", ".avi"}:
            return "🎬"
        elif ext in {".mp3", ".wav", ".m4a"}:
            return "🎵"
        elif ext in {".py", ".bat", ".ps1", ".sh", ".js", ".ts", ".html", ".css", ".sql"}:
            return "💻"
        elif ext in {".txt", ".md", ".log"}:
            return "📝"
        elif ext in {".json", ".xml", ".yaml", ".yml", ".ini", ".cfg", ".conf"}:
            return "⚙️"
        return "📄"

    @property
    def tag_name(self):
        if self.is_dir:
            return "folder"
        ext = self.path.suffix.lower()
        if ext in {".xls", ".xlsx", ".xlsm"}:
            return "excel"
        elif ext in {".csv", ".tsv"}:
            return "csv"
        elif ext in {".doc", ".docx"}:
            return "word"
        elif ext in {".ppt", ".pptx"}:
            return "ppt"
        elif ext == ".pdf":
            return "pdf"
        elif ext in {".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp", ".svg", ".ico"}:
            return "image"
        elif ext in {".mp4", ".mp3", ".wav", ".m4a", ".mkv", ".avi"}:
            return "media"
        elif ext in {".py", ".bat", ".ps1", ".sh", ".js", ".ts", ".html", ".css", ".sql"}:
            return "code"
        elif ext in {".txt", ".md", ".log", ".json", ".xml", ".yaml", ".yml"}:
            return "text"
        return "file"

    @property
    def size_text(self):
        if self.is_dir:
            return "—"
        b = self.size_bytes
        if b < 1024:
            return f"{b} B"
        elif b < 1024 * 1024:
            return f"{b / 1024:.1f} KB"
        elif b < 1024 * 1024 * 1024:
            return f"{b / (1024 * 1024):.1f} MB"
        return f"{b / (1024 * 1024 * 1024):.2f} GB"

    @property
    def modified_text(self):
        modified = datetime.fromtimestamp(self.modified_time)
        today = datetime.now().date()
        mod_date = modified.date()
        if mod_date == today:
            return f"今天 {modified:%H:%M}"
        if mod_date == (today - timedelta(days=1)):
            return f"昨天 {modified:%H:%M}"
        if mod_date.year == today.year:
            return modified.strftime("%m-%d %H:%M")
        return modified.strftime("%Y-%m-%d %H:%M")


class ModernButton(tk.Button):
    """A sleek, modern flat button with responsive hover and active effects."""
    def __init__(self, parent, text="", command=None, bg="#1a293f", fg="#e2e8f0",
                 hover_bg="#263b59", active_bg="#162337", disabled_bg="#131b26",
                 disabled_fg="#4b5b70", font=None, padx=12, pady=7, anchor="center", **kwargs):
        self.default_bg = bg
        self.default_fg = fg
        self.hover_bg = hover_bg
        self.active_bg = active_bg
        self.disabled_bg = disabled_bg
        self.disabled_fg = disabled_fg
        super().__init__(
            parent, text=text, command=command,
            bg=bg, fg=fg, activebackground=active_bg, activeforeground=fg,
            relief="flat", bd=0, padx=padx, pady=pady, anchor=anchor,
            font=font or (FONT_FAMILY, 9), cursor="hand2", **kwargs
        )
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)

    def _on_enter(self, _event=None):
        if self["state"] != "disabled":
            self.configure(bg=self.hover_bg)

    def _on_leave(self, _event=None):
        if self["state"] != "disabled":
            self.configure(bg=self.default_bg)

    def set_active_style(self, is_active=True, active_color=PALETTE["accent_blue_wash"],
                         active_text=PALETTE["text_main"]):
        """Highlight button for filter toggle states."""
        if is_active:
            self.default_bg = active_color
            self.default_fg = active_text
            self.configure(bg=active_color, fg=active_text)
        else:
            self.default_bg = "#162338"
            self.default_fg = PALETTE["text_muted"]
            self.configure(bg="#162338", fg=PALETTE["text_muted"])

    def set_enabled(self, enabled=True):
        if enabled:
            self.configure(state="normal", bg=self.default_bg, fg=self.default_fg, cursor="hand2")
        else:
            self.configure(state="disabled", bg=self.disabled_bg, fg=self.disabled_fg, cursor="arrow")


class SettingsDialog(tk.Toplevel):
    """Dialog for configuring scanned root directories and search preferences."""
    def __init__(self, parent, current_roots, current_search_days, current_file_scope, on_save_callback, root_counts=None):
        super().__init__(parent)
        self.title("設定掃描目錄與搜尋範圍 - Recent Office Finder")
        # Keep the directory list and bottom action buttons visible together.
        self.geometry("780x820")
        self.minsize(700, 700)
        self.configure(bg=PALETTE["bg_main"])
        self.transient(parent)
        self.grab_set()

        # Windows dark mode title bar
        self._apply_dark_titlebar()

        self.on_save_callback = on_save_callback
        self.roots = list(current_roots)
        self.search_days_var = tk.IntVar(value=current_search_days)
        self.file_scope_var = tk.StringVar(value=current_file_scope or "common")
        self.path_input_var = tk.StringVar()
        self.days_buttons = []
        self.scope_buttons = []
        self.root_counts = root_counts or {}

        self._build_ui()
        self.center_window(parent)

    def _apply_dark_titlebar(self):
        if sys.platform.startswith("win"):
            try:
                import ctypes
                hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
                value = ctypes.c_int(2)
                ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 20, ctypes.byref(value), ctypes.sizeof(value))
            except Exception:
                pass

    def center_window(self, parent):
        self.update_idletasks()
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        px = parent.winfo_rootx()
        py = parent.winfo_rooty()
        w = self.winfo_width()
        h = self.winfo_height()
        x = max(0, px + (pw - w) // 2)
        y = max(0, py + (ph - h) // 2)
        self.geometry(f"+{x}+{y}")

    def _build_ui(self):
        shell = tk.Frame(self, bg=PALETTE["bg_main"], padx=20, pady=18)
        shell.pack(fill=tk.BOTH, expand=True)

        header = tk.Frame(shell, bg=PALETTE["bg_main"])
        header.pack(fill=tk.X, pady=(0, 10))
        tk.Label(header, text="⚙️  設定掃描目錄與索引範圍", font=(FONT_FAMILY, 15, "bold"),
                 fg=PALETTE["text_main"], bg=PALETTE["bg_main"]).pack(anchor=tk.W)
        tk.Label(header, text="支援 Office (Word/Excel/PPT)、CSV 表格、PDF 文件、筆記、圖片、影音與程式碼。",
                 font=(FONT_FAMILY, 9), fg=PALETTE["text_muted"], bg=PALETTE["bg_main"]).pack(anchor=tk.W, pady=(3, 0))

        # Add Directory Input Card
        input_card = tk.Frame(shell, bg=PALETTE["bg_card"], padx=14, pady=10)
        input_card.pack(fill=tk.X, pady=(0, 10))
        input_card.columnconfigure(0, weight=1)

        tk.Label(input_card, text="加入新目錄 (貼上完整路徑或點選瀏覽選擇)：", font=(FONT_FAMILY, 9, "bold"),
                 fg=PALETTE["accent_blue_hover"], bg=PALETTE["bg_card"]).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 6))

        entry_box = tk.Frame(input_card, bg=PALETTE["border_subtle"], padx=1, pady=1)
        entry_box.grid(row=1, column=0, sticky="ew", padx=(0, 8))
        entry_box.columnconfigure(0, weight=1)

        self.path_entry = tk.Entry(
            entry_box, textvariable=self.path_input_var, font=(FONT_FAMILY, 10),
            bg=PALETTE["bg_input"], fg=PALETTE["text_main"], insertbackground=PALETTE["accent_blue"],
            relief="flat", bd=6
        )
        self.path_entry.pack(fill=tk.BOTH, expand=True)
        self.path_entry.bind("<FocusIn>", lambda _e: entry_box.configure(bg=PALETTE["border_focus"]))
        self.path_entry.bind("<FocusOut>", lambda _e: entry_box.configure(bg=PALETTE["border_subtle"]))
        self.path_entry.bind("<Return>", lambda _e: self._add_from_text())

        ModernButton(input_card, text="📁 瀏覽...", command=self._browse_and_set,
                     bg="#23354e", fg=PALETTE["text_main"], hover_bg="#314768", padx=12, pady=5).grid(row=1, column=1, padx=(0, 6))

        ModernButton(input_card, text="＋ 加入目錄", command=self._add_from_text,
                     bg=PALETTE["accent_blue"], fg="#ffffff", hover_bg=PALETTE["accent_blue_hover"],
                     active_bg=PALETTE["accent_blue_active"], font=(FONT_FAMILY, 9, "bold"), padx=14, pady=5).grid(row=1, column=2)

        # Options Card (Time Scope & File Scope)
        opts_card = tk.Frame(shell, bg=PALETTE["bg_card"], padx=14, pady=10)
        opts_card.pack(fill=tk.X, pady=(0, 10))

        # Time Scope Row
        time_row = tk.Frame(opts_card, bg=PALETTE["bg_card"])
        time_row.pack(fill=tk.X, pady=(0, 6))
        tk.Label(time_row, text="📅 索引時間範圍：", font=(FONT_FAMILY, 9, "bold"),
                 fg=PALETTE["accent_blue_hover"], bg=PALETTE["bg_card"]).pack(side=tk.LEFT, padx=(0, 8))

        scope_options = [
            ("不限時間 (全部檔案)", 0),
            ("最近 30 天", 30),
            ("最近 90 天", 90),
            ("最近 180 天", 180),
            ("最近 365 天", 365),
        ]
        self.days_buttons = []
        for text, days in scope_options:
            btn = ModernButton(time_row, text=text, command=lambda d=days: self._set_days(d),
                               padx=8, pady=3, font=(FONT_FAMILY, 8))
            btn.pack(side=tk.LEFT, padx=2)
            self.days_buttons.append((btn, days))

        # File Scope Row
        file_row = tk.Frame(opts_card, bg=PALETTE["bg_card"])
        file_row.pack(fill=tk.X)
        tk.Label(file_row, text="📁 檔案類型範圍：", font=(FONT_FAMILY, 9, "bold"),
                 fg=PALETTE["accent_blue_hover"], bg=PALETTE["bg_card"]).pack(side=tk.LEFT, padx=(0, 8))

        type_options = [
            ("全部常用 (Office/圖片/影音/程式碼)", "common"),
            ("僅 Office 與文件", "office"),
            ("不限副檔名 (*.*)", "all"),
        ]
        self.scope_buttons = []
        for text, scope in type_options:
            btn = ModernButton(file_row, text=text, command=lambda s=scope: self._set_scope(s),
                               padx=8, pady=3, font=(FONT_FAMILY, 8))
            btn.pack(side=tk.LEFT, padx=2)
            self.scope_buttons.append((btn, scope))

        self._update_days_buttons()
        self._update_scope_buttons()

        # List frame
        list_header = tk.Frame(shell, bg=PALETTE["bg_main"])
        list_header.pack(fill=tk.X, pady=(0, 4))
        tk.Label(list_header, text="目前掃描目錄清單 (各目錄索引筆數)：", font=(FONT_FAMILY, 9, "bold"),
                 fg=PALETTE["text_dim"], bg=PALETTE["bg_main"]).pack(side=tk.LEFT)

        list_container = tk.Frame(shell, bg=PALETTE["border_subtle"], padx=1, pady=1)
        list_container.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        self.listbox = tk.Listbox(
            list_container, bg=PALETTE["bg_card"], fg=PALETTE["text_main"],
            selectbackground=PALETTE["accent_blue"], selectforeground="#ffffff",
            relief="flat", bd=6, font=(FONT_MONO, 9), activestyle="none"
        )
        scroll = tk.Scrollbar(list_container, orient=tk.VERTICAL, command=self.listbox.yview, bg=PALETTE["bg_card"])
        self.listbox.configure(yscrollcommand=scroll.set)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self._refresh_listbox()

        # Action toolbar
        toolbar = tk.Frame(shell, bg=PALETTE["bg_main"])
        toolbar.pack(fill=tk.X, pady=(0, 12))

        ModernButton(toolbar, text="－ 移除選取項目", command=self._remove_selected,
                     bg="#25354e", fg=PALETTE["text_main"], hover_bg="#344869").pack(side=tk.LEFT, padx=(0, 8))

        ModernButton(toolbar, text="↺ 設為常用目錄 (文件 / 桌面)", command=self._reset_common,
                     bg="#1c283c", fg=PALETTE["text_muted"], hover_bg="#263750").pack(side=tk.LEFT, padx=(0, 8))

        ModernButton(toolbar, text="🗑️ 清空清單", command=self._clear_all,
                     bg="#221e24", fg="#f87171", hover_bg="#3b242a").pack(side=tk.LEFT)

        # Bottom buttons
        bottom = tk.Frame(shell, bg=PALETTE["bg_main"])
        bottom.pack(fill=tk.X)

        ModernButton(bottom, text="取消", command=self.destroy,
                     bg="#1e2c40", fg=PALETTE["text_muted"], hover_bg="#2a3c56").pack(side=tk.RIGHT, padx=(8, 0))

        ModernButton(bottom, text="💾 儲存並重新掃描", command=self._save_and_close,
                     bg=PALETTE["accent_green"], fg="#ffffff", hover_bg=PALETTE["accent_green_hover"],
                     font=(FONT_FAMILY, 9, "bold"), padx=18).pack(side=tk.RIGHT)

    def _set_days(self, days):
        self.search_days_var.set(days)
        self._update_days_buttons()

    def _update_days_buttons(self):
        curr = self.search_days_var.get()
        for btn, d in self.days_buttons:
            btn.set_active_style(curr == d)

    def _set_scope(self, scope):
        self.file_scope_var.set(scope)
        self._update_scope_buttons()

    def _update_scope_buttons(self):
        curr = self.file_scope_var.get()
        for btn, s in self.scope_buttons:
            btn.set_active_style(curr == s)

    def _refresh_listbox(self):
        self.listbox.delete(0, tk.END)
        for path_str in self.roots:
            p = Path(path_str)
            if not p.exists():
                status = "⚠️ 不存在"
            else:
                cnt = self.root_counts.get(path_str, None)
                if cnt is not None:
                    status = f"✅ 存在 ({cnt:,} 筆)"
                else:
                    status = "✅ 存在"
            self.listbox.insert(tk.END, f"{status:<16}  {path_str}")

    def _add_from_text(self):
        raw = self.path_input_var.get().strip()
        if not raw:
            return False
        clean_path = raw.strip('"').strip("'").strip()
        if not clean_path:
            return False
        p = Path(clean_path).resolve()
        path_str = str(p)
        if any(os.path.normcase(r) == os.path.normcase(path_str) for r in self.roots):
            messagebox.showinfo("提示", f"目錄「{path_str}」已在清單中。", parent=self)
            self.path_input_var.set("")
            return False
        if not p.exists():
            if not messagebox.askyesno("路徑不存在", f"目錄「{path_str}」目前在電腦上不存在。\n\n是否仍要加入清單（如尚未掛載的磁碟或遠端資料夾）？", parent=self):
                return False
        elif not p.is_dir():
            messagebox.showwarning("提示", "所輸入的路徑不是資料夾目錄，請確認後再試。", parent=self)
            return False

        self.roots.append(path_str)
        self.path_input_var.set("")
        self._refresh_listbox()
        return True

    def _browse_and_set(self):
        chosen = filedialog.askdirectory(title="選擇要加入索引的資料夾", parent=self)
        if chosen:
            chosen = str(Path(chosen).resolve())
            self.path_input_var.set(chosen)
            self._add_from_text()

    def _remove_selected(self):
        selected_idx = self.listbox.curselection()
        if not selected_idx:
            return
        idx = selected_idx[0]
        del self.roots[idx]
        self._refresh_listbox()

    def _clear_all(self):
        if not self.roots:
            return
        if messagebox.askyesno("確認清空", "是否清空所有已設定的掃描目錄？", parent=self):
            self.roots.clear()
            self._refresh_listbox()

    def _reset_common(self):
        home = Path.home()
        candidates = [home / "Documents", home / "Desktop"]
        new_roots = [str(p) for p in candidates if p.exists()]
        if new_roots:
            self.roots = new_roots
            self._refresh_listbox()
        else:
            messagebox.showinfo("提示", "找不到常用目錄。", parent=self)

    def _save_and_close(self):
        # Auto-add any pending directory path in the entry box
        if self.path_input_var.get().strip():
            self._add_from_text()
        if not self.roots:
            messagebox.showwarning("提示", "請至少指定一個掃描目錄。", parent=self)
            return
        self.on_save_callback(self.roots, self.search_days_var.get(), self.file_scope_var.get())
        self.destroy()


class RecentOfficeFinder(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Recent Office Finder - 最近文件索引與快速檢索")

        # Keep the wide desktop layout while using nearly the full vertical
        # workspace, matching the reference layout. The screen-relative lower
        # bound keeps the window usable on smaller displays.
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        default_w = min(1500, max(1280, int(screen_w * 0.82)))
        default_h = min(1580, max(680, screen_h - 20))
        x = max(0, (screen_w - default_w) // 2)
        y = max(0, (screen_h - default_h) // 2 - 25)
        self.geometry(f"{default_w}x{default_h}+{x}+{y}")
        self.minsize(1100, 680)
        self.configure(bg=PALETTE["bg_main"])

        # Enable Windows dark mode native title bar
        self._apply_dark_titlebar()

        # Load persisted config or defaults
        self.scan_roots, self.search_days, self.file_scope = load_config()
        self.root_counts = {}

        self.all_items = []
        self.filtered_items = []
        self.scan_running = False
        self.scan_pending = False
        self.scan_queue = Queue()
        self.missing_roots = []
        self._filter_after_id = None
        self.current_directory = None
        self.navigation_history = []
        self.navigation_forward = []
        self.sidebar_collapsed = False

        # Filter states
        self.active_type_filter = "all"  # "all", "word", "excel", "pdf_ppt", "text", "media", "code", "folder"
        self.active_time_filter = "all"  # "all" (default, all scanned files), "today", "7d", "30d", "90d"

        # Sorting states: (column_name, reverse_boolean)
        self.sort_col = "modified"
        self.sort_reverse = True

        # Observables
        self.search_var = tk.StringVar()
        self.status_var = tk.StringVar(value="準備完成")
        self.index_state_var = tk.StringVar(value="索引待命")
        self.result_var = tk.StringVar(value="正在建立索引")
        self.selection_path_var = tk.StringVar(value="選取項目可查看完整路徑")
        self.toast_var = tk.StringVar(value="")
        self.navigation_path_var = tk.StringVar(value="索引根目錄")

        self.build_ui()
        self.search_var.trace_add("write", lambda *_args: self._schedule_filter())
        self.bind_shortcuts()

        # Poll scan queue and kick off initial indexing
        self.after(100, self.check_scan_queue)
        self.after(200, self.refresh_items)

    def _apply_dark_titlebar(self):
        """Set Windows native title bar to dark mode."""
        if sys.platform.startswith("win"):
            try:
                import ctypes
                hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
                value = ctypes.c_int(2)
                ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 20, ctypes.byref(value), ctypes.sizeof(value))
            except Exception:
                pass

    def build_ui(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        # Configure Treeview and Scrollbars
        style.configure(
            "Treeview",
            background=PALETTE["tree_row_even"],
            fieldbackground=PALETTE["tree_row_even"],
            foreground=PALETTE["text_main"],
            rowheight=35,
            font=(FONT_FAMILY, 9),
            borderwidth=0
        )
        style.configure(
            "Treeview.Heading",
            background="#172439",
            foreground="#9db1cc",
            font=(FONT_FAMILY, 9, "bold"),
            padding=(10, 8),
            relief="flat"
        )
        style.map("Treeview",
                  background=[("selected", PALETTE["tree_selected"])],
                  foreground=[("selected", PALETTE["tree_selected_fg"])])
        style.map("Treeview.Heading",
                  background=[("active", "#223552")])

        style.configure("Vertical.TScrollbar",
                        background="#1e2f47",
                        troughcolor=PALETTE["bg_card"],
                        bordercolor=PALETTE["bg_card"],
                        arrowcolor=PALETTE["text_muted"])
        style.configure("Horizontal.TScrollbar",
                        background="#1e2f47",
                        troughcolor=PALETTE["bg_card"],
                        bordercolor=PALETTE["bg_card"],
                        arrowcolor=PALETTE["text_muted"])
        style.configure("Scan.Horizontal.TProgressbar",
                        background=PALETTE["accent_blue"],
                        troughcolor="#1e2c40",
                        bordercolor="#1e2c40",
                        lightcolor=PALETTE["accent_blue"],
                        darkcolor=PALETTE["accent_blue"])

        # Main shell grid: Column 0 = Left Rail, Column 1 = Main Workspace
        shell = tk.Frame(self, bg=PALETTE["bg_main"], padx=16, pady=16)
        shell.pack(fill=tk.BOTH, expand=True)
        self.shell = shell
        shell.columnconfigure(1, weight=1)
        shell.rowconfigure(0, weight=1)

        # -------------------------------------------------------------
        # Left Rail (Sidebar)
        # -------------------------------------------------------------
        rail = tk.Frame(shell, bg=PALETTE["bg_sidebar"], width=230, padx=14, pady=16)
        rail.grid(row=0, column=0, sticky="nsew", padx=(0, 14))
        rail.grid_propagate(False)
        self.sidebar = rail

        self.sidebar_toggle = ModernButton(
            rail, text="◀  收合", command=self.toggle_sidebar,
            bg="#1b2b43", fg=PALETTE["accent_blue_hover"], hover_bg="#294469",
            font=(FONT_FAMILY, 8, "bold"), anchor="w", padx=8, pady=5
        )
        self.sidebar_toggle.pack(fill=tk.X, pady=(0, 12))

        # Brand header
        tk.Label(rail, text="OFFICE FINDER", font=(FONT_MONO, 9, "bold"),
                 fg=PALETTE["accent_blue_hover"], bg=PALETTE["bg_sidebar"]).pack(anchor=tk.W)
        tk.Label(rail, text="最近文件總管", font=(FONT_FAMILY, 16, "bold"),
                 fg=PALETTE["text_main"], bg=PALETTE["bg_sidebar"]).pack(anchor=tk.W, pady=(4, 2))
        tk.Label(rail, text="依時間追蹤文件、表格與資料夾", font=(FONT_FAMILY, 8),
                 fg=PALETTE["text_muted"], bg=PALETTE["bg_sidebar"]).pack(anchor=tk.W, pady=(0, 14))

        # Divider
        tk.Frame(rail, height=1, bg=PALETTE["border_subtle"]).pack(fill=tk.X, pady=(0, 14))

        # Section: 類型篩選
        tk.Label(rail, text="類型篩選", font=(FONT_MONO, 8, "bold"),
                 fg=PALETTE["text_dim"], bg=PALETTE["bg_sidebar"]).pack(anchor=tk.W, pady=(0, 6))

        self.btn_filter_all = ModernButton(
            rail, text="● 全部項目", command=lambda: self.set_type_filter("all"),
            anchor="w", font=(FONT_FAMILY, 9, "bold"), pady=5
        )
        self.btn_filter_all.pack(fill=tk.X, pady=1)

        self.btn_filter_word = ModernButton(
            rail, text="📄 Word 文件", command=lambda: self.set_type_filter("word"),
            anchor="w", pady=5
        )
        self.btn_filter_word.pack(fill=tk.X, pady=1)

        self.btn_filter_excel = ModernButton(
            rail, text="📊 Excel / CSV", command=lambda: self.set_type_filter("excel"),
            anchor="w", pady=5
        )
        self.btn_filter_excel.pack(fill=tk.X, pady=1)

        self.btn_filter_pdf_ppt = ModernButton(
            rail, text="📑 PDF / 簡報", command=lambda: self.set_type_filter("pdf_ppt"),
            anchor="w", pady=5
        )
        self.btn_filter_pdf_ppt.pack(fill=tk.X, pady=1)

        self.btn_filter_text = ModernButton(
            rail, text="📝 筆記文本", command=lambda: self.set_type_filter("text"),
            anchor="w", pady=5
        )
        self.btn_filter_text.pack(fill=tk.X, pady=1)

        self.btn_filter_media = ModernButton(
            rail, text="🖼️ 圖片 / 影音", command=lambda: self.set_type_filter("media"),
            anchor="w", pady=5
        )
        self.btn_filter_media.pack(fill=tk.X, pady=1)

        self.btn_filter_code = ModernButton(
            rail, text="💻 程式腳本", command=lambda: self.set_type_filter("code"),
            anchor="w", pady=5
        )
        self.btn_filter_code.pack(fill=tk.X, pady=1)

        self.btn_filter_folder = ModernButton(
            rail, text="📁 資料夾目錄", command=lambda: self.set_type_filter("folder"),
            anchor="w", pady=5
        )
        self.btn_filter_folder.pack(fill=tk.X, pady=1)

        # Section: 時間範圍
        tk.Label(rail, text="時間範圍", font=(FONT_MONO, 8, "bold"),
                 fg=PALETTE["text_dim"], bg=PALETTE["bg_sidebar"]).pack(anchor=tk.W, pady=(10, 5))

        time_frame = tk.Frame(rail, bg=PALETTE["bg_sidebar"])
        time_frame.pack(fill=tk.X)
        self.btn_time_all = ModernButton(time_frame, text="全部", command=lambda: self.set_time_filter("all"),
                                         padx=4, pady=5, font=(FONT_FAMILY, 8))
        self.btn_time_all.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 2))

        self.btn_time_today = ModernButton(time_frame, text="今天", command=lambda: self.set_time_filter("today"),
                                           padx=4, pady=5, font=(FONT_FAMILY, 8))
        self.btn_time_today.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 2))

        self.btn_time_7d = ModernButton(time_frame, text="7天", command=lambda: self.set_time_filter("7d"),
                                        padx=4, pady=5, font=(FONT_FAMILY, 8))
        self.btn_time_7d.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 2))

        self.btn_time_30d = ModernButton(time_frame, text="30天", command=lambda: self.set_time_filter("30d"),
                                         padx=4, pady=5, font=(FONT_FAMILY, 8))
        self.btn_time_30d.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 2))

        self.btn_time_90d = ModernButton(time_frame, text="90天", command=lambda: self.set_time_filter("90d"),
                                         padx=4, pady=5, font=(FONT_FAMILY, 8))
        self.btn_time_90d.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Divider
        tk.Frame(rail, height=1, bg=PALETTE["border_subtle"]).pack(fill=tk.X, pady=14)

        # Section: 快捷操作
        tk.Label(rail, text="快速建立與操作", font=(FONT_MONO, 8, "bold"),
                 fg=PALETTE["text_dim"], bg=PALETTE["bg_sidebar"]).pack(anchor=tk.W, pady=(0, 6))

        ModernButton(rail, text="＋  新增 Word 文件", command=self.create_word,
                     bg="#1e2c42", fg="#d8b4fe", hover_bg="#2b3e5e", anchor="w", pady=6).pack(fill=tk.X, pady=2)

        ModernButton(rail, text="＋  新增 Excel 活頁簿", command=self.create_excel,
                     bg="#1e2c42", fg="#6ee7b7", hover_bg="#2b3e5e", anchor="w", pady=6).pack(fill=tk.X, pady=2)

        self.copy_button = ModernButton(rail, text="📑  複製所選檔案...", command=self.duplicate_file,
                                        bg="#1a2538", fg=PALETTE["text_main"], hover_bg="#25354e",
                                        anchor="w", pady=6)
        self.copy_button.pack(fill=tk.X, pady=2)
        self.copy_button.set_enabled(False)

        # Divider
        tk.Frame(rail, height=1, bg=PALETTE["border_subtle"]).pack(fill=tk.X, pady=14)

        # Section: 掃描目錄管理
        tk.Label(rail, text="目錄管理", font=(FONT_MONO, 8, "bold"),
                 fg=PALETTE["text_dim"], bg=PALETTE["bg_sidebar"]).pack(anchor=tk.W, pady=(0, 6))

        ModernButton(rail, text="＋  快速加入目錄...", command=self.quick_add_directory,
                     bg="#1e2c42", fg="#93c5fd", hover_bg="#2b3e5e", anchor="w",
                     font=(FONT_FAMILY, 9), pady=6).pack(fill=tk.X, pady=2)

        ModernButton(rail, text="⚙️  管理掃描目錄", command=self.open_settings_dialog,
                     bg="#19283e", fg=PALETTE["text_muted"], hover_bg="#263d5e", anchor="w",
                     font=(FONT_FAMILY, 9), pady=6).pack(fill=tk.X, pady=2)

        # Keyboard shortcuts hint at bottom of rail
        rail_footer = tk.Frame(rail, bg=PALETTE["bg_sidebar"])
        rail_footer.pack(side=tk.BOTTOM, fill=tk.X)
        tk.Label(rail_footer, text="快捷鍵指引", font=(FONT_MONO, 8, "bold"),
                 fg=PALETTE["text_dim"], bg=PALETTE["bg_sidebar"]).pack(anchor=tk.W, pady=(0, 4))
        shortcuts = [
            ("Enter", "開啟／進入"),
            ("Alt+← / ⌫", "返回上層"),
            ("Alt+→", "前進"),
            ("Alt+↑", "上一層資料夾"),
            ("Ctrl+C", "複製路徑"),
            ("Ctrl+F", "聚焦搜尋"),
            ("Esc", "清除搜尋"),
            ("F5", "重新整理"),
        ]
        for key, desc in shortcuts:
            row = tk.Frame(rail_footer, bg=PALETTE["bg_sidebar"])
            row.pack(fill=tk.X, pady=1)
            tk.Label(row, text=key, font=(FONT_MONO, 8), fg=PALETTE["accent_blue_hover"], bg=PALETTE["bg_sidebar"]).pack(side=tk.LEFT)
            tk.Label(row, text=desc, font=(FONT_FAMILY, 8), fg=PALETTE["text_dim"], bg=PALETTE["bg_sidebar"]).pack(side=tk.RIGHT)

        # -------------------------------------------------------------
        # Main Content Workspace
        # -------------------------------------------------------------
        content = tk.Frame(shell, bg=PALETTE["bg_main"])
        content.grid(row=0, column=1, sticky="nsew")
        content.columnconfigure(0, weight=1)
        content.rowconfigure(2, weight=1)  # Table takes up remaining vertical space

        # Row 0: Top Header & Status Bar
        top_bar = tk.Frame(content, bg=PALETTE["bg_main"])
        top_bar.grid(row=0, column=0, sticky="ew", pady=(0, 12))

        header_left = tk.Frame(top_bar, bg=PALETTE["bg_main"])
        header_left.pack(side=tk.LEFT, fill=tk.Y)
        tk.Label(header_left, text="檔案與目錄索引", font=(FONT_FAMILY, 17, "bold"),
                 fg=PALETTE["text_main"], bg=PALETTE["bg_main"]).pack(side=tk.LEFT)

        # Status badge
        self.status_badge = tk.Label(header_left, textvariable=self.index_state_var,
                                     font=(FONT_MONO, 9, "bold"), fg="#6ee7b7", bg="#142b2f",
                                     padx=9, pady=3)
        self.status_badge.pack(side=tk.LEFT, padx=(12, 0))

        self.navigation_label = tk.Label(
            header_left, textvariable=self.navigation_path_var,
            font=(FONT_MONO, 8), fg=PALETTE["text_muted"], bg=PALETTE["bg_main"],
            anchor=tk.W, width=42
        )
        self.navigation_label.pack(side=tk.LEFT, padx=(10, 0))

        header_right = tk.Frame(top_bar, bg=PALETTE["bg_main"])
        header_right.pack(side=tk.RIGHT, fill=tk.Y)

        self.scan_progress = ttk.Progressbar(header_right, mode="indeterminate", length=90,
                                             style="Scan.Horizontal.TProgressbar")

        self.refresh_button = ModernButton(
            header_right, text="🔄 重新掃描 (F5)", command=self.refresh_items,
            bg=PALETTE["accent_blue"], fg="#ffffff", hover_bg=PALETTE["accent_blue_hover"],
            active_bg=PALETTE["accent_blue_active"], font=(FONT_FAMILY, 9, "bold"), padx=14
        )
        self.refresh_button.pack(side=tk.RIGHT, padx=(10, 0))

        self.home_button = ModernButton(
            header_right, text="⌂ 索引", command=self.navigate_home,
            bg="#1d3049", fg="#dbeafe", hover_bg="#294469", padx=10, pady=6
        )
        self.home_button.pack(side=tk.RIGHT, padx=(8, 0))

        self.back_button = ModernButton(
            header_right, text="↩ 上層", command=self.navigate_back,
            bg="#1d3049", fg="#dbeafe", hover_bg="#294469", padx=10, pady=6
        )
        self.back_button.pack(side=tk.RIGHT, padx=(8, 0))
        self.back_button.set_enabled(False)

        self.forward_button = ModernButton(
            header_right, text="↪ 前進", command=self.navigate_forward,
            bg="#1d3049", fg="#dbeafe", hover_bg="#294469", padx=10, pady=6
        )
        self.forward_button.pack(side=tk.RIGHT, padx=(8, 0))
        self.forward_button.set_enabled(False)

        # Row 1: Search Card
        search_card = tk.Frame(content, bg=PALETTE["bg_card"], padx=14, pady=12)
        search_card.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        search_card.columnconfigure(1, weight=1)

        # Search icon prefix
        tk.Label(search_card, text="🔍", font=(FONT_FAMILY, 12),
                 fg=PALETTE["text_muted"], bg=PALETTE["bg_card"]).grid(row=0, column=0, padx=(0, 8))

        # Search entry container with custom border
        entry_box = tk.Frame(search_card, bg=PALETTE["border_subtle"], padx=1, pady=1)
        entry_box.grid(row=0, column=1, sticky="ew")
        entry_box.columnconfigure(0, weight=1)

        self.search_entry = tk.Entry(
            entry_box, textvariable=self.search_var, font=(FONT_FAMILY, 11),
            bg=PALETTE["bg_input"], fg=PALETTE["text_main"], insertbackground=PALETTE["accent_blue"],
            relief="flat", bd=7
        )
        self.search_entry.pack(fill=tk.BOTH, expand=True)
        self.search_entry.bind("<FocusIn>", lambda _e: entry_box.configure(bg=PALETTE["border_focus"]))
        self.search_entry.bind("<FocusOut>", lambda _e: entry_box.configure(bg=PALETTE["border_subtle"]))
        self.search_entry.bind("<KeyPress-Down>", self._on_search_down_key)
        self.search_entry.bind("<KeyPress-KP_Down>", self._on_search_down_key)
        self.search_entry.bind("<Return>", self._on_search_enter_key)

        # Clear button
        self.clear_button = ModernButton(
            search_card, text="✕ 清除", command=self.clear_search,
            bg="#203047", fg=PALETTE["text_muted"], hover_bg="#2c4263", padx=12, pady=6
        )
        self.clear_button.grid(row=0, column=2, padx=(8, 0))
        self.clear_button.set_enabled(False)

        # Search meta & count row
        meta_row = tk.Frame(search_card, bg=PALETTE["bg_card"])
        meta_row.grid(row=1, column=1, columnspan=2, sticky="ew", pady=(8, 0))

        tk.Label(meta_row, text="💡 雙擊資料夾可逐層進入；按 Esc 清除搜尋，按 ↓ 進入清單",
                 font=(FONT_FAMILY, 8), fg=PALETTE["text_dim"], bg=PALETTE["bg_card"]).pack(side=tk.LEFT)

        self.count_badge = tk.Label(meta_row, textvariable=self.result_var,
                                    font=(FONT_FAMILY, 8, "bold"), fg=PALETTE["accent_blue_hover"],
                                    bg="#15273f", padx=7, pady=1)
        self.count_badge.pack(side=tk.RIGHT)

        # Row 2: Table Card (Treeview)
        table_card = tk.Frame(content, bg=PALETTE["border_subtle"], padx=1, pady=1)
        table_card.grid(row=2, column=0, sticky="nsew")
        table_card.columnconfigure(0, weight=1)
        table_card.rowconfigure(0, weight=1)

        columns = ("type", "name", "modified", "size", "path")
        self.tree = ttk.Treeview(table_card, columns=columns, show="headings", selectmode="browse")

        # Configure column headings with sorting hooks
        self.column_titles = {
            "type": "類型",
            "name": "名稱",
            "modified": "修改時間",
            "size": "大小",
            "path": "所在目錄",
        }
        for key in columns:
            self.tree.heading(key, text=self.column_titles[key],
                              command=lambda c=key: self.sort_by_column(c))

        self.tree.column("type", width=105, anchor=tk.CENTER, stretch=False)
        self.tree.column("name", width=380, minwidth=200)
        self.tree.column("modified", width=150, anchor=tk.CENTER, stretch=False)
        self.tree.column("size", width=100, anchor=tk.E, stretch=False)
        self.tree.column("path", width=460, minwidth=220)

        # Tree tags for item types
        self.tree.tag_configure("even", background=PALETTE["tree_row_even"])
        self.tree.tag_configure("odd", background=PALETTE["tree_row_odd"])
        self.tree.tag_configure("folder", foreground="#60a5fa")   # Sky blue
        self.tree.tag_configure("word", foreground="#c084fc")     # Soft violet
        self.tree.tag_configure("excel", foreground="#34d399")    # Mint green
        self.tree.tag_configure("csv", foreground="#2dd4bf")      # Teal green
        self.tree.tag_configure("pdf", foreground="#f87171")      # Coral red
        self.tree.tag_configure("ppt", foreground="#fb923c")      # Warm orange
        self.tree.tag_configure("text", foreground="#94a3b8")     # Slate gray
        self.tree.tag_configure("image", foreground="#f472b6")    # Pink
        self.tree.tag_configure("media", foreground="#fb7185")    # Rose
        self.tree.tag_configure("code", foreground="#38bdf8")     # Cyan
        self.tree.tag_configure("file", foreground="#cbd5e1")     # Light slate

        scroll_y = ttk.Scrollbar(table_card, orient=tk.VERTICAL, command=self.tree.yview)
        scroll_x = ttk.Scrollbar(table_card, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")

        # Treeview event bindings
        self.tree.bind("<Double-1>", lambda _event: self.open_selected())
        self.tree.bind("<Return>", lambda _event: self.open_selected())
        self.tree.bind("<KeyPress-Down>", self._on_tree_down_key)
        self.tree.bind("<KeyPress-KP_Down>", self._on_tree_down_key)
        self.tree.bind("<<TreeviewSelect>>", self.update_selection)
        self.tree.bind("<Button-3>", self._show_context_menu)  # Right-click context menu (Windows)
        self.tree.bind("<Button-2>", self._show_context_menu)  # macOS trackpad/mouse
        self.tree.bind("<Control-c>", lambda _e: self.copy_selected_path())

        # Empty state overlay container
        self.empty_overlay = tk.Frame(table_card, bg=PALETTE["bg_card"], padx=30, pady=30)
        self.empty_icon = tk.Label(self.empty_overlay, text="🔍", font=(FONT_FAMILY, 28),
                                   fg=PALETTE["text_muted"], bg=PALETTE["bg_card"])
        self.empty_icon.pack(pady=(0, 6))
        self.empty_title = tk.Label(self.empty_overlay, text="", font=(FONT_FAMILY, 13, "bold"),
                                    fg=PALETTE["text_main"], bg=PALETTE["bg_card"])
        self.empty_title.pack()
        self.empty_hint = tk.Label(self.empty_overlay, text="", font=(FONT_FAMILY, 9),
                                   fg=PALETTE["text_muted"], bg=PALETTE["bg_card"], justify=tk.CENTER)
        self.empty_hint.pack(pady=(6, 12))

        self.empty_btn_box = tk.Frame(self.empty_overlay, bg=PALETTE["bg_card"])
        ModernButton(
            self.empty_btn_box, text="＋ 快速加入目錄...", command=self.quick_add_directory,
            bg=PALETTE["accent_blue"], fg="#ffffff", hover_bg=PALETTE["accent_blue_hover"],
            font=(FONT_FAMILY, 9, "bold"), padx=14
        ).pack(side=tk.LEFT, padx=(0, 8))
        ModernButton(
            self.empty_btn_box, text="⚙️  管理掃描目錄", command=self.open_settings_dialog,
            bg="#22354e", fg=PALETTE["text_main"], hover_bg="#304464",
            font=(FONT_FAMILY, 9), padx=14
        ).pack(side=tk.LEFT)

        # Context Menu
        self.context_menu = tk.Menu(self, tearoff=0, bg="#18253b", fg=PALETTE["text_main"],
                                    activebackground=PALETTE["accent_blue"], activeforeground="#ffffff",
                                    relief="solid", bd=1, font=(FONT_FAMILY, 9))
        self.context_menu.add_command(label="🚀  開啟／進入項目", command=self.open_selected)
        self.context_menu.add_command(label="📂  在檔案總管中定位", command=self.reveal_in_explorer)
        self.context_menu.add_command(label="📁  開啟所在資料夾", command=self.open_parent_folder)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="📋  複製完整路徑", command=self.copy_selected_path)
        self.context_menu.add_command(label="📝  複製名稱", command=self.copy_selected_name)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="📑  複製檔案複本...", command=self.duplicate_file)

        # Row 3: Bottom Action & Selected Details Card
        bottom_bar = tk.Frame(content, bg=PALETTE["bg_sidebar"], padx=14, pady=10)
        bottom_bar.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        bottom_bar.columnconfigure(0, weight=1)

        detail_left = tk.Frame(bottom_bar, bg=PALETTE["bg_sidebar"])
        detail_left.grid(row=0, column=0, sticky="w")

        tk.Label(detail_left, text="📍 選取項目：", font=(FONT_FAMILY, 9, "bold"),
                 fg=PALETTE["accent_blue_hover"], bg=PALETTE["bg_sidebar"]).pack(side=tk.LEFT)
        self.selection_label = tk.Label(detail_left, textvariable=self.selection_path_var,
                                        font=(FONT_MONO, 9), fg=PALETTE["text_main"], bg=PALETTE["bg_sidebar"])
        self.selection_label.pack(side=tk.LEFT, padx=(4, 10))

        self.toast_label = tk.Label(detail_left, textvariable=self.toast_var,
                                    font=(FONT_FAMILY, 9, "bold"), fg=PALETTE["accent_green"],
                                    bg=PALETTE["bg_sidebar"])
        self.toast_label.pack(side=tk.LEFT)

        # Quick action buttons on the right side of the bottom bar
        detail_right = tk.Frame(bottom_bar, bg=PALETTE["bg_sidebar"])
        detail_right.grid(row=0, column=1, sticky="e")

        self.btn_bottom_open = ModernButton(
            detail_right, text="🚀 開啟／進入", command=self.open_selected,
            bg="#1d3049", fg="#ffffff", hover_bg="#294469", padx=11, pady=4
        )
        self.btn_bottom_open.pack(side=tk.LEFT, padx=3)
        self.btn_bottom_open.set_enabled(False)

        self.btn_bottom_reveal = ModernButton(
            detail_right, text="📂 檔案總管", command=self.reveal_in_explorer,
            bg="#1d3049", fg="#ffffff", hover_bg="#294469", padx=11, pady=4
        )
        self.btn_bottom_reveal.pack(side=tk.LEFT, padx=3)
        self.btn_bottom_reveal.set_enabled(False)

        self.btn_bottom_copy = ModernButton(
            detail_right, text="📋 複製路徑", command=self.copy_selected_path,
            bg="#1d3049", fg="#ffffff", hover_bg="#294469", padx=11, pady=4
        )
        self.btn_bottom_copy.pack(side=tk.LEFT, padx=3)
        self.btn_bottom_copy.set_enabled(False)

        # Row 4: Status line
        footer_line = tk.Frame(content, bg=PALETTE["bg_main"])
        footer_line.grid(row=4, column=0, sticky="ew", pady=(6, 0))

        tk.Label(footer_line, textvariable=self.status_var, font=(FONT_FAMILY, 8),
                 fg=PALETTE["text_muted"], bg=PALETTE["bg_main"]).pack(side=tk.LEFT)
        tk.Label(footer_line, text="Recent Office Finder · Pure Python", font=(FONT_FAMILY, 8),
                 fg=PALETTE["text_dim"], bg=PALETTE["bg_main"]).pack(side=tk.RIGHT)

        # Update initial filter styles
        self._sidebar_widgets = [
            widget for widget in rail.winfo_children()
            if widget is not self.sidebar_toggle
        ]
        self._sidebar_pack_info = []
        for widget in self._sidebar_widgets:
            pack_info = dict(widget.pack_info())
            pack_info.pop("in", None)  # All saved widgets already belong to rail.
            self._sidebar_pack_info.append((widget, pack_info))
        self._update_filter_button_styles()
        self._update_navigation_controls()

    def bind_shortcuts(self):
        self.bind("<Control-f>", lambda _event: self.focus_search())
        self.bind("<Control-d>", lambda _event: self.duplicate_file())
        self.bind("<F5>", lambda _event: self.refresh_items())
        self.bind("<Alt-Left>", lambda _event: self.navigate_back())
        self.bind("<Alt-Right>", lambda _event: self.navigate_forward())
        self.bind("<Alt-Up>", lambda _event: self.navigate_parent())
        self.bind("<BackSpace>", lambda _event: self.navigate_back())
        self.bind("<Control-Shift-b>", lambda _event: self.toggle_sidebar())
        self.bind("<Escape>", lambda _event: self.clear_search())
        self.bind_all("<KeyPress-Down>", self._on_global_down_key, add="+")
        self.bind_all("<KeyPress-KP_Down>", self._on_global_down_key, add="+")

    def toggle_sidebar(self):
        """Collapse or restore the left filter and action rail."""
        if self.sidebar_collapsed:
            for widget, pack_info in self._sidebar_pack_info:
                widget.pack(**pack_info)
            self.sidebar.configure(width=230, padx=14)
            self.shell.columnconfigure(0, minsize=0)
            self.sidebar_toggle.configure(text="◀  收合", anchor="w")
            self.sidebar_collapsed = False
        else:
            for widget in self._sidebar_widgets:
                widget.pack_forget()
            self.sidebar.configure(width=56, padx=6)
            self.shell.columnconfigure(0, minsize=56)
            self.sidebar_toggle.configure(text="▶", anchor="center")
            self.sidebar_collapsed = True
        self.update_idletasks()

    def _on_search_down_key(self, _event=None):
        """When pressing down arrow in search box, jump focus to table and select first item."""
        self._cancel_scheduled_filter()
        self.apply_filter()
        self._select_first_tree_item()
        return "break"

    def _on_tree_down_key(self, _event=None):
        """Select the first row when Down is pressed with no current selection."""
        if not self.tree.selection():
            self._select_first_tree_item()
            return "break"
        return None

    def _on_global_down_key(self, event=None):
        """Fallback Down-key routing when a Tk class binding consumes the key."""
        widget = self.focus_get()
        if widget is None:
            return None
        try:
            if widget.winfo_toplevel() is not self:
                return None
        except tk.TclError:
            return None

        if widget is self.search_entry:
            return self._on_search_down_key(event)
        if widget is self.tree:
            return self._on_tree_down_key(event)
        if widget is self:
            self._cancel_scheduled_filter()
            self.apply_filter()
            self._select_first_tree_item()
            return "break"
        return None

    def _select_first_tree_item(self):
        """Focus, select, and reveal the first currently displayed row."""
        children = self.tree.get_children()
        if not children:
            return False
        first = children[0]
        self.tree.focus_set()
        self.tree.focus(first)
        self.tree.selection_set(first)
        self.tree.see(first)
        self.update_selection()
        return True

    def _on_search_enter_key(self, _event=None):
        """When pressing Enter in search box, open the first item or selected item."""
        self._cancel_scheduled_filter()
        self.apply_filter()
        selected = self.get_selected_item()
        if selected:
            self.open_selected()
        else:
            if self._select_first_tree_item():
                self.open_selected()
        return "break"

    def focus_search(self):
        self.search_entry.focus_set()
        self.search_entry.select_range(0, tk.END)

    def clear_search(self):
        self.search_var.set("")
        self._cancel_scheduled_filter()
        self.apply_filter()
        self.search_entry.focus_set()

    def _schedule_filter(self):
        """Debounce live search so each keystroke does not rebuild the table."""
        if self._filter_after_id is not None:
            try:
                self.after_cancel(self._filter_after_id)
            except tk.TclError:
                pass
        self._filter_after_id = self.after(180, self._run_scheduled_filter)

    def _run_scheduled_filter(self):
        self._filter_after_id = None
        self.apply_filter()

    def _cancel_scheduled_filter(self):
        """Run an immediate filter without leaving a duplicate debounced callback."""
        if self._filter_after_id is not None:
            try:
                self.after_cancel(self._filter_after_id)
            except tk.TclError:
                pass
            self._filter_after_id = None

    def set_type_filter(self, filter_type):
        self._cancel_scheduled_filter()
        self.active_type_filter = filter_type
        self._update_filter_button_styles()
        self.apply_filter()

    def set_time_filter(self, time_scope):
        self._cancel_scheduled_filter()
        self.active_time_filter = time_scope
        self._update_filter_button_styles()
        self.apply_filter()

    @staticmethod
    def _canonical_path(path):
        """Return a stable path key so Windows alias paths navigate correctly."""
        normalized = os.path.abspath(os.path.normpath(str(path)))
        try:
            # realpath resolves Windows 8.3 aliases such as USER~1.  A folder
            # selected from the index may otherwise use a different spelling
            # than the same folder stored as an indexed item's parent path.
            return os.path.normcase(os.path.realpath(normalized))
        except (OSError, ValueError):
            return os.path.normcase(normalized)

    def _get_navigation_items(self):
        """Return indexed items directly inside the current folder, if any."""
        if self.current_directory is None:
            return self.all_items
        current_key = self._canonical_path(self.current_directory)
        return [
            item for item in self.all_items
            if self._canonical_path(item.path.parent) == current_key
        ]

    def _update_navigation_controls(self):
        """Refresh the breadcrumb and navigation button states."""
        if not hasattr(self, "back_button"):
            return
        if self.current_directory is None:
            self.navigation_path_var.set("索引根目錄")
        else:
            self.navigation_path_var.set(f"📁 {self.current_directory}")
        self.back_button.set_enabled(bool(self.navigation_history))
        self.forward_button.set_enabled(bool(self.navigation_forward))
        self.home_button.set_enabled(self.current_directory is not None)

    def _clear_search_for_navigation(self):
        """Clear search immediately when changing folders."""
        self._cancel_scheduled_filter()
        if self.search_var.get():
            self.search_var.set("")
            self._cancel_scheduled_filter()

    def enter_selected_directory(self):
        """Enter the selected folder and show its direct children."""
        item = self.get_selected_item()
        if not item or not item.is_dir:
            return
        directory = Path(item.path)
        try:
            if not directory.is_dir():
                self.show_toast("資料夾已不存在，請重新掃描")
                return
            directory = directory.resolve()
        except OSError as exc:
            messagebox.showerror("無法進入資料夾", f"無法讀取資料夾：\n{exc}", parent=self)
            return

        self.navigation_history.append(self.current_directory)
        self.current_directory = directory
        self.navigation_forward.clear()
        self._clear_search_for_navigation()
        self._update_navigation_controls()
        self.apply_filter()
        self.show_toast(f"已進入：{directory.name or directory}")

    def navigate_back(self):
        """Match Explorer's Back action for in-app folder history."""
        if not self.navigation_history:
            return
        self.navigation_forward.append(self.current_directory)
        self.current_directory = self.navigation_history.pop()
        self._clear_search_for_navigation()
        self._update_navigation_controls()
        self.apply_filter()

    def navigate_forward(self):
        """Match Explorer's Forward action for in-app folder history."""
        if not self.navigation_forward:
            return
        self.navigation_history.append(self.current_directory)
        self.current_directory = self.navigation_forward.pop()
        self._clear_search_for_navigation()
        self._update_navigation_controls()
        self.apply_filter()

    def navigate_parent(self):
        """Move to the physical parent folder, like Explorer's Alt+Up."""
        if self.current_directory is None:
            return
        parent = self.current_directory.parent
        if parent == self.current_directory:
            return

        parent_key = self._canonical_path(parent)
        inside_scan_root = False
        for root in self.scan_roots:
            root_path = Path(root)
            if not root_path.exists():
                continue
            try:
                if os.path.commonpath([parent_key, self._canonical_path(root_path)]) == self._canonical_path(root_path):
                    inside_scan_root = True
                    break
            except ValueError:
                continue
        if not inside_scan_root:
            return

        self.navigation_history.append(self.current_directory)
        self.current_directory = parent
        self.navigation_forward.clear()
        self._clear_search_for_navigation()
        self._update_navigation_controls()
        self.apply_filter()

    def navigate_home(self):
        """Return to the full indexed view."""
        if self.current_directory is None and not self.navigation_history:
            return
        self.current_directory = None
        self.navigation_history.clear()
        self.navigation_forward.clear()
        self._clear_search_for_navigation()
        self._update_navigation_controls()
        self.apply_filter()

    def _update_filter_button_styles(self):
        # Update type filter buttons
        self.btn_filter_all.set_active_style(self.active_type_filter == "all")
        self.btn_filter_word.set_active_style(self.active_type_filter == "word")
        self.btn_filter_excel.set_active_style(self.active_type_filter == "excel")
        self.btn_filter_pdf_ppt.set_active_style(self.active_type_filter == "pdf_ppt")
        self.btn_filter_text.set_active_style(self.active_type_filter == "text")
        self.btn_filter_media.set_active_style(self.active_type_filter == "media")
        self.btn_filter_code.set_active_style(self.active_type_filter == "code")
        self.btn_filter_folder.set_active_style(self.active_type_filter == "folder")

        # Update time filter buttons
        self.btn_time_all.set_active_style(self.active_time_filter == "all")
        self.btn_time_today.set_active_style(self.active_time_filter == "today")
        self.btn_time_7d.set_active_style(self.active_time_filter == "7d")
        self.btn_time_30d.set_active_style(self.active_time_filter == "30d")
        self.btn_time_90d.set_active_style(self.active_time_filter == "90d")

    def sort_by_column(self, col):
        """Sort items when clicking a column header."""
        if self.sort_col == col:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_col = col
            self.sort_reverse = (col in {"modified", "size"})

        self._update_header_arrows()
        self._sort_and_render()

    def _update_header_arrows(self):
        arrow = " ▼" if self.sort_reverse else " ▲"
        for col, title in self.column_titles.items():
            if col == self.sort_col:
                self.tree.heading(col, text=f"{title}{arrow}")
            else:
                self.tree.heading(col, text=title)

    def quick_add_directory(self):
        """Quickly add a single directory to scan roots via directory picker."""
        chosen = filedialog.askdirectory(title="選擇要加入索引的資料夾", parent=self)
        if chosen:
            chosen = str(Path(chosen).resolve())
            if not any(os.path.normcase(r) == os.path.normcase(chosen) for r in self.scan_roots):
                self.scan_roots.append(chosen)
                save_config(self.scan_roots, self.search_days, self.file_scope)
                self.show_toast(f"已新增目錄：{Path(chosen).name}")
                self.refresh_items()
            else:
                self.show_toast(f"目錄「{Path(chosen).name}」已在索引清單中")

    def open_settings_dialog(self):
        SettingsDialog(
            self,
            self.scan_roots,
            self.search_days,
            self.file_scope,
            self._on_settings_saved,
            self.root_counts
        )

    def _on_settings_saved(self, new_roots, new_search_days, new_file_scope):
        self.scan_roots = list(new_roots)
        self.search_days = new_search_days
        self.file_scope = new_file_scope
        save_config(self.scan_roots, self.search_days, self.file_scope)
        self.show_toast("已儲存設定！正在重新整理索引...")
        self.refresh_items()

    def show_toast(self, message, timeout_ms=3000):
        self.toast_var.set(f"✓ {message}")
        self.after(timeout_ms, lambda: self.toast_var.set(""))

    def refresh_items(self):
        if self.scan_running:
            self.scan_pending = True
            return
        self.scan_running = True
        self.scan_pending = False
        self.refresh_button.set_enabled(False)
        self.status_var.set("正在進行全域掃描...")
        self.index_state_var.set("掃描中...")
        self.status_badge.configure(bg="#2b2612", fg="#facc15")  # Warning/yellow badge
        self.scan_progress.pack(side=tk.RIGHT, padx=(0, 6))
        self.scan_progress.start(10)
        self.apply_filter()

        roots_snapshot = list(self.scan_roots)
        search_days_snapshot = self.search_days
        file_scope_snapshot = self.file_scope
        threading.Thread(
            target=self.scan_worker,
            args=(roots_snapshot, search_days_snapshot, file_scope_snapshot),
            daemon=True
        ).start()

    def scan_worker(self, scan_roots, search_days, file_scope):
        cutoff = (datetime.now() - timedelta(days=search_days)).timestamp() if search_days > 0 else 0
        result = {}
        missing_roots = []
        root_counts = {}

        if file_scope == "all":
            allowed_exts = None
        elif file_scope == "office":
            allowed_exts = OFFICE_EXTENSIONS
        else:
            allowed_exts = COMMON_EXTENSIONS

        # Remove duplicate and nested existing roots before walking the tree. This
        # avoids scanning the same files multiple times when a broad workspace and
        # one of its subdirectories are both configured.
        effective_roots = []
        for root_text in scan_roots:
            root = Path(root_text)
            root_counts[root_text] = 0
            try:
                canonical = os.path.normcase(os.path.abspath(os.path.normpath(str(root))))
            except (OSError, ValueError):
                canonical = os.path.normcase(str(root))

            if root.exists():
                covered = False
                for parent_canonical, parent_text in effective_roots:
                    if not Path(parent_text).exists():
                        continue
                    try:
                        covered = os.path.commonpath([canonical, parent_canonical]) == parent_canonical
                    except ValueError:
                        covered = False
                    if covered:
                        break
                if covered:
                    continue

                # If this root is broader than an earlier existing root, retain
                # only the broader root. Non-existing roots are never discarded.
                retained = []
                for child_canonical, child_text in effective_roots:
                    try:
                        contains_child = os.path.commonpath([canonical, child_canonical]) == canonical
                    except ValueError:
                        contains_child = False
                    if contains_child:
                        root_counts[child_text] = 0
                    else:
                        retained.append((child_canonical, child_text))
                effective_roots = retained

            if not any(canonical == existing_canonical for existing_canonical, _ in effective_roots):
                effective_roots.append((canonical, root_text))

        for _canonical, root_text in effective_roots:
            root = Path(root_text)
            if not root.exists():
                missing_roots.append(root_text)
                continue
            count_before = len(result)
            for current_root, dirs, files in os.walk(root, onerror=lambda _error: None):
                dirs[:] = [name for name in dirs if not name.startswith(".") and name not in IGNORE_DIR_NAMES]
                current = Path(current_root)
                if current != root:
                    self.add_recent(current, True, cutoff, result)
                for filename in files:
                    if filename.startswith("~$") or filename.startswith("."):
                        continue
                    path = current / filename
                    if allowed_exts is None or path.suffix.lower() in allowed_exts:
                        self.add_recent(path, False, cutoff, result)
            root_counts[root_text] = len(result) - count_before

        items = sorted(result.values(), key=lambda item: item.modified_time, reverse=True)
        self.scan_queue.put((items, missing_roots, root_counts, scan_roots, search_days, file_scope))

    @staticmethod
    def add_recent(path, is_dir, cutoff, result):
        try:
            stat_res = path.stat()
            modified = stat_res.st_mtime
            if cutoff <= 0 or modified >= cutoff:
                key = os.path.normcase(os.path.abspath(str(path)))
                size = 0 if is_dir else stat_res.st_size
                result[key] = FinderItem(path, is_dir, modified, size)
                return True
        except OSError:
            pass
        return False

    def check_scan_queue(self):
        try:
            items, missing, root_counts, scanned_roots, scanned_days, scanned_scope = self.scan_queue.get_nowait()
            self.scan_running = False
            self.refresh_button.set_enabled(True)
            self.scan_progress.stop()
            self.scan_progress.pack_forget()

            # A settings change during a scan queues a new scan. Do not display
            # the old result with the new settings in the short transition.
            current_snapshot = (list(self.scan_roots), self.search_days, self.file_scope)
            scanned_snapshot = (list(scanned_roots), scanned_days, scanned_scope)
            if current_snapshot != scanned_snapshot and self.scan_pending:
                self.scan_running = False
                self.refresh_items()
                return

            self.missing_roots = missing
            self.all_items = items
            self.root_counts = root_counts

            if missing and not items:
                self.index_state_var.set("目錄需確認")
                self.status_badge.configure(bg="#3a1c1c", fg="#f87171")
            elif missing:
                self.index_state_var.set(f"部分路徑遺失 ({len(missing)})")
                self.status_badge.configure(bg="#362913", fg="#fbbf24")
            else:
                self.index_state_var.set(f"索引就緒 · {len(items):,} 筆")
                self.status_badge.configure(bg="#142b24", fg="#34d399")

            time_desc = "全部時間" if scanned_days == 0 else f"最近 {scanned_days} 天"

            breakdown_parts = []
            for r in self.scan_roots:
                name = Path(r).name or r
                cnt = root_counts.get(r, 0)
                breakdown_parts.append(f"{name}: {cnt:,} 筆")
            breakdown_str = " ｜ ".join(breakdown_parts)

            status = f"掃描完成：共 {len(self.scan_roots)} 個目錄，索引 {len(items):,} 筆 ({time_desc}) ｜ {breakdown_str}"
            if missing:
                status += f" ｜ ⚠️ {len(missing)} 個目錄不存在"
            self.status_var.set(status)

            # Update category counts on buttons
            word_count = sum(1 for it in items if it.type_text == "Word")
            excel_count = sum(1 for it in items if it.type_text == "Excel")
            pdf_ppt_count = sum(1 for it in items if it.type_text in {"PDF", "PPT"})
            text_count = sum(1 for it in items if it.type_text == "Text")
            media_count = sum(1 for it in items if it.type_text == "Media")
            code_count = sum(1 for it in items if it.type_text == "Code")
            folder_count = sum(1 for it in items if it.is_dir)

            self.btn_filter_all.configure(text=f"● 全部項目 ({len(items):,})")
            self.btn_filter_word.configure(text=f"📄 Word ({word_count:,})")
            self.btn_filter_excel.configure(text=f"📊 Excel/CSV ({excel_count:,})")
            self.btn_filter_pdf_ppt.configure(text=f"📑 PDF/簡報 ({pdf_ppt_count:,})")
            self.btn_filter_text.configure(text=f"📝 筆記文本 ({text_count:,})")
            self.btn_filter_media.configure(text=f"🖼️ 圖片影音 ({media_count:,})")
            self.btn_filter_code.configure(text=f"💻 程式腳本 ({code_count:,})")
            self.btn_filter_folder.configure(text=f"📁 資料夾 ({folder_count:,})")

            self.apply_filter()

            # If rescan was requested during active scan, execute now
            if self.scan_pending:
                self.scan_pending = False
                self.refresh_items()
        except Empty:
            pass
        self.after(100, self.check_scan_queue)

    def apply_filter(self):
        query = self.search_var.get().strip()
        self.clear_button.set_enabled(bool(query))

        items = self._get_navigation_items()

        # 1. Type filtering
        if self.active_type_filter == "word":
            items = [it for it in items if it.type_text == "Word"]
        elif self.active_type_filter == "excel":
            items = [it for it in items if it.type_text == "Excel"]
        elif self.active_type_filter == "pdf_ppt":
            items = [it for it in items if it.type_text in {"PDF", "PPT"}]
        elif self.active_type_filter == "text":
            items = [it for it in items if it.type_text == "Text"]
        elif self.active_type_filter == "media":
            items = [it for it in items if it.type_text == "Media"]
        elif self.active_type_filter == "code":
            items = [it for it in items if it.type_text == "Code"]
        elif self.active_type_filter == "folder":
            items = [it for it in items if it.is_dir]

        # 2. Time filtering (default is "all", which displays all items sorted by recency)
        today = datetime.now().date()
        if self.active_time_filter == "today":
            cutoff_today = datetime.combine(today, datetime.min.time()).timestamp()
            items = [it for it in items if it.modified_time >= cutoff_today]
        elif self.active_time_filter == "7d":
            cutoff_7d = datetime.combine(today - timedelta(days=7), datetime.min.time()).timestamp()
            items = [it for it in items if it.modified_time >= cutoff_7d]
        elif self.active_time_filter == "30d":
            cutoff_30d = datetime.combine(today - timedelta(days=30), datetime.min.time()).timestamp()
            items = [it for it in items if it.modified_time >= cutoff_30d]
        elif self.active_time_filter == "90d":
            cutoff_90d = datetime.combine(today - timedelta(days=90), datetime.min.time()).timestamp()
            items = [it for it in items if it.modified_time >= cutoff_90d]
        # "all": no time cutoff, show all items

        # 3. Fuzzy search matching
        if query:
            scored = []
            for item in items:
                score = self.fuzzy_score(query, item.name)
                if score:
                    scored.append((score, item))
            scored.sort(key=lambda row: (row[0], row[1].modified_time), reverse=True)
            items = [item for _score, item in scored]

        self.filtered_items = items
        self._sort_and_render(is_initial_fuzzy=bool(query and self.sort_col == "modified"))

    @staticmethod
    def fuzzy_score(query, name):
        query, name = query.casefold(), name.casefold()
        if query == name:
            return 1000
        if name.startswith(query):
            return 900
        if query in name:
            return 800
        pos = 0
        for char in name:
            if pos < len(query) and char == query[pos]:
                pos += 1
        if pos == len(query):
            return 600 + len(query) / max(len(name), 1)
        ratio = SequenceMatcher(None, query, name).ratio()
        return ratio * 500 if ratio >= 0.45 else 0

    def _sort_and_render(self, is_initial_fuzzy=False):
        """Sort filtered_items according to active sort column and reverse state."""
        items = list(self.filtered_items)
        if not is_initial_fuzzy:
            if self.sort_col == "type":
                items.sort(key=lambda it: it.type_text, reverse=self.sort_reverse)
            elif self.sort_col == "name":
                items.sort(key=lambda it: it.name.lower(), reverse=self.sort_reverse)
            elif self.sort_col == "modified":
                items.sort(key=lambda it: it.modified_time, reverse=self.sort_reverse)
            elif self.sort_col == "size":
                items.sort(key=lambda it: it.size_bytes, reverse=self.sort_reverse)
            elif self.sort_col == "path":
                items.sort(key=lambda it: str(it.path.parent).lower(), reverse=self.sort_reverse)

        display_items = items[:MAX_RESULTS]
        self.displayed_items = display_items

        # Update count badge text
        query = self.search_var.get().strip()
        time_desc = {
            "all": "全部時間",
            "today": "今天",
            "7d": "最近 7 天",
            "30d": "最近 30 天",
            "90d": "最近 90 天",
        }.get(self.active_time_filter, "全部時間")
        self.result_var.set(f"{time_desc} · 顯示 {len(display_items)} / {len(items)} 筆")

        self.render_items(display_items)

    def render_items(self, items):
        self.tree.delete(*self.tree.get_children())
        self.update_selection()

        for index, item in enumerate(items):
            tags = ["odd" if (index % 2) else "even"]
            tags.append(item.tag_name)

            self.tree.insert(
                "", tk.END, iid=str(index),
                values=(
                    item.type_badge,
                    item.name,
                    item.modified_text,
                    item.size_text,
                    str(item.path.parent)
                ),
                tags=tuple(tags)
            )

        if items:
            self.empty_overlay.place_forget()
        else:
            searching = bool(self.search_var.get().strip())
            if self.scan_running and not self.all_items:
                self.empty_icon.configure(text="⏳")
                self.empty_title.configure(text="正在建立索引...")
                self.empty_hint.configure(text="掃描完成後會自動顯示最近異動的檔案與資料夾。")
                self.empty_btn_box.pack_forget()
            elif not self.all_items and self.missing_roots:
                self.empty_icon.configure(text="📁")
                self.empty_title.configure(text="找不到已設定的掃描目錄")
                self.empty_hint.configure(text="目前的掃描路徑不存在於本機磁碟，請點擊下方按鈕加入資料夾。")
                self.empty_btn_box.pack()
            elif searching:
                self.empty_icon.configure(text="🔍")
                self.empty_title.configure(text="沒有找到符合名稱的項目")
                self.empty_hint.configure(text="試試縮短關鍵字，或按 Esc 清除搜尋以檢視全部檔案。")
                self.empty_btn_box.pack_forget()
            elif self.current_directory is not None:
                self.empty_icon.configure(text="📂")
                self.empty_title.configure(text="此資料夾沒有可顯示的項目")
                self.empty_hint.configure(text="已到目前索引範圍的最底層，請按「↩ 上層」返回。")
                self.empty_btn_box.pack_forget()
            else:
                self.empty_icon.configure(text="📅")
                self.empty_title.configure(text="該時間範圍內沒有異動項目")
                self.empty_hint.configure(text="點擊左側「30天」或在上方搜尋欄輸入名稱，可搜尋更早異動的檔案。")
                self.empty_btn_box.pack_forget()
            self.empty_overlay.place(relx=0.5, rely=0.5, anchor="center")

    def update_selection(self, _event=None):
        item = self.get_selected_item()
        has_item = (item is not None)
        is_file = (has_item and not item.is_dir)

        # Update sidebar duplicate button
        self.copy_button.set_enabled(is_file)

        # Update bottom bar buttons
        self.btn_bottom_open.set_enabled(has_item)
        self.btn_bottom_reveal.set_enabled(has_item)
        self.btn_bottom_copy.set_enabled(has_item)

        if item:
            size_part = f" ({item.size_text})" if not item.is_dir else ""
            self.selection_path_var.set(f"{item.type_badge} {item.name}{size_part} · {item.path.parent}")
        else:
            self.selection_path_var.set("選取項目可查看完整路徑與執行快捷操作")

    def get_selected_item(self):
        selected = self.tree.selection()
        if not selected:
            return None
        try:
            idx = int(selected[0])
            if hasattr(self, "displayed_items") and idx < len(self.displayed_items):
                return self.displayed_items[idx]
        except (ValueError, IndexError):
            pass
        return None

    def _show_context_menu(self, event):
        """Show context menu on right click and select target row."""
        row_id = self.tree.identify_row(event.y)
        if row_id:
            self.tree.selection_set(row_id)
            self.update_selection()
            self.context_menu.post(event.x_root, event.y_root)

    def open_selected(self):
        item = self.get_selected_item()
        if not item:
            return
        if item.is_dir:
            self.enter_selected_directory()
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(item.path))
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(item.path)])
            else:
                subprocess.Popen(["xdg-open", str(item.path)])
        except Exception as exc:
            messagebox.showerror("開啟失敗", f"無法開啟項目：\n{item.path}\n\n錯誤資訊：{exc}", parent=self)

    def reveal_in_explorer(self):
        """Reveal and select the item in Windows File Explorer or native file manager."""
        item = self.get_selected_item()
        if not item:
            return
        path_str = str(item.path.resolve())
        try:
            if sys.platform.startswith("win"):
                subprocess.Popen(["explorer", f"/select,{path_str}"])
            elif sys.platform == "darwin":
                subprocess.Popen(["open", "-R", path_str])
            else:
                subprocess.Popen(["xdg-open", str(item.path.parent)])
        except Exception as exc:
            messagebox.showerror("定位失敗", f"無法在檔案總管中定位：\n{exc}", parent=self)

    def open_parent_folder(self):
        item = self.get_selected_item()
        if not item:
            return
        folder = item.path if item.is_dir else item.path.parent
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(folder))
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(folder)])
            else:
                subprocess.Popen(["xdg-open", str(folder)])
        except Exception as exc:
            messagebox.showerror("開啟資料夾失敗", str(exc), parent=self)

    def copy_selected_path(self):
        item = self.get_selected_item()
        if not item:
            return
        path_str = str(item.path.resolve())
        self.clipboard_clear()
        self.clipboard_append(path_str)
        self.update()
        self.show_toast(f"已複製路徑：{item.name}")

    def copy_selected_name(self):
        item = self.get_selected_item()
        if not item:
            return
        self.clipboard_clear()
        self.clipboard_append(item.name)
        self.update()
        self.show_toast(f"已複製名稱：{item.name}")

    def duplicate_file(self):
        item = self.get_selected_item()
        if not item:
            messagebox.showinfo("尚未選取", "請先選擇要建立複本的檔案。", parent=self)
            return
        if item.is_dir:
            messagebox.showinfo("無法複製", "複本功能僅適用於檔案，不支援資料夾。", parent=self)
            return
        source = item.path
        target_text = filedialog.asksaveasfilename(
            title="以此為範本建立複本",
            initialdir=str(source.parent),
            initialfile=f"{source.stem}_複本{source.suffix}",
            defaultextension=source.suffix,
            parent=self
        )
        if not target_text:
            return
        target = Path(target_text)
        try:
            if source.resolve() == target.resolve():
                messagebox.showwarning("無法複製", "來源檔案與目標檔案路徑相同，請指定不同的檔名或資料夾。", parent=self)
                return
        except Exception:
            pass

        if target.exists() and not messagebox.askyesno("檔案已存在", f"{target.name}\n\n是否覆寫此檔案？", parent=self):
            return
        try:
            shutil.copy2(source, target)
            now = datetime.now().timestamp()
            os.utime(target, (now, now))
            self.show_toast(f"已成功建立複本：{target.name}")
            self.refresh_items()
        except Exception as exc:
            messagebox.showerror("複製失敗", f"複製檔案時發生錯誤：\n{exc}", parent=self)

    def create_word(self):
        self.create_office_file("Word", ".docx", "新文件.docx", self.write_blank_docx)

    def create_excel(self):
        self.create_office_file("Excel", ".xlsx", "新活頁簿.xlsx", self.write_blank_xlsx)

    def create_office_file(self, kind, suffix, initial_name, writer):
        target_text = filedialog.asksaveasfilename(
            title=f"新增 {kind} 檔案",
            initialdir=str(self.get_preferred_directory()),
            initialfile=initial_name,
            defaultextension=suffix,
            filetypes=[(f"{kind} 檔案", f"*{suffix}")],
            parent=self
        )
        if not target_text:
            return
        target = Path(target_text).with_suffix(suffix)
        if target.exists():
            messagebox.showwarning("檔案已存在", f"{target.name} 已經存在，請更換名稱。", parent=self)
            return
        try:
            writer(target)
            self.show_toast(f"已建立 {target.name}")
            self.refresh_items()
        except Exception as exc:
            messagebox.showerror(f"建立 {kind} 失敗", str(exc), parent=self)

    @staticmethod
    def write_blank_docx(target):
        files = {
            "[Content_Types].xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>''',
            "_rels/.rels": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>''',
            "word/document.xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>''',
        }
        with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
            for name, content in files.items():
                archive.writestr(name, content)

    @staticmethod
    def write_blank_xlsx(target):
        files = {
            "[Content_Types].xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>''',
            "_rels/.rels": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>''',
            "xl/workbook.xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>''',
            "xl/_rels/workbook.xml.rels": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>''',
            "xl/worksheets/sheet1.xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>''',
        }
        with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
            for name, content in files.items():
                archive.writestr(name, content)

    def get_preferred_directory(self):
        item = self.get_selected_item()
        if item:
            return item.path if item.is_dir else item.path.parent
        for root in self.scan_roots:
            p = Path(root)
            if p.exists():
                return p
        home_docs = Path.home() / "Documents"
        if home_docs.exists():
            return home_docs
        return Path.home()


def main():
    app = RecentOfficeFinder()
    app.mainloop()


if __name__ == "__main__":
    main()
