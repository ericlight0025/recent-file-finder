# Recent Office Finder (現代化桌面檔案與最近文件索引器)

A modern, high-performance Windows desktop app for quickly indexing and finding recently modified folders, documents, spreadsheets, images, media, notes, and code. It uses **only the Python standard library**; no `pip` packages are required.

---

## ✨ Features & UI / UX Highlights

### 🎨 Modern Dark Theme & Visual Polish
- **High-DPI Awareness**: Fully Per-Monitor DPI aware on Windows. Fonts, borders, and controls remain crisp and sharp without blurriness.
- **Windows Immersive Dark Title Bar**: Native OS window title bar integrates seamlessly with the dark palette via Windows DWM attributes.
- **Distinct Type Badges & Colors**: Quick visual recognition with clean icons and colors:
  - 📁 **資料夾 (Folder)** (Sky Blue)
  - 📄 **Word** (Soft Violet)
  - 📊 **Excel / CSV** (Emerald Green)
  - 📑 **PDF / 簡報** (Coral Red / Warm Orange)
  - 📝 **筆記文本** (Slate Gray)
  - 🖼️ **圖片影音** (Rose Pink)
  - 💻 **程式腳本** (Cyan Blue)
- **File Size Column**: Displays human-readable file sizes (`B`, `KB`, `MB`, `GB`) for quick inspection.
- **Alternating Zebra Rows**: Subtle alternating backgrounds improve scan readability.

### 🔍 Interactive Search & Filter Controls
- **Category Filter Buttons**: Instant one-click filtering for:
  - `● 全部項目`
  - `📄 Word`
  - `📊 Excel / CSV`
  - `📑 PDF / 簡報`
  - `📝 筆記文本`
  - `🖼️ 圖片影音`
  - `💻 程式腳本`
  - `📁 資料夾`
  (complete with live count counters for each category)
- **Time Scope Filter**: Choose between `全部 (All)`, `今天 (Today)`, `7天 (7 Days)`, `30天 (30 Days)`, or `90天 (90 Days)`.
- **Fuzzy Item Name Search**: Matches file and folder names dynamically as you type.
- **Instant Clear Button**: A recessed `✕ 清除` button resets the search box instantly.
- **Keyboard Navigation**:
  - `↓ (Down Arrow)` in the search box moves focus directly into the table.
  - `Enter` opens the selected item or first matching item with default system app.
  - `Esc` clears search or selection.
  - `Ctrl + F` focuses the search box.
  - `Ctrl + C` copies the full path of the selected item to clipboard.
  - `F5` refreshes the index.

### 📊 Sortable Columns
- Click on any table column header (`類型`, `名稱`, `修改時間`, `大小`, `所在目錄`) to sort ascending or descending.
- Visual sorting indicators (`▲` / `▼`) show active sort order.

### 🖱️ Right-Click Context Menu & Actions
Right-click any item in the table to access:
- 🚀 **開啟項目** (Open file or folder with system default app)
- 📂 **在檔案總管中定位** (Reveal in Windows File Explorer: highlights the file directly)
- 📁 **開啟所在資料夾** (Open parent directory)
- 📋 **複製完整路徑** (Copy full path to clipboard)
- 📝 **複製名稱** (Copy file/folder name)
- 📑 **以此建立複本...** (Duplicate file with automatic timestamp touch; protects against copying onto source)

### ⚙️ GUI Folder Configuration & Persistence
- Click `⚙️ 管理掃描目錄` in the sidebar or `＋ 快速加入目錄...` to manage indexing:
  - Add directories by pasting path directly or clicking `📁 瀏覽...`.
  - Automatically trims enclosing quotes (supporting Windows "Copy as path").
  - Auto-commits pending input when clicking `💾 儲存並重新掃描`.
  - Shows per-directory item count breakdown (e.g. `✅ 存在 (18,188 筆)`).
  - Choose **時間範圍**: `不限時間 (全部檔案)`、`最近 30 天`、`最近 90 天` 等。
  - Choose **檔案類型範圍**:
    - `全部常用 (Office/圖片/影音/程式碼)` [預設]
    - `僅 Office 與文件`
    - `不限副檔名 (*.*)`
- Configurations are automatically saved to `finder_config.json` next to the script.

### 📄 Quick Blank Office Document Generation
- `＋ 新增 Word 文件`: Generates a valid standard `.docx` directly without external tools.
- `＋ 新增 Excel 活頁簿`: Generates a valid standard `.xlsx` directly without external tools.

### 🛠️ 其他優化規劃（目前未啟用）
以下項目先記錄為後續優化方向，尚未加入目前穩定版本，以避免啟動與掃描流程變得過重：
- SQLite 索引快取與啟動時快取驗證。
- 可取消掃描、掃描進度與更細緻的掃描狀態提示。
- Explorer 風格的進階搜尋語法（類型、副檔名、日期、檔案大小）。
- 自訂排除資料夾與副檔名規則。
- 目前結果匯出 CSV、文字檔預覽、直接輸入路徑跳轉。
- 掃描目錄自動變更偵測與介面狀態保存。

---

## 💻 Requirements

- Python 3.9 or newer recommended (tested with Python 3.12).
- Windows with Tkinter enabled (included with standard Windows Python installer).
- Zero third-party dependencies (`no pip install` required).

---

## 🚀 How to Run

From this folder, run:

```powershell
python recent_office_finder.py
```

---

## 📁 Configuration File (`finder_config.json`)

The application automatically creates and manages `finder_config.json` in the same directory:

```json
{
  "scan_roots": [
    "C:\\Users\\your_name\\Documents",
    "C:\\DevWorkspace\\your_project"
  ],
  "search_days": 0,
  "file_scope": "common"
}
```

You can either modify this JSON file directly or use the in-app **⚙️ 管理掃描目錄** interface.
