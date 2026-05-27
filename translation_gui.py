#!/usr/bin/env python3
"""
MUI Translation Tool — GUI Version
雙擊執行，選 zip + Excel，按 Run 即可。
"""

import json, re, sys, zipfile, threading
from pathlib import Path
from collections import defaultdict, OrderedDict
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

APP_VERSION  = "v1.7.BUILD_DATETIME"   # replaced by build script at package time
APP_AUTHOR   = "Pocky Wu"
TOOL_VERSION = "5"   # bump when index structure changes (forces cache rebuild)

CHANGELOG = """\
v1.7
────────────────────────────────────────
新功能
  • 新增「轉換 Ignore」模式：一鍵將語言掃描報告（含 <語言> Extracted Sheet）轉換為 Ignore Excel
    → 自動偵測 Sheet 名稱（格式：FIL Extracted / HEB Extracted / POL Extracted …）
    → Comment = Pass 的字串按 Sheet 所屬語言收錄，不混用「語言欄」的多語言值
    → 同一 Key 在多個語言都 Pass 時自動合併為同一列（語言欄逗號分隔）
    → 多個 Key 同格（換行分隔）自動拆分為獨立列

────────────────────────────────────────
v1.6
────────────────────────────────────────
UI 優化
  • 語言全掃描模式：輸入／輸出 Excel 欄位自動灰化，避免使用者混淆
  • 語言全掃描模式：輸出報告「選取」按鈕不再被截斷（改用彈性欄寬）
  • 移除「部分 Key 未翻譯」分類，改統一顯示為「未翻譯 / 空白」，避免使用者混淆

新功能
  • Ignore Excel 語言欄支援 ALL_LANGUAGES，表示該字串所有語言皆與英文相同，直接跳過不報錯

Bug 修正
  • 修正特定翻譯字串含 \x10 等控制字元導致輸出 Excel 時崩潰（IllegalCharacterError）
    → 寫入前自動清除 Excel XML 不合法字元，不影響翻譯內容完整度

────────────────────────────────────────
v1.5
────────────────────────────────────────
新功能
  • 新增 YCP、YCVB 兩個 App（YCP 支援 20 語言含希臘/波蘭/羅馬尼亞，YCVB 16 語言）
  • Android XML parser 新增 <plurals> 支援，以 quantity="other" 作為代表值

UI 優化
  • App 選擇按鈕格式改為 YMK(20)，縮短寬度
  • 輸出 Excel 檔名欄寬加大（32 字元），減少截斷

────────────────────────────────────────
v1.4
────────────────────────────────────────
新功能
  • 更新記錄：Header 新增「📋 更新記錄」入口，可查看每版變更
  • 自動發佈腳本：build 完可一鍵上傳 exe 至 Google Drive 並更新版號
  • 自動更新提示：app 啟動時背景檢查是否有新版（版號格式改為 1.4.yyMMdd_HHmm）

UI 優化
  • 全介面字型放大兩號
  • Log 區域撐滿視窗底部，不再留白也不出現多餘捲動條
  • 檔名過長時截斷顯示，前綴 … 保留尾端資訊
  • 欄位標籤加冒號，「Zip 檔」改為「翻譯 Zip 檔:」
  • 語言勾選區塊字型縮小兩號、改為 4 欄，語言代碼完整顯示不截斷

Bug 修正
  • 修正 fuzzy match 時翻譯與英文相同未標紅（如 Choose 1-{{count}} styles）
  • 修正「輸出報告」選取鈕超出視窗右邊界

建置
  • build_windows.bat 改用 PowerShell 取時間戳，修正 Windows 11 移除 wmic 後無法 build 的問題

────────────────────────────────────────
v1.3
────────────────────────────────────────
新功能
  • 取消按鈕：Run / 開始掃描執行中可隨時中止
  • 進度條百分比：顯示「處理第 N / M」並同步更新進度條
  • 快速查詢獨立結果框：查詢結果不再混入 log，右欄獨立顯示
  • 語言全掃描記憶：切換 App 或重開程式自動帶入上次勾選的語言
  • 掃描輸出路徑記憶：與 Zip / Excel 路徑同步儲存於 config
  • 自動更新提示：啟動時背景檢查 Google Drive 是否有新版

Bug 修正
  • 修正 Android gen/ 資料夾的中文字串被誤當英文 base
  • 修正 .strings 檔解析 regex typo 導致部分字串漏抓
  • 修正 fuzzy match 情況下翻譯與英文相同未標紅（Choose 1-{{count}} styles 等）
  • 修正語言全掃描面板「輸出報告」選取鈕超出視窗邊界
  • 修正捲動條在內容未超出視窗時仍可捲動的問題

效能 / 架構
  • 索引快取邏輯重構，三個模式共用同一 _load_index()
  • _normalize() 改用 module-level 預編譯 regex，加速重複呼叫
  • UI 操作全面改用 self.after() 確保 thread safety

────────────────────────────────────────
v1.2
────────────────────────────────────────
  • 新增語言全掃描模式（不需 Excel，直接掃描整包 Zip）
  • 新增快速查詢（右側文字框，Ctrl+Enter 執行）
  • 新增 Ignore Excel 支援（跳過預期與英文相同的字串）
  • 翻譯問題報告新增「部分 Key 未翻譯」橘色分類
  • 模糊比對支援（數字 ↔ 變數 正規化，橘色字標示）
  • 路徑設定自動記憶，切換 App 各自保留上次路徑
"""

# ── Auto-update ───────────────────────────────────────────────────────────────
# latest_version.txt on Google Drive (public, "Anyone with link can view")
# Content format: full version with build stamp, e.g.  1.3.260508_1030
_UPDATE_CHECK_URL    = "https://drive.google.com/uc?export=download&id=11dRxedUqKwtD26P4EcNK17qmICefJ4-w"
_UPDATE_DOWNLOAD_URL = "https://drive.google.com/drive/folders/1-DumwCUoqOV-zgMRhUtI4QRMtqGn3qDq"
try:
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    import pandas as pd
except ImportError:
    import subprocess, sys
    root = tk.Tk(); root.withdraw()
    if messagebox.askyesno("缺少套件", "需要安裝 openpyxl 和 pandas，是否立即安裝？"):
        subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl", "pandas"])
        messagebox.showinfo("完成", "安裝完成，請重新啟動程式。")
    sys.exit(0)

# ┌─────────────────────────────────────────────────────────────────────────┐
# │  APP_CONFIGS — 每個 App 的語言設定                                        │
# │  新增 App:  在下面複製一個 block，改名稱和語言清單                          │
# │  新增語言:  在對應 App 的 dict 裡加一行 "代碼": "lang code"                │
# │  設為 None = 輸出全部語言（不過濾）                                        │
# └─────────────────────────────────────────────────────────────────────────┘
APP_CONFIGS: dict[str, dict | None] = {

    "YMK": {
        "ENU": "en",
        "CHT": "zh-Hant",
        "CHS": "zh-Hans",
        "JPN": "ja",
        "KOR": "ko",
        "ESP": "es",
        "ARA": "ar",
        "FRA": "fr",
        "DEU": "de",
        "RUS": "ru",
        "PTB": "pt",
        "IND": "id",
        "THA": "th",
        "TRK": "tr",
        "MSL": "ms",
        "ITA": "it",
        "NLD": "nl",
        "HEB": "he",
        "POL": "pl",
        "FIL": "fil",
    },

    "YCE": {
        "ENU": "en",
        "CHT": "zh-Hant",
        "CHS": "zh-Hans",
        "JPN": "ja",
        "KOR": "ko",
        "RUS": "ru",
        "DEU": "de",
        "FRA": "fr",
        "THA": "th",
        "ESP": "es",
        "PTB": "pt",
        "IND": "id",
        "TRK": "tr",
        "MSL": "ms",
        "NLD": "nl",
        "ITA": "it",
        "HUM": "hu",
        "SWE": "sv",
        "ARA": "ar",
    },

    "YCP": {
        "ENU": "en",
        "CHT": "zh-Hant",
        "CHS": "zh-Hans",
        "JPN": "ja",
        "KOR": "ko",
        "DEU": "de",
        "ESP": "es",
        "FRA": "fr",
        "ITA": "it",
        "RUS": "ru",
        "PTB": "pt",
        "IND": "id",
        "THA": "th",
        "TRK": "tr",
        "MSL": "ms",
        "NLD": "nl",
        "ELL": "el",
        "PLO": "pl",
        "RON": "ro",
        "ARA": "ar",
    },

    "YCVB": {
        "ENU": "en",
        "CHT": "zh-Hant",
        "CHS": "zh-Hans",
        "JPN": "ja",
        "KOR": "ko",
        "DEU": "de",
        "ESP": "es",
        "FRA": "fr",
        "ITA": "it",
        "RUS": "ru",
        "PTB": "pt",
        "IND": "id",
        "THA": "th",
        "TRK": "tr",
        "MSL": "ms",
        "NLD": "nl",
    },

}

REGIONAL_ALLOWLIST: set = set()

# ── Language metadata ─────────────────────────────────────────────────────────
LANG_NAMES = {
    "en":      "英語",
    "zh-Hant": "繁體中文",  "zh-Hans": "簡體中文",  "zh": "中文",
    "ja": "日語",           "ko": "韓語",           "ar": "阿拉伯語",
    "de": "德語",           "es": "西班牙語",        "fr": "法語",
    "it": "義大利語",       "nl": "荷蘭語",          "pt": "葡萄牙語",
    "pt-PT": "葡萄牙語(葡)", "ru": "俄語",           "th": "泰語",
    "tr": "土耳其語",       "id": "印尼語",          "ms": "馬來語",
    "fil": "菲律賓語",      "he": "希伯來語",        "pl": "波蘭語",
    "sv": "瑞典語",         "da": "丹麥語",          "hu": "匈牙利語",
    "el": "希臘語",         "uk": "烏克蘭語",        "bg": "保加利亞語",
    "hr": "克羅埃西亞語",   "ro": "羅馬尼亞語",      "cs": "捷克語",
}
LANG_PRIORITY = ["en", "zh-Hant", "zh-Hans", "zh", "ja", "ko"]

# ── Styles ────────────────────────────────────────────────────────────────────
YELLOW  = PatternFill("solid", fgColor="FFD700")
YELLOW2 = PatternFill("solid", fgColor="FFF176")   # 2nd longest when ARA is longest
RED     = PatternFill("solid", fgColor="FF6B6B")
HEADER = PatternFill("solid", fgColor="302B63")
H_FONT = Font(bold=True, color="FFFFFF", name="Arial", size=10)
D_FONT = Font(name="Arial", size=10)
B_FONT = Font(name="Arial", size=10, bold=True)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT   = Alignment(horizontal="left",   vertical="top",    wrap_text=True)
thin   = Side(style="thin", color="DDDDDD")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

# ── Config persistence (#6) ───────────────────────────────────────────────────
import tempfile as _tempfile
_CONFIG_PATH = Path(_tempfile.gettempdir()) / "TranslationTool" / "config.json"

def load_config() -> dict:
    try:
        if _CONFIG_PATH.exists():
            return json.loads(_CONFIG_PATH.read_text("utf-8"))
    except Exception:
        pass
    return {}

def save_config(cfg: dict):
    try:
        _CONFIG_PATH.parent.mkdir(exist_ok=True)
        _CONFIG_PATH.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass

# ── Parsers ───────────────────────────────────────────────────────────────────

def _is_regional_variant(lang: str) -> bool:
    if lang in REGIONAL_ALLOWLIST: return False
    m = re.match(r'^[a-z]{2,3}-([A-Z][a-z]{3}|[A-Z]{2})$', lang)
    if not m: return False
    return len(m.group(1)) == 2

def parse_arb(text: str) -> tuple:
    try: data = json.loads(text)
    except: return "", {}
    loc  = data.get("@@locale") or data.get("localeCode")
    lang = loc.replace("_", "-") if loc else ""
    return lang, {k: v for k, v in data.items()
                  if not k.startswith("@") and isinstance(v, str)}

def parse_strings(text: str) -> dict:
    text = re.sub(r'/\*.*?\*/', '', text, flags=re.DOTALL)
    entries = {}
    for line in text.splitlines():
        s = line.strip()
        if not s or s.startswith('//'): continue
        m = re.match(r'^"((?:[^"\\]|\\.)*)"[ \t]*=[ \t]*"((?:[^"\\]|\\.)*)"\s*;', s)
        if m: entries[m.group(1)] = m.group(2)
    return entries

def lang_from_arb_name(stem: str) -> str:
    m = re.search(r'_([a-z]{2}(?:_[A-Za-z]{2,4})*)$', stem)
    return m.group(1).replace("_", "-") if m else ""

def lang_from_strings_name(name: str) -> str:
    m = re.match(r'.+?_(.+?)(?:\.lproj)?\.strings$', name)
    return m.group(1).replace("_", "-") if m else ""

ANDROID_LANG_MAP = {
    "zh-rCN": "zh-Hans", "zh-rHK": "zh-Hant", "zh-rTW": "zh-Hant",
    "in": "id", "iw": "he", "pt-rBR": "pt", "nl-rNL": "nl", "nl-rBE": "nl",
}

def lang_from_android_folder(folder: str) -> str:
    if folder == "values": return "en"
    m = re.match(r'^values-(.+)$', folder)
    if not m: return ""
    raw = m.group(1)
    if re.match(r'^v\d+$', raw): return ""
    if raw in ANDROID_LANG_MAP: return ANDROID_LANG_MAP[raw]
    return re.sub(r'-r([A-Z]{2})$', r'-\1', raw)

def parse_android_xml(text: str) -> dict:
    import html as _html

    def _clean(raw: str) -> str:
        raw = raw.strip()
        raw = raw.replace("\\'", "'").replace('\\"', '"')
        raw = raw.replace("\\n", "\n").replace("\\t", "\t")
        raw = re.sub(r'<!\[CDATA\[(.*?)\]\]>', r'\1', raw, flags=re.DOTALL)
        raw = re.sub(r'<[^>]+>', '', raw)
        return _html.unescape(raw).strip()

    entries = {}
    for m in re.finditer(r'<string\s+name="([^"]+)"[^>]*>(.*?)</string>', text, re.DOTALL):
        val = _clean(m.group(2))
        if m.group(1) and val:
            entries[m.group(1)] = val

    # <plurals>: use "other" quantity as representative value
    for m in re.finditer(r'<plurals\s+name="([^"]+)"[^>]*>(.*?)</plurals>', text, re.DOTALL):
        key, block = m.group(1), m.group(2)
        item = re.search(r'<item\s+quantity="other"[^>]*>(.*?)</item>', block, re.DOTALL)
        if not item:
            item = re.search(r'<item[^>]*>(.*?)</item>', block, re.DOTALL)
        if item:
            val = _clean(item.group(1))
            if key and val:
                entries[key] = val

    return entries

GARBLED_RE = re.compile(r'[\ufffd\x00-\x08\x0b\x0c\x0e-\x1f]')

# Characters illegal in Excel XML (openpyxl will raise IllegalCharacterError)
_XL_ILLEGAL_RE = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff]')

def _xl_safe(s: str) -> str:
    """Strip characters that openpyxl cannot write to Excel cells."""
    if not isinstance(s, str): return s
    return _XL_ILLEGAL_RE.sub('', s)

def is_garbled(val: str) -> bool:
    if not val or not val.strip(): return True
    if GARBLED_RE.search(val): return True
    if len(val) > 3 and val.count('?') > len(val) * 0.4: return True
    return False

def _visual_len(s: str) -> int:
    """
    Count visible characters only, excluding Unicode combining marks
    (e.g. Thai tone marks ่ ้ ๊ ๋, Arabic diacritics ً ٌ ٍ etc.)
    so that 'หลีกเลี่ยง' and 'Vermeiden' compare by display width.
    """
    import unicodedata
    return sum(1 for c in s if unicodedata.category(c) not in ('Mn', 'Mc', 'Me'))


_RE_PRINTF  = re.compile(r'%\d+\$[@disfeEgGuoxXld]+')
_RE_PRINTF2 = re.compile(r'%[@disfeEgGuoxXld]+')
_RE_BRACE   = re.compile(r'\{[^}]+\}')
_RE_NUM     = re.compile(r'(?<!\w)\d+(?!\w)')
_RE_SPACE   = re.compile(r'\s+')

def _normalize(s: str) -> str:
    s = s.replace('\\n', ' ').replace('\n', ' ')
    s = _RE_PRINTF.sub('__X__', s)
    s = _RE_PRINTF2.sub('__X__', s)
    s = _RE_BRACE.sub('__X__', s)
    s = _RE_NUM.sub('__X__', s)
    return _RE_SPACE.sub(' ', s).strip().lower()

# ── Index builder ─────────────────────────────────────────────────────────────

def build_index(source: Path, base_lang: str, log) -> tuple:
    log(f"🔍 掃描翻譯來源: {source.name}")
    projects: dict = defaultdict(dict)

    def add(proj, lang, kvs):
        if proj and lang and kvs:
            if lang in projects[proj]: projects[proj][lang].update(kvs)
            else: projects[proj][lang] = dict(kvs)

    def process_file(proj, fname, raw, folder=""):
        if fname.endswith(".arb"):
            lc, kvs = parse_arb(raw)
            lang = lc or lang_from_arb_name(Path(fname).stem)
            add(proj, lang, kvs)
        elif fname.endswith(".strings"):
            add(proj, lang_from_strings_name(fname), parse_strings(raw))
        elif fname == "strings.xml" and folder:
            lang = lang_from_android_folder(folder)
            if lang: add(proj, lang, parse_android_xml(raw))

    if zipfile.is_zipfile(source):
        with zipfile.ZipFile(source) as zf:
            for info in zf.infolist():
                if info.is_dir(): continue
                if "/gen/" in info.filename.replace("\\", "/"): continue
                parts = Path(info.filename).parts
                if len(parts) < 2: continue
                fname, folder = parts[-1], parts[-2]
                raw = zf.read(info).decode("utf-8", errors="replace")
                if fname == "strings.xml":
                    proj = parts[1] if len(parts) >= 7 else parts[0]
                    process_file(proj, fname, raw, folder)
                else:
                    process_file(parts[-2], fname, raw)
    elif source.is_dir():
        for pd_ in sorted(source.iterdir()):
            if not pd_.is_dir(): continue
            for fp in sorted(pd_.rglob("*")):
                if fp.is_dir(): continue
                try: raw = fp.read_text("utf-8", errors="replace")
                except: continue
                process_file(pd_.name, fp.name, raw, fp.parent.name)

    lookup: dict = {}
    for proj, lang_data in projects.items():
        base_kvs = lang_data.get(base_lang, {})
        if not base_kvs: continue

        # First pass: group all keys by their en_val
        # key_groups[en_val] = [ (key, {lang: val}) ]
        key_groups: dict[str, list] = {}
        for key, en_val in base_kvs.items():
            if not en_val.strip(): continue
            pt = {}
            for lang, kvs in lang_data.items():
                if lang == base_lang: continue
                val = kvs.get(key, "")
                if val: pt[lang] = val
            key_groups.setdefault(en_val, []).append((key, pt))

        # Second pass: merge per en_val
        for en_val, key_list in key_groups.items():
            # All langs that exist in any key's translations
            all_langs = set()
            for _, pt in key_list:
                all_langs.update(pt.keys())

            # Best translation per lang (from whichever key has it)
            best: dict = {base_lang: en_val}
            for lang in all_langs:
                for _, pt in key_list:
                    val = pt.get(lang, "")
                    if val:
                        best[lang] = val
                        break

            # Track which keys are missing each lang
            key_missing: dict[str, list] = {}   # lang → [missing key names]
            if len(key_list) > 1:
                # Only worth tracking when multiple keys share the same en_val
                for lang in all_langs:
                    missing_keys = [k for k, pt in key_list if not pt.get(lang)]
                    if missing_keys:
                        key_missing[lang] = missing_keys

            # Also flag langs that are missing from ALL keys
            for lang in all_langs:
                if not best.get(lang):
                    key_missing.setdefault(lang, [k for k, _ in key_list])

            best["_key_missing"] = key_missing
            best["_all_keys"]    = [k for k, _ in key_list]
            best["_key_trans"]   = {k: dict(pt) for k, pt in key_list}  # per-key translations
            lookup.setdefault(en_val, {})[proj] = best

    norm_index: dict = {}
    for en_val in lookup:
        nk = _normalize(en_val)
        if nk not in norm_index:
            norm_index[nk] = en_val

    log(f"✅ 索引完成: {len(projects)} 個 project，{len(lookup)} 個英文字串")
    return lookup, norm_index

# ── Ignore list loader ────────────────────────────────────────────────────────

def load_ignore_list(path: Path, target_langs: dict | None = None) -> set:
    """
    Parse ignore Excel with columns: Module | ENU翻譯檔字串 | 問題Key | 語言
    語言 column contains comma-separated lang codes (e.g. "HEB, POL, FIL")
    using the APP_CONFIGS keys (e.g. HEB→he, POL→pl, FIL→fil).
    Special value "ALL_LANGUAGES" skips the string for every language.
    Returns set of (module, key_name, lang_code) to skip in reports.
    Entries with ALL_LANGUAGES use lang_code="*" as a wildcard sentinel.
    """
    if not path or not path.exists():
        return set()

    import pandas as pd_
    try:
        df = pd_.read_excel(str(path), header=0)
    except Exception:
        return set()

    # Build code→lang mapping e.g. {"HEB": "he", "POL": "pl"}
    code_to_lang = {k: v for k, v in (target_langs or {}).items()}

    ignore: set = set()
    for _, row in df.iterrows():
        module    = str(row.iloc[0]).strip() if pd_.notna(row.iloc[0]) else ""
        key_name  = str(row.iloc[2]).strip() if pd_.notna(row.iloc[2]) else ""
        langs_str = str(row.iloc[3]).strip() if pd_.notna(row.iloc[3]) else ""

        if not module or not key_name or not langs_str:
            continue

        if langs_str.upper() == "ALL_LANGUAGES":
            ignore.add((module, key_name, "*"))
            continue

        for code in [c.strip() for c in langs_str.split(",")]:
            lang = code_to_lang.get(code, "")
            if lang:
                ignore.add((module, key_name, lang))

    return ignore


def _in_ignore(ignore_set: set, module: str, key: str, lang: str) -> bool:
    """Check ignore_set with wildcard support for ALL_LANGUAGES entries."""
    return (module, key, lang) in ignore_set or (module, key, "*") in ignore_set


# ── Scan-report → Ignore-table converter ─────────────────────────────────────

def _detect_lang_sheets(xl: "pd.ExcelFile") -> dict:
    """
    Return {sheet_name: LANG_CODE} for every language-analysis sheet.
    Detects by finding the first 2–5 consecutive uppercase letters in the
    sheet name (e.g. "FIL Extracted", "FIL_Extraction", "HEB_Extraction").
    Sheet must also contain a Comment column to avoid false positives.
    """
    result = {}
    for name in xl.sheet_names:
        m = re.search(r'[A-Z]{2,5}', name)
        if not m:
            continue
        try:
            header = pd.read_excel(xl, sheet_name=name, nrows=0)
        except Exception:
            continue
        cols = [str(c).strip() for c in header.columns]
        if len(cols) >= 6 and "Comment" in cols:
            result[name] = m.group(0)
    return result


def convert_scan_to_ignore(input_xlsx: Path, output_xlsx: Path, log) -> int:
    """
    Convert a language scan report to an ignore table Excel.
    Auto-detects language sheets via column names or sheet name pattern;
    for each Pass row assigns only that sheet's own language.
    Returns number of rows written.
    """
    xl = pd.ExcelFile(str(input_xlsx))

    lang_sheets = _detect_lang_sheets(xl)

    if not lang_sheets:
        raise ValueError(
            "找不到語言分析 Sheet。\n"
            "支援的格式（擇一即可）：\n"
            "  • 欄位名稱含「<語言> Translation」，例如：FIL Translation\n"
            "  • Sheet 名稱以語言代碼開頭，例如：FIL Extracted、FIL_Extraction"
        )

    log(f"🔍 偵測到語言 Sheet: {list(lang_sheets.values())}")

    records: dict[tuple, set] = {}
    for sheet, lang in lang_sheets.items():
        df = pd.read_excel(xl, sheet_name=sheet)
        if df.shape[1] < 6:
            log(f"⚠️  '{sheet}' 欄位不足，跳過")
            continue
        pass_rows = df[df.iloc[:, 5].astype(str).str.strip().str.lower() == "pass"]
        count = 0
        for _, row in pass_rows.iterrows():
            module   = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            enu      = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
            keys_raw = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ""
            for key in keys_raw.split("\n"):
                key = key.strip()
                if key:
                    records.setdefault((module, enu, key), set()).add(lang)
                    count += 1
        log(f"  {lang}: {len(pass_rows)} 列 Pass（{count} 個 Key）")

    rows = []
    for (module, enu, key), langs in sorted(records.items()):
        rows.append([module, enu, key, ", ".join(sorted(langs))])

    wb = Workbook()
    ws = wb.active
    ws.title = "Ignore Table"

    hdr_fill  = PatternFill("solid", fgColor="4472C4")
    hdr_font  = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    thin_s    = Side(style="thin", color="BFBFBF")
    cell_bdr  = Border(left=thin_s, right=thin_s, top=thin_s, bottom=thin_s)
    alt_fill  = PatternFill("solid", fgColor="EEF2FA")
    data_font = Font(name="Arial", size=10)

    for c, h in enumerate(["Module", "ENU翻譯檔字串", "問題Key", "語言"], 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font      = hdr_font
        cell.fill      = hdr_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border    = cell_bdr

    for r, row_data in enumerate(rows, 2):
        fill = alt_fill if r % 2 == 0 else None
        for c, val in enumerate(row_data, 1):
            cell = ws.cell(row=r, column=c, value=val)
            cell.font      = data_font
            cell.alignment = Alignment(vertical="center")
            cell.border    = cell_bdr
            if fill:
                cell.fill = fill

    ws.row_dimensions[1].height = 20
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 48
    ws.column_dimensions["D"].width = 18
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:D{len(rows) + 1}"

    wb.save(str(output_xlsx))
    log(f"✅ 輸出完成: {output_xlsx.name}（共 {len(rows)} 筆）")
    return len(rows)


# ── Language scan report (scan mode) ─────────────────────────────────────────

def generate_scan_report(index: dict, output_xlsx: Path, log,
                         scan_langs: list[str],
                         target_langs: dict | None = None,
                         ignore_set: set = None,
                         cancel_event=None):
    """
    Scan ALL strings in the index for selected languages.
    Output: one report sheet grouped by Module.
    """
    lang_label = {v: k for k, v in target_langs.items()} if target_langs else {}

    log(f"🔎 掃描語言: {[lang_label.get(l, l) for l in scan_langs]}")

    groups: dict[tuple, list] = OrderedDict()

    all_records = []
    for en_val, proj_dict in index.items():
        for proj, trans in proj_dict.items():
            all_records.append((proj, en_val, trans))

    total = len(all_records)
    for i, (module, en_val, trans) in enumerate(all_records):
        if cancel_event and cancel_event.is_set():
            log("⚠️  已取消"); return
        if (i + 1) % 200 == 0 or i == total - 1:
            log(f"  掃描第 {i+1} / {total} 筆...")

        enu_val     = trans.get("en", "")
        key_missing = trans.get("_key_missing", {})
        all_keys    = trans.get("_all_keys", [])

        for lang in scan_langs:
            if lang == "en": continue
            val   = trans.get(lang, "")
            label = lang_label.get(lang, lang)
            missing_keys = key_missing.get(lang, [])  # keys specifically missing this lang

            if not val or is_garbled(val):
                issue    = "未翻譯 / 空白"
                keys_str = "\n".join(all_keys)
            elif enu_val and val.strip() == enu_val.strip():
                # Check ignore: pass if ALL keys for this string are in ignore_set
                if ignore_set and all(
                    _in_ignore(ignore_set, module, k, lang) for k in (all_keys or [""])
                ):
                    continue
                issue    = "與英文翻譯檔相同"
                keys_str = "\n".join(all_keys)
            elif missing_keys:
                issue    = "未翻譯 / 空白"
                keys_str = "\n".join(missing_keys)
            else:
                continue

            grp_key = (module, enu_val, issue, keys_str)
            if grp_key not in groups:
                groups[grp_key] = []
            if label not in groups[grp_key]:
                groups[grp_key].append(label)

    sorted_groups = sorted(groups.items(), key=lambda x: (x[0][0], x[0][2], x[0][1]))
    log(f"📋 發現 {len(sorted_groups)} 筆問題")

    ISSUE_COLORS = {
        "未翻譯 / 空白":    "FF6B6B",
        "與英文翻譯檔相同": "FFD740",
    }

    wb = Workbook()
    rpt = wb.active
    rpt.title = "語言掃描報告"
    RPT_H_FILL = PatternFill("solid", fgColor="C0392B")

    for ci, h in enumerate(["Module", "ENU 翻譯檔字串", "問題 Key", "語言", "問題類型"], 1):
        c = rpt.cell(row=1, column=ci, value=h)
        c.fill = RPT_H_FILL
        c.font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
        c.alignment, c.border = CENTER, BORDER
    rpt.row_dimensions[1].height = 24

    prev_module = None
    for ri, ((module, enu_val, issue, keys_str), langs_list) in enumerate(sorted_groups, 2):
        is_new = (module != prev_module)
        prev_module = module
        row_border = Border(
            left=thin, right=thin,
            top=Side(style="medium" if is_new else "thin",
                     color="C0392B" if is_new else "DDDDDD"),
            bottom=thin
        )
        row_h = max(15, len(keys_str.split("\n")) * 14) if keys_str else 15
        rpt.row_dimensions[ri].height = row_h

        for ci, v in enumerate([module, enu_val, keys_str, ", ".join(langs_list), issue], 1):
            c = rpt.cell(row=ri, column=ci, value=_xl_safe(v) or None)
            c.border = row_border
            if ci == 5:
                c.fill = PatternFill("solid", fgColor=ISSUE_COLORS.get(issue, "FFFFFF"))
                c.font = Font(name="Arial", size=10, bold=True)
                c.alignment = LEFT
            elif ci == 3:
                c.font = Font(name="Arial", size=10)
                c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            else:
                c.font = Font(name="Arial", size=10)
                c.alignment = LEFT

    rpt.column_dimensions["A"].width = 40   # Module
    rpt.column_dimensions["B"].width = 44
    rpt.column_dimensions["C"].width = 36
    rpt.column_dimensions["D"].width = 32
    rpt.column_dimensions["E"].width = 22
    rpt.freeze_panes = "A2"
    _auto_row_height(rpt, {"A": 40, "B": 44, "C": 36, "D": 32, "E": 22}, content_cols={"B", "C"})

    wb.save(str(output_xlsx))
    log(f"✅ 輸出完成: {output_xlsx.name}")


def _auto_row_height(ws, col_widths: dict, content_cols: set = None,
                     min_h: int = 15, line_h: int = 15):
    """
    Set row heights based on cell content and column width.
    content_cols: set of column letters to include in calculation.
                  If None, use all columns.
    Uses a 0.7 correction factor for proportional fonts.
    """
    from math import ceil
    FONT_FACTOR = 0.7   # proportional font avg char width vs Excel width unit
    for row in ws.iter_rows(min_row=2):
        max_lines = 1
        for cell in row:
            if not cell.value: continue
            col_letter = get_column_letter(cell.column)
            if content_cols and col_letter not in content_cols:
                continue
            col_w   = col_widths.get(col_letter, 20)
            eff_w   = col_w / FONT_FACTOR   # effective character capacity
            text    = str(cell.value)
            lines   = 0
            for part in text.split("\n"):
                lines += max(1, ceil(len(part) / max(eff_w, 1)))
            max_lines = max(max_lines, lines)
        ws.row_dimensions[row[0].row].height = max(min_h, max_lines * line_h)


# ── Excel generator ───────────────────────────────────────────────────────────

def generate_excel(index: dict, norm_index: dict, input_xlsx: Path,
                   output_xlsx: Path, log, target_langs: dict | None = None,
                   ignore_set: set = None, cancel_event=None):
    import pandas as pd_
    df       = pd_.read_excel(str(input_xlsx), header=0)
    page_col = df.columns[0]
    en_col   = df.columns[1]
    total    = len(df)

    rows_data = []
    seen_langs: set = set()

    for i, (_, row) in enumerate(df.iterrows()):
        if cancel_event and cancel_event.is_set():
            log("⚠️  已取消"); return
        en   = str(row[en_col]).strip()   if pd_.notna(row[en_col])   else ""
        page = str(row[page_col]).strip() if pd_.notna(row[page_col]) else ""

        if (i + 1) % 5 == 0 or i == total - 1:
            log(f"  處理第 {i+1} / {total} 筆...")

        rec = index.get(en)
        if not rec and en:
            for k, v in index.items():
                if k.lower() == en.lower(): rec = v; break

        matched_key = en
        if not rec and en:
            nq = _normalize(en)
            orig = norm_index.get(nq)
            if orig:
                rec = index.get(orig)
                matched_key = orig

        if rec:
            first_proj = sorted(rec.keys())[0]
            for proj, proj_trans in sorted(rec.items()):
                all_keys  = proj_trans.get("_all_keys", [])
                key_trans = proj_trans.get("_key_trans", {})
                seen_langs.update(k for k in proj_trans if not k.startswith("_"))

                if len(all_keys) > 1:
                    # Multiple keys → one row per key
                    key_missing = proj_trans.get("_key_missing", {})
                    for ki, key in enumerate(all_keys):
                        kt = key_trans.get(key, {})
                        is_first_row = (proj == first_proj and ki == 0)
                        # For this key: which langs are missing (only if this key is in key_missing)
                        this_key_missing_langs = [l for l, keys in key_missing.items() if key in keys]
                        rows_data.append({
                            "page":            page,
                            "en":              en,
                            "module":          proj,
                            "key_name":        key,
                            "trans":           kt,
                            "enu_val":         proj_trans.get("en", ""),
                            "is_first":        is_first_row,
                            "fuzzy":           matched_key != en,
                            "fuzzy_key":       matched_key,
                            "not_found":       False,
                            "key_missing_langs": this_key_missing_langs,  # langs missing for THIS key
                        })
                else:
                    # Single key → one row as before
                    rows_data.append({
                        "page":            page,
                        "en":              en,
                        "module":          proj,
                        "key_name":        all_keys[0] if all_keys else "",
                        "trans":           {k: v for k, v in proj_trans.items() if not k.startswith("_")},
                        "enu_val":         proj_trans.get("en", ""),
                        "is_first":        proj == first_proj,
                        "fuzzy":           matched_key != en,
                        "fuzzy_key":       matched_key,
                        "not_found":       False,
                        "key_missing_langs": [],
                    })
        else:
            rows_data.append({
                "page": page, "en": en, "module": "",
                "key_name": "", "trans": {}, "enu_val": "",
                "is_first": True, "fuzzy": False,
                "fuzzy_key": en, "not_found": True,
            })

    if target_langs:
        allowed   = set(target_langs.values())
        key_order = {v: i for i, v in enumerate(target_langs.values())}
        sorted_langs = sorted(
            [l for l in seen_langs if l in allowed and not _is_regional_variant(l)],
            key=lambda l: key_order.get(l, 999)
        )
        lang_label = {v: k for k, v in target_langs.items()}
    else:
        filtered     = [l for l in seen_langs if not _is_regional_variant(l)]
        sorted_langs = ([l for l in LANG_PRIORITY if l in filtered] +
                        sorted(l for l in filtered if l not in LANG_PRIORITY))
        lang_label   = {}

    log(f"📊 語言: {[lang_label.get(l, l) for l in sorted_langs]}")

    wb = Workbook()

    # ── Sheet 1: 翻譯對照 ─────────────────────────────────────────────────────
    ws = wb.active
    ws.title = "翻譯對照"

    def _hdr(l):
        code = lang_label.get(l, l)
        name = LANG_NAMES.get(l, "")
        return f"{code}\n{name}" if name else code

    MODULE_FILL   = PatternFill("solid", fgColor="FFFFFF")
    MODULE_FONT   = Font(name="Arial", size=9, color="302B63")
    NF_FILL       = PatternFill("solid", fgColor="FFF3CD")
    NF_FONT       = Font(name="Arial", size=9, color="cc4400", bold=True)
    DIM_FONT      = Font(name="Arial", size=10, color="aaaaaa")
    GROUP_TOP     = Border(left=thin, right=thin,
                           top=Side(style="medium", color="302B63"),
                           bottom=Side(style="thin", color="DDDDDD"))
    KEY_COL_FILL  = PatternFill("solid", fgColor="F5F3FF")

    for ci, h in enumerate(["Page", "Module", "EN (Base, from App)", "Key(s)"] +
                            [_hdr(l) for l in sorted_langs], 1):
        c = ws.cell(row=1, column=ci, value=h)
        c.fill, c.font, c.alignment, c.border = HEADER, H_FONT, CENTER, BORDER
    ws.row_dimensions[1].height = 32

    for ri, rd in enumerate(rows_data, 2):
        en_val    = rd["en"]
        trans     = rd["trans"]
        enu_val   = rd.get("enu_val", "")
        key_name  = rd.get("key_name", "")
        is_first  = rd["is_first"]
        not_found = rd["not_found"]
        border    = GROUP_TOP if is_first else BORDER

        valid    = {l: _visual_len(v) for l, v in trans.items()
                    if l in sorted_langs and v and not is_garbled(v)
                    and l != "en" and v.strip() != en_val.strip()}
        max_len  = max(valid.values()) if valid else 0
        max_langs = {l for l, n in valid.items() if n == max_len}

        # If ARA is among the longest, also find the second longest tier
        second_langs: set = set()
        if "ar" in max_langs and len(valid) > len(max_langs):
            second_len   = max((n for l, n in valid.items() if l not in max_langs), default=0)
            second_langs = {l for l, n in valid.items() if n == second_len and l not in max_langs}

        # Col 1: Page — dim on repeated rows within same group
        c = ws.cell(row=ri, column=1, value=rd["page"] if is_first else None)
        c.font = B_FONT if is_first else DIM_FONT
        c.alignment, c.border = LEFT, border

        # Col 2: Module — ALWAYS show (each row = different module or key)
        if not_found:
            c = ws.cell(row=ri, column=2, value="⚠️ 查無字串")
            c.fill, c.font, c.alignment, c.border = NF_FILL, NF_FONT, LEFT, border
        else:
            c = ws.cell(row=ri, column=2, value=rd["module"] or None)
            c.fill = MODULE_FILL
            c.font = MODULE_FONT
            c.alignment, c.border = LEFT, border

        # Col 3: EN Base — dim on repeated rows
        c = ws.cell(row=ri, column=3, value=_xl_safe(en_val) if is_first else None)
        if rd.get("fuzzy") and is_first:
            c.font = Font(name="Arial", size=10, bold=True, color="E65100")
            from openpyxl.comments import Comment
            fuzzy_key = rd["fuzzy_key"]
            note  = f"⚠️ 模糊比對（變數/換行符號已正規化）\n翻譯檔實際字串:\n{fuzzy_key}"
            lines = note.split("\n")
            w = max(280, max(len(l) for l in lines) * 7 + 20)
            h = 20 + len(lines) * 18
            c.comment = Comment(note, "MUI Tool", height=h, width=w)
        else:
            c.font = B_FONT if is_first else DIM_FONT
        c.alignment, c.border = LEFT, border

        # Col 4: Key name (one key per row, no symbols)
        c = ws.cell(row=ri, column=4, value=_xl_safe(key_name) or None)
        c.fill   = KEY_COL_FILL
        c.font = Font(name="Arial", size=10)
        c.alignment = LEFT
        c.border = border

        # Col 5+: translations (per this key's own translations)
        for ci, lang in enumerate(sorted_langs, 5):
            val     = trans.get(lang, "")
            is_base = (lang == "en")
            if is_base:
                val = enu_val
            c = ws.cell(row=ri, column=ci, value=_xl_safe(val) or None)
            c.alignment, c.border = LEFT, border
            if not val or is_garbled(val):
                c.fill, c.font = RED, D_FONT
            elif not is_base and (
                val.strip() == en_val.strip()
                or (enu_val and val.strip() == enu_val.strip())
            ):
                module = rd.get("module", "")
                if ignore_set and _in_ignore(ignore_set, module, key_name, lang):
                    c.font = D_FONT
                else:
                    c.fill, c.font = RED, D_FONT
            elif lang in max_langs:
                c.fill, c.font = YELLOW, B_FONT
            elif lang in second_langs:
                c.fill, c.font = YELLOW2, B_FONT
            else:
                c.font = D_FONT

    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 40   # Module — longest: pf-flutter-photo-editing-copilot-doc (36)
    ws.column_dimensions["C"].width = 42
    ws.column_dimensions["D"].width = 36   # Key(s)
    for i in range(len(sorted_langs)):
        ws.column_dimensions[get_column_letter(i + 5)].width = 28
    ws.freeze_panes = "E2"

    main_col_widths = {"A": 16, "B": 40, "C": 42, "D": 36}
    lang_cols = set()
    for i in range(len(sorted_langs)):
        col = get_column_letter(i + 5)
        main_col_widths[col] = 28
        lang_cols.add(col)
    # Only measure translation columns (E+) and EN column (C) for height calculation
    _auto_row_height(ws, main_col_widths, content_cols=lang_cols | {"C"})

    # ── Sheet 2: 翻譯問題報告 ────────────────────────────────────────────────
    rpt = wb.create_sheet("翻譯問題報告")
    RPT_H_FILL = PatternFill("solid", fgColor="C0392B")

    for ci, h in enumerate(["Module", "ENU 翻譯檔字串", "問題 Key", "語言", "問題類型"], 1):
        c = rpt.cell(row=1, column=ci, value=h)
        c.fill = RPT_H_FILL
        c.font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
        c.alignment, c.border = CENTER, BORDER
    rpt.row_dimensions[1].height = 24

    ISSUE_COLORS = {
        "未翻譯 / 空白":    "FF6B6B",
        "與英文翻譯檔相同": "FFD740",
    }
    rpt_groups: dict[tuple, list] = OrderedDict()

    for rd in rows_data:
        if rd["not_found"]: continue
        trans             = rd["trans"]
        module            = rd["module"]
        enu_val           = rd.get("enu_val", "")
        key_name          = rd.get("key_name", "")
        key_missing_langs = rd.get("key_missing_langs", [])

        for lang in sorted_langs:
            if lang == "en": continue
            val   = trans.get(lang, "")
            label = lang_label.get(lang, lang)

            if not val or is_garbled(val):
                issue    = "未翻譯 / 空白"
                keys_str = key_name
            elif enu_val and val.strip() == enu_val.strip():
                if ignore_set and _in_ignore(ignore_set, module, key_name, lang):
                    continue   # intentional same-as-English → PASS
                issue    = "與英文翻譯檔相同"
                keys_str = key_name
            elif lang in key_missing_langs:
                issue    = "未翻譯 / 空白"
                keys_str = key_name
            else:
                continue

            grp_key = (module, enu_val, issue, keys_str)
            if grp_key not in rpt_groups:
                rpt_groups[grp_key] = []
            if label not in rpt_groups[grp_key]:
                rpt_groups[grp_key].append(label)

    sorted_groups = sorted(rpt_groups.items(), key=lambda x: (x[0][0], x[0][2], x[0][1]))

    prev_module = None
    for ri, ((module, enu_val, issue, keys_str), langs_list) in enumerate(sorted_groups, 2):
        is_new_mod = (module != prev_module)
        prev_module = module
        row_border = Border(
            left=thin, right=thin,
            top=Side(style="medium" if is_new_mod else "thin",
                     color="C0392B" if is_new_mod else "DDDDDD"),
            bottom=thin
        )
        langs_str = ", ".join(langs_list)
        row_h = max(15, len(keys_str.split("\n")) * 14) if keys_str else 15
        rpt.row_dimensions[ri].height = row_h

        for ci, v in enumerate([module, enu_val, keys_str, langs_str, issue], 1):
            c = rpt.cell(row=ri, column=ci, value=v or None)
            c.border = row_border
            if ci == 5:
                color = ISSUE_COLORS.get(issue, "FFFFFF")
                c.fill = PatternFill("solid", fgColor=color)
                c.font = Font(name="Arial", size=10, bold=True)
                c.alignment = LEFT
            elif ci == 3:   # Keys column
                c.font = Font(name="Arial", size=10)
                c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            else:
                c.font = Font(name="Arial", size=10)
                c.alignment = LEFT
                c.font = Font(name="Arial", size=10)

    rpt.column_dimensions["A"].width = 40   # Module
    rpt.column_dimensions["B"].width = 44
    rpt.column_dimensions["C"].width = 36   # Keys
    rpt.column_dimensions["D"].width = 32   # Languages
    rpt.column_dimensions["E"].width = 22   # Issue type
    rpt.freeze_panes = "A2"
    _auto_row_height(rpt, {"A": 40, "B": 44, "C": 36, "D": 32, "E": 22}, content_cols={"B", "C"})
    log(f"📋 翻譯問題報告: {len(sorted_groups)} 筆（{sum(len(v) for v in rpt_groups.values())} 個問題）")

    # ── Sheet 3: 說明 ─────────────────────────────────────────────────────────
    ls = wb.create_sheet("說明")
    for ri, (a, b) in enumerate([
        ("色彩說明", ""),
        ("  黃色底", "該語言在此 module 翻譯中字串最長（唯一最長才標示）"),
        ("  紅色底", "未翻譯、空白、亂碼、或翻譯與英文相同"),
        ("  橘色字", "模糊比對（Excel 數字對應翻譯檔變數，hover 看原始字串）"),
        ("  黃色警示欄", "查無字串（此英文字串在翻譯檔中找不到）"),
        ("", ""),
        ("翻譯問題報告色彩", ""),
        ("  紅底", "未翻譯 / 空白"),
        ("  黃底", "與英文翻譯檔相同（ENU）"),
        ("", ""), ("語言代碼", "語言名稱"),
    ] + list(LANG_NAMES.items()), 1):
        ca = ls.cell(row=ri, column=1, value=a)
        cb = ls.cell(row=ri, column=2, value=b)
        bold = ri in (1, 7, 12)
        ca.font = Font(name="Arial", size=10, bold=bold)
        cb.font = Font(name="Arial", size=10, bold=bold)
        if ri == 2:  ca.fill = YELLOW
        if ri == 3:  ca.fill = RED
        if ri == 8:  ca.fill = PatternFill("solid", fgColor="FF6B6B")
        if ri == 9:  ca.fill = PatternFill("solid", fgColor="FFAB40")
        if ri == 10: ca.fill = PatternFill("solid", fgColor="FFD740")
    ls.column_dimensions["A"].width = 22
    ls.column_dimensions["B"].width = 55

    wb.save(str(output_xlsx))
    log(f"✅ 輸出完成: {output_xlsx.name}")

# ── GUI ───────────────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"MUI Translation Tool  {APP_VERSION}")
        self.resizable(True, True)
        self.minsize(700, 520)
        self.configure(bg="#1e1b35")
        self._cfg = load_config()
        self._zip_path = self._excel_path = self._out_path = None
        self._ignore_path = None
        self._scan_out_path = None
        self._convert_in_path = self._convert_out_path = None
        self._cancel_event = threading.Event()
        self._ignore_var = tk.StringVar(value="未設定（可選）")

        # Scrollable main canvas
        self._canvas = tk.Canvas(self, bg="#1e1b35", highlightthickness=0)
        self._vscroll = ttk.Scrollbar(self, orient="vertical",
                                      command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=self._vscroll.set)
        self._vscroll.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)

        self._inner = tk.Frame(self._canvas, bg="#1e1b35")
        self._inner_id = self._canvas.create_window(
            (0, 0), window=self._inner, anchor="nw")
        self._inner.columnconfigure(0, weight=1)

        self._inner.bind("<Configure>", self._on_frame_configure)
        self._canvas.bind("<Configure>", self._on_canvas_configure)
        # Mouse wheel scrolling
        self.bind_all("<MouseWheel>",
                      lambda e: self._canvas.yview_scroll(-1*(e.delta//120), "units"))
        self.bind_all("<Button-4>",
                      lambda e: self._canvas.yview_scroll(-1, "units"))
        self.bind_all("<Button-5>",
                      lambda e: self._canvas.yview_scroll(1, "units"))

        self._build_ui()
        self._restore_paths()
        self._center()
        threading.Thread(target=self._check_for_update, daemon=True).start()

    def _on_frame_configure(self, event=None):
        canvas_h = self._canvas.winfo_height()
        req_h = self._inner.winfo_reqheight()
        target_h = max(canvas_h if canvas_h > 1 else req_h, req_h)
        self._canvas.itemconfig(self._inner_id, height=target_h)
        self._canvas.configure(scrollregion=(0, 0, self._canvas.winfo_width(), target_h))

    def _on_canvas_configure(self, event):
        self._canvas.itemconfig(self._inner_id, width=event.width)
        req_h = self._inner.winfo_reqheight()
        target_h = max(event.height, req_h)
        self._canvas.itemconfig(self._inner_id, height=target_h)
        self._canvas.configure(scrollregion=(0, 0, event.width, target_h))

    def _restore_paths(self):
        """Restore last app and its paths from config."""
        last_app = self._cfg.get("last_app")
        if last_app and last_app in APP_CONFIGS:
            self._select_app(last_app, init=True)   # restore app selection (already built UI)
            # Re-highlight correct button
            for name, btn in self._app_buttons.items():
                btn.configure(
                    bg="#7c6ef5" if name == last_app else "#252040",
                    fg="#ffffff"  if name == last_app else "#888888"
                )
            self._app_var.set(last_app)
            cfg = APP_CONFIGS.get(last_app, {})
            self._lang_preview.configure(text="  ".join(cfg.keys()) if cfg else "（全部語言）")

        app = self._app_var.get()
        app_cfg = self._cfg.get(app, {})

        def _try_set(path_str, var, attr):
            if path_str and Path(path_str).exists():
                setattr(self, attr, Path(path_str))
                var.set(self._fmt_name(Path(path_str).name))

        _try_set(app_cfg.get("zip"),    self._zip_var,    "_zip_path")
        _try_set(app_cfg.get("excel"),  self._excel_var,  "_excel_path")
        _try_set(app_cfg.get("out"),    self._out_var,    "_out_path")
        _try_set(app_cfg.get("ignore"), self._ignore_var, "_ignore_path")

        scan_out = app_cfg.get("scan_out")
        if scan_out and Path(scan_out).parent.exists():
            self._scan_out_path = Path(scan_out)
            self._scan_out_var.set(self._fmt_name(Path(scan_out).name))

        _try_set(app_cfg.get("convert_in"),  self._convert_in_var,  "_convert_in_path")
        _try_set(app_cfg.get("convert_out"), self._convert_out_var, "_convert_out_path")

        if self._zip_path or self._excel_path:
            self._log(f"📂 已載入上次 {app} 設定")

        # Rebuild scan checkboxes (also restores saved lang selections)
        self._rebuild_scan_checkboxes()

    def _save_app_paths(self):
        """Save current app's paths to config."""
        app = self._app_var.get()
        if app not in self._cfg:
            self._cfg[app] = {}
        if self._zip_path:
            self._cfg[app]["zip"]    = str(self._zip_path)
        if self._excel_path:
            self._cfg[app]["excel"]  = str(self._excel_path)
        if self._out_path:
            self._cfg[app]["out"]    = str(self._out_path)
        if self._ignore_path:
            self._cfg[app]["ignore"] = str(self._ignore_path)
        if self._scan_out_path:
            self._cfg[app]["scan_out"] = str(self._scan_out_path)
        if self._convert_in_path:
            self._cfg[app]["convert_in"]  = str(self._convert_in_path)
        if self._convert_out_path:
            self._cfg[app]["convert_out"] = str(self._convert_out_path)
        if hasattr(self, '_scan_lang_vars'):
            self._cfg[app]["scan_langs"] = [l for l, v in self._scan_lang_vars.items() if v.get()]
        self._cfg["last_app"] = app
        save_config(self._cfg)

    # ── Layout ───────────────────────────────────────────────────────────────

    def _build_ui(self):
        BG = "#1e1b35"
        # ── Title bar ─────────────────────────────────────────────────────────
        hdr = tk.Frame(self._inner, bg="#0f0c29")
        hdr.pack(fill="x")
        title_f = tk.Frame(hdr, bg="#0f0c29")
        title_f.pack(pady=(14, 2))
        tk.Label(title_f, text="MUI Translation Tool",
                 font=("Arial", 16, "bold"), fg="#ffffff", bg="#0f0c29").pack(side="left")
        tk.Label(title_f, text=f" {APP_VERSION}",
                 font=("Arial", 11, "bold"), fg="#7c6ef5", bg="#0f0c29").pack(side="left", pady=(4,0))
        sub_f = tk.Frame(hdr, bg="#0f0c29")
        sub_f.pack(pady=(0, 10))
        tk.Label(sub_f, text=f"by {APP_AUTHOR}",
                 font=("Arial", 10), fg="#444466", bg="#0f0c29").pack(side="left")
        tk.Button(sub_f, text="📋 更新記錄",
                  font=("Arial", 9), fg="#7c6ef5", bg="#0f0c29",
                  activeforeground="#aaaaff", activebackground="#0f0c29",
                  relief="flat", cursor="hand2", bd=0,
                  command=self._show_changelog).pack(side="left", padx=(12, 0))

        # ── Two-column body ───────────────────────────────────────────────────
        body = tk.Frame(self._inner, bg=BG)
        body.pack(fill="x")
        body.columnconfigure(0, weight=0, minsize=400)
        body.columnconfigure(1, weight=1, minsize=340)
        body.rowconfigure(0, weight=0)   # top two columns fixed

        # ── LEFT COLUMN ───────────────────────────────────────────────────────
        left = tk.Frame(body, bg=BG)
        left.grid(row=0, column=0, sticky="new", padx=(16, 8), pady=12)

        # App selector
        app_hdr = tk.Frame(left, bg=BG)
        app_hdr.pack(fill="x", pady=(0, 3))
        tk.Label(app_hdr, text="選擇 App：", font=("Arial", 11, "bold"),
                 fg="#aaaaaa", bg=BG).pack(side="left", padx=(0, 8))
        self._app_var     = tk.StringVar(value=list(APP_CONFIGS.keys())[0])
        self._app_buttons = {}
        for app_name in APP_CONFIGS:
            langs = APP_CONFIGS[app_name]
            count = len(langs) if langs else "全部"
            btn = tk.Button(app_hdr, text=f"{app_name}({count})",
                            font=("Arial", 10, "bold"), relief="flat",
                            cursor="hand2", padx=7, pady=4, bd=0,
                            command=lambda n=app_name: self._select_app(n))
            btn.pack(side="left", padx=(0, 4))
            self._app_buttons[app_name] = btn
        self._lang_preview = tk.Label(left, text="", font=("Arial", 10),
                                      fg="#7c6ef5", bg=BG,
                                      wraplength=370, justify="left")
        self._lang_preview.pack(anchor="w", pady=(0, 8))
        self._select_app(list(APP_CONFIGS.keys())[0], init=True)

        # Mode selector
        mode_hdr = tk.Frame(left, bg=BG)
        mode_hdr.pack(fill="x", pady=(0, 10))
        tk.Label(mode_hdr, text="選擇模式：", font=("Arial", 11, "bold"),
                 fg="#aaaaaa", bg=BG).pack(side="left", padx=(0, 8))
        mode_btns = tk.Frame(mode_hdr, bg="#252040")
        mode_btns.pack(side="left")
        self._mode_var     = tk.StringVar(value="excel")
        self._mode_buttons = {}
        for val, label in [("excel", "Excel 查詢"), ("scan", "語言全掃描"), ("convert", "轉換 Ignore")]:
            btn = tk.Button(mode_btns, text=label,
                            font=("Arial", 11, "bold"), relief="flat",
                            cursor="hand2", padx=14, pady=5, bd=0,
                            command=lambda v=val: self._set_mode(v))
            btn.pack(side="left")
            self._mode_buttons[val] = btn
        self._set_mode("excel", init=True)

        ttk.Separator(left, orient="horizontal").pack(fill="x", pady=(0, 8))

        # File pickers container — swapped between standard and convert mode
        fp_container = tk.Frame(left, bg=BG)
        fp_container.pack(fill="x")

        fp = tk.Frame(fp_container, bg=BG)
        fp.pack(fill="x")
        self._std_fp = fp
        self._zip_var   = tk.StringVar(value="尚未選取")
        self._excel_var = tk.StringVar(value="尚未選取")
        self._out_var   = tk.StringVar(value="尚未選取")
        self._ignore_var = tk.StringVar(value="未設定（可選）")
        self._make_file_row(fp, "📦 翻譯 Zip 檔:",  self._zip_var,   self._pick_zip,   row=0)
        self._excel_row = self._make_file_row(fp, "📄 輸入 Excel:",   self._excel_var, self._pick_excel, row=1)
        self._out_row   = self._make_file_row(fp, "💾 輸出 Excel:",   self._out_var,   self._pick_out,   row=2, is_save=True)
        # Ignore row with clear button
        tk.Label(fp, text="🚫 Ignore Excel:", font=("Arial", 11), fg="#dddddd",
                 bg=BG, anchor="w").grid(row=3, column=0, sticky="w", pady=4)
        tk.Label(fp, textvariable=self._ignore_var, font=("Arial", 9), fg="#888888",
                 bg=BG, anchor="w", width=22).grid(row=3, column=1, sticky="w", padx=(4, 0))
        ig_btns = tk.Frame(fp, bg=BG)
        ig_btns.grid(row=3, column=2, padx=(4, 0))
        tk.Button(ig_btns, text="選取", font=("Arial", 10),
                  bg="#302b63", fg="white", activebackground="#443d7a",
                  activeforeground="white", relief="flat", padx=8, pady=2,
                  cursor="hand2", command=self._pick_ignore).pack(side="left", padx=(0, 3))
        tk.Button(ig_btns, text="清除", font=("Arial", 10),
                  bg="#4a1030", fg="#ffaaaa", activebackground="#6a1040",
                  activeforeground="white", relief="flat", padx=8, pady=2,
                  cursor="hand2", command=self._clear_ignore).pack(side="left")

        # Convert mode file pickers (hidden by default)
        self._conv_fp = tk.Frame(fp_container, bg=BG)
        self._convert_in_var  = tk.StringVar(value="尚未選取")
        self._convert_out_var = tk.StringVar(value="尚未選取")
        self._conv_in_row  = self._make_file_row(
            self._conv_fp, "📊 掃描報告:", self._convert_in_var, self._pick_convert_in, row=0)
        self._conv_out_row = self._make_file_row(
            self._conv_fp, "💾 Ignore 輸出:", self._convert_out_var, self._pick_convert_out,
            row=1, is_save=True)

        ttk.Separator(left, orient="horizontal").pack(fill="x", pady=(10, 8))

        # Run buttons (left column bottom)
        run_row = tk.Frame(left, bg=BG)
        run_row.pack(fill="x")
        self._run_btn = tk.Button(run_row, text="▶  Run",
                                  font=("Arial", 14, "bold"),
                                  bg="#7c6ef5", fg="white",
                                  activebackground="#6a5ce0",
                                  activeforeground="white", relief="flat",
                                  padx=20, pady=8, cursor="hand2",
                                  command=self._run)
        self._run_btn.pack(side="left", padx=(0, 8))
        self._scan_run_btn = tk.Button(run_row, text="▶  開始掃描",
                                       font=("Arial", 14, "bold"),
                                       bg="#e65100", fg="white",
                                       activebackground="#bf360c",
                                       activeforeground="white", relief="flat",
                                       padx=20, pady=8, cursor="hand2",
                                       command=self._run_scan)
        self._convert_run_btn = tk.Button(run_row, text="▶  轉換",
                                          font=("Arial", 14, "bold"),
                                          bg="#2e7d32", fg="white",
                                          activebackground="#1b5e20",
                                          activeforeground="white", relief="flat",
                                          padx=20, pady=8, cursor="hand2",
                                          command=self._run_convert)
        self._cancel_btn = tk.Button(run_row, text="■  取消",
                                     font=("Arial", 14, "bold"),
                                     bg="#c0392b", fg="white",
                                     activebackground="#a93226",
                                     activeforeground="white", relief="flat",
                                     padx=20, pady=8, cursor="hand2",
                                     command=self._cancel_run)
        # Don't pack scan_run_btn, convert_run_btn, or cancel_btn here — _set_mode / _finish_run controls visibility

        # ── RIGHT COLUMN ──────────────────────────────────────────────────────
        right = tk.Frame(body, bg="#16132b")
        right.grid(row=0, column=1, sticky="nsew", padx=(0, 12), pady=12)

        # ── Excel panel (right) ───────────────────────────────────────────────
        self._excel_panel = tk.Frame(right, bg="#16132b")
        self._excel_panel.pack(fill="both", expand=True, padx=12, pady=10)

        tk.Label(self._excel_panel, text="🔍  快速查詢",
                 font=("Arial", 12, "bold"), fg="#aaaaaa", bg="#16132b").pack(anchor="w")
        tk.Label(self._excel_panel, text="每行一個英文字串，Ctrl+Enter 執行",
                 font=("Arial", 10), fg="#444466", bg="#16132b").pack(anchor="w", pady=(0, 4))
        self._ql_text = tk.Text(self._excel_panel, height=5, bg="#252040", fg="#ffffff",
                                font=("Arial", 12), relief="flat",
                                insertbackground="white", padx=8, pady=6)
        self._ql_text.pack(fill="x", pady=(0, 6))
        self._ql_text.bind("<Control-Return>", lambda e: self._quick_lookup())
        self._ql_btn = tk.Button(self._excel_panel, text="🔍  快速查詢",
                                 font=("Arial", 12, "bold"),
                                 bg="#252040", fg="#aaaaaa",
                                 activebackground="#302b63",
                                 activeforeground="white", relief="flat",
                                 padx=14, pady=7, cursor="hand2",
                                 command=self._quick_lookup)
        self._ql_btn.pack(anchor="w")

        tk.Label(self._excel_panel, text="查詢結果", font=("Arial", 10, "bold"),
                 fg="#555577", bg="#16132b").pack(anchor="w", pady=(8, 2))
        ql_result_f = tk.Frame(self._excel_panel, bg="#16132b")
        ql_result_f.pack(fill="both", expand=True)
        self._ql_result_box = tk.Text(ql_result_f, height=8, bg="#0f0c29", fg="#cccccc",
                                      font=("Courier", 11), relief="flat",
                                      state="disabled", wrap="word")
        ql_sb = ttk.Scrollbar(ql_result_f, command=self._ql_result_box.yview)
        self._ql_result_box.configure(yscrollcommand=ql_sb.set)
        self._ql_result_box.pack(side="left", fill="both", expand=True)
        ql_sb.pack(side="right", fill="y")
        self._ql_result_box.tag_config("ok",   foreground="#4caf50")
        self._ql_result_box.tag_config("err",  foreground="#f44336")
        self._ql_result_box.tag_config("info", foreground="#aaaaaa")
        self._ql_result_box.tag_config("ql",   foreground="#7c6ef5")

        # ── Scan panel (right) ────────────────────────────────────────────────
        self._scan_panel = tk.Frame(right, bg="#16132b")
        # shown/hidden by _set_mode

        # ── Convert panel (right) ─────────────────────────────────────────────
        self._convert_panel = tk.Frame(right, bg="#16132b")
        tk.Label(self._convert_panel, text="📋  轉換掃描報告為 Ignore Excel",
                 font=("Arial", 12, "bold"), fg="#aaaaaa", bg="#16132b").pack(anchor="w")
        tk.Label(self._convert_panel,
                 text=(
                     "將語言掃描報告（含 「<語言> Extracted」Sheet）\n"
                     "一鍵轉換成可直接套用的 Ignore Excel。\n\n"
                     "運作規則：\n"
                     "  • 自動偵測所有 「X Extracted」格式的 Sheet\n"
                     "  • 只收錄 Comment = Pass 的列\n"
                     "  • 每個 Sheet 只記錄該 Sheet 自己的語言\n"
                     "    （忽略「語言」欄可能包含的多語言值）\n"
                     "  • 多個 Key 同格（換行分隔）自動拆分\n"
                     "  • 同一 Key 多語言皆 Pass → 合併為同一列"
                 ),
                 font=("Arial", 10), fg="#555577", bg="#16132b",
                 justify="left").pack(anchor="w", pady=(6, 0))

        tk.Label(self._scan_panel, text="選擇要掃描的語言：",
                 font=("Arial", 10, "bold"), fg="#aaaaaa", bg="#16132b"
                 ).pack(anchor="w", pady=(0, 6))

        self._scan_lang_vars: dict[str, tk.BooleanVar] = {}
        cb_frame = tk.Frame(self._scan_panel, bg="#16132b")
        cb_frame.pack(fill="x", pady=(0, 4))

        app  = self._app_var.get()
        cfg  = APP_CONFIGS.get(app, {}) or {}
        langs = [(code, lang) for code, lang in cfg.items() if lang != "en"]
        cols = 4
        for idx, (code, lang) in enumerate(langs):
            var = tk.BooleanVar(value=False)
            self._scan_lang_vars[lang] = var
            tk.Checkbutton(cb_frame, text=code, variable=var,
                           font=("Arial", 9), fg="#dddddd", bg="#16132b",
                           activebackground="#16132b", activeforeground="#7c6ef5",
                           selectcolor="#302b63", relief="flat", cursor="hand2"
                           ).grid(row=idx // cols, column=idx % cols,
                                  padx=6, pady=2, sticky="w")

        sel_row = tk.Frame(self._scan_panel, bg="#16132b")
        sel_row.pack(anchor="w", pady=(4, 8))
        tk.Button(sel_row, text="全選", font=("Arial", 10),
                  bg="#302b63", fg="white", relief="flat", padx=8, pady=2,
                  cursor="hand2",
                  command=lambda: [v.set(True) for v in self._scan_lang_vars.values()]
                  ).pack(side="left", padx=(0, 6))
        tk.Button(sel_row, text="全不選", font=("Arial", 10),
                  bg="#302b63", fg="white", relief="flat", padx=8, pady=2,
                  cursor="hand2",
                  command=lambda: [v.set(False) for v in self._scan_lang_vars.values()]
                  ).pack(side="left")

        scan_out_f = tk.Frame(self._scan_panel, bg="#16132b")
        scan_out_f.pack(fill="x", pady=(0, 6), padx=(0, 12))
        self._scan_out_var = tk.StringVar(value="尚未選取")
        self._make_file_row(scan_out_f, "💾 輸出報告:",
                            self._scan_out_var, self._pick_scan_out,
                            row=0, is_save=True)

        # ── Bottom: progress + log (full width) ───────────────────────────────
        bot = tk.Frame(self._inner, bg=BG)
        bot.pack(fill="both", expand=True, padx=16, pady=(0, 12))

        self._progress = ttk.Progressbar(bot, mode="indeterminate")
        self._progress.pack(fill="x", pady=(0, 2))
        self._progress_label = tk.Label(bot, text="", font=("Arial", 10),
                                        fg="#7c6ef5", bg=BG)
        self._progress_label.pack(anchor="w", pady=(0, 4))

        log_f = tk.Frame(bot, bg=BG)
        log_f.pack(fill="both", expand=True)
        self._log_box = tk.Text(log_f, height=1, bg="#0f0c29", fg="#cccccc",
                                font=("Courier", 11), relief="flat",
                                state="disabled", wrap="word",
                                insertbackground="white")
        sb = ttk.Scrollbar(log_f, command=self._log_box.yview)
        self._log_box.configure(yscrollcommand=sb.set)
        self._log_box.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self._log_box.tag_config("ok",   foreground="#4caf50")
        self._log_box.tag_config("err",  foreground="#f44336")
        self._log_box.tag_config("info", foreground="#aaaaaa")
        self._log_box.tag_config("ql",   foreground="#7c6ef5")

    @staticmethod
    def _fmt_name(name: str, n: int = 32) -> str:
        return ("…" + name[-(n - 1):]) if len(name) > n else name

    def _make_file_row(self, parent, label, var, cmd, row, is_save=False):
        parent.columnconfigure(1, weight=1)   # filename column stretches; button always visible
        lbl = tk.Label(parent, text=label, font=("Arial", 11), fg="#dddddd",
                       bg=parent["bg"], width=14, anchor="w")
        lbl.grid(row=row, column=0, sticky="w", pady=4)
        val = tk.Label(parent, textvariable=var, font=("Arial", 9), fg="#888888",
                       bg=parent["bg"], anchor="w")
        val.grid(row=row, column=1, sticky="ew", padx=(4, 0))
        btn = tk.Button(parent, text="選取", font=("Arial", 10),
                        bg="#302b63", fg="white", activebackground="#443d7a",
                        activeforeground="white", relief="flat", padx=8, pady=2,
                        cursor="hand2", command=cmd)
        btn.grid(row=row, column=2, padx=(4, 0))
        return lbl, val, btn

    def _set_file_row_state(self, widgets, enabled: bool):
        lbl, val, btn = widgets
        if enabled:
            lbl.configure(fg="#dddddd")
            val.configure(fg="#888888")
            btn.configure(state="normal", bg="#302b63", fg="white", cursor="hand2")
        else:
            lbl.configure(fg="#444455")
            val.configure(fg="#333344")
            btn.configure(state="disabled", bg="#222035", fg="#444455", cursor="")

    # ── App selector ─────────────────────────────────────────────────────────

    def _select_app(self, app_name: str, init: bool = False):
        if not init:
            # Save current app's paths before switching
            self._save_app_paths()

        self._app_var.set(app_name)
        for name, btn in self._app_buttons.items():
            if name == app_name:
                btn.configure(bg="#7c6ef5", fg="#ffffff",
                              activebackground="#6a5ce0", activeforeground="#ffffff")
            else:
                btn.configure(bg="#252040", fg="#888888",
                              activebackground="#302b63", activeforeground="#ffffff")
        cfg = APP_CONFIGS.get(app_name)
        self._lang_preview.configure(
            text="  ".join(cfg.keys()) if cfg else "（全部語言）"
        )

        if not init:
            # Load this app's saved paths (don't clear — restore from config)
            self._zip_path = self._excel_path = self._out_path = self._ignore_path = None
            self._scan_out_path = None
            self._convert_in_path = self._convert_out_path = None
            self._zip_var.set("尚未選取")
            self._excel_var.set("尚未選取")
            self._out_var.set("尚未選取（將放在 Excel 同目錄）")
            self._ignore_var.set("未設定（可選）")
            self._scan_out_var.set("尚未選取")
            self._convert_in_var.set("尚未選取")
            self._convert_out_var.set("尚未選取")

            app_cfg = self._cfg.get(app_name, {})
            def _try_set(path_str, var, attr):
                if path_str and Path(path_str).exists():
                    setattr(self, attr, Path(path_str))
                    var.set(self._fmt_name(Path(path_str).name))
            _try_set(app_cfg.get("zip"),    self._zip_var,    "_zip_path")
            _try_set(app_cfg.get("excel"),  self._excel_var,  "_excel_path")
            _try_set(app_cfg.get("out"),    self._out_var,    "_out_path")
            _try_set(app_cfg.get("ignore"), self._ignore_var, "_ignore_path")

            scan_out = app_cfg.get("scan_out")
            if scan_out and Path(scan_out).parent.exists():
                self._scan_out_path = Path(scan_out)
                self._scan_out_var.set(self._fmt_name(Path(scan_out).name))

            _try_set(app_cfg.get("convert_in"),  self._convert_in_var,  "_convert_in_path")
            _try_set(app_cfg.get("convert_out"), self._convert_out_var, "_convert_out_path")

            self._cfg["last_app"] = app_name
            save_config(self._cfg)
            self._rebuild_scan_checkboxes()

    def _rebuild_scan_checkboxes(self):
        """Rebuild lang checkboxes when App selection changes."""
        if not hasattr(self, '_scan_lang_vars'): return
        for widget in self._scan_panel.winfo_children():
            if isinstance(widget, tk.Frame):
                for child in widget.winfo_children():
                    child.destroy()
                break
        self._scan_lang_vars.clear()
        app   = self._app_var.get()
        cfg   = APP_CONFIGS.get(app, {}) or {}
        langs = [(code, lang) for code, lang in cfg.items() if lang != "en"]
        cols  = 4
        frames = [w for w in self._scan_panel.winfo_children() if isinstance(w, tk.Frame)]
        cb_frame = frames[0] if frames else self._scan_panel
        saved_langs = set(self._cfg.get(self._app_var.get(), {}).get("scan_langs", []))
        for idx, (code, lang) in enumerate(langs):
            var = tk.BooleanVar(value=lang in saved_langs)
            self._scan_lang_vars[lang] = var
            tk.Checkbutton(cb_frame, text=code, variable=var,
                           font=("Arial", 9), fg="#dddddd", bg="#16132b",
                           activebackground="#16132b", activeforeground="#7c6ef5",
                           selectcolor="#302b63", relief="flat", cursor="hand2"
                           ).grid(row=idx // cols, column=idx % cols,
                                  padx=6, pady=2, sticky="w")

    # ── Mode toggle ───────────────────────────────────────────────────────────

    def _set_mode(self, mode: str, init: bool = False):
        self._mode_var.set(mode)
        for val, btn in self._mode_buttons.items():
            if val == mode:
                btn.configure(bg="#7c6ef5", fg="#ffffff",
                              activebackground="#6a5ce0", activeforeground="#ffffff")
            else:
                btn.configure(bg="#252040", fg="#666666",
                              activebackground="#302b63", activeforeground="#ffffff")

        # Switch run button (guard for init before buttons are created)
        if hasattr(self, '_run_btn') and hasattr(self, '_scan_run_btn'):
            self._run_btn.pack_forget()
            self._scan_run_btn.pack_forget()
            if hasattr(self, '_convert_run_btn'):
                self._convert_run_btn.pack_forget()
            if mode == "excel":
                self._run_btn.pack(side="left", padx=(0, 8))
            elif mode == "scan":
                self._scan_run_btn.pack(side="left")
            else:
                self._convert_run_btn.pack(side="left")

        # Show/hide file picker panels based on mode
        if hasattr(self, '_std_fp') and hasattr(self, '_conv_fp'):
            if mode == "convert":
                self._std_fp.pack_forget()
                self._conv_fp.pack(fill="x")
            else:
                self._conv_fp.pack_forget()
                self._std_fp.pack(fill="x")

        # Enable/disable 輸入 & 輸出 Excel rows based on mode
        if hasattr(self, '_excel_row') and hasattr(self, '_out_row'):
            excel_mode = (mode == "excel")
            self._set_file_row_state(self._excel_row, excel_mode)
            self._set_file_row_state(self._out_row,   excel_mode)

        if not init:
            if mode == "excel":
                self._scan_panel.pack_forget()
                if hasattr(self, '_convert_panel'):
                    self._convert_panel.pack_forget()
                self._excel_panel.pack(fill="both", expand=True, padx=12, pady=10)
            elif mode == "scan":
                self._excel_panel.pack_forget()
                if hasattr(self, '_convert_panel'):
                    self._convert_panel.pack_forget()
                self._scan_panel.pack(fill="both", expand=True, padx=12, pady=10)
            else:
                self._excel_panel.pack_forget()
                self._scan_panel.pack_forget()
                self._convert_panel.pack(fill="both", expand=True, padx=12, pady=10)

    def _clear_ignore(self):
        self._ignore_path = None
        self._ignore_var.set("未設定（可選）")
        self._save_app_paths()

    # ── Index loader (shared by _run, _run_scan, _quick_lookup) ──────────────

    def _load_index(self, zip_path: Path) -> tuple:
        import tempfile, hashlib
        zip_stat  = zip_path.stat()
        cache_key = hashlib.md5(
            f"{zip_path}|{zip_stat.st_mtime}|{zip_stat.st_size}|{TOOL_VERSION}".encode()
        ).hexdigest()[:12]
        cache_dir  = Path(tempfile.gettempdir()) / "TranslationTool"
        cache_dir.mkdir(exist_ok=True)
        index_path = cache_dir / f"lookup_{zip_path.stem}_{cache_key}.json"

        if index_path.exists():
            self._log("📂 發現快取索引，直接載入...")
            raw        = json.loads(index_path.read_text("utf-8"))
            index      = {r["en"]: r["projects"] for r in raw}
            norm_index = {_normalize(k): k for k in index}
            self._log(f"✅ 索引載入: {len(index)} 個字串")
            return index, norm_index

        index, norm_index = build_index(zip_path, "en", self._log)
        try:
            out_list = [{"en": k, "projects": v} for k, v in index.items()]
            index_path.write_text(
                json.dumps(out_list, ensure_ascii=False, separators=(",", ":")),
                encoding="utf-8")
            self._log("💾 索引已快取 (下次更快)")
        except Exception:
            self._log("⚠️  快取寫入失敗，不影響本次執行")
        return index, norm_index

    # ── Scan output picker ────────────────────────────────────────────────────

    def _pick_scan_out(self):
        p = filedialog.asksaveasfilename(title="選取輸出報告路徑",
                                         defaultextension=".xlsx",
                                         filetypes=[("Excel files", "*.xlsx")])
        if p:
            self._scan_out_path = Path(p)
            self._scan_out_var.set(self._fmt_name(Path(p).name))
            self._save_app_paths()

    # ── Convert Ignore pickers & run ─────────────────────────────────────────

    def _pick_convert_in(self):
        p = filedialog.askopenfilename(title="選取語言掃描報告",
                                       filetypes=[("Excel files", "*.xlsx *.xls"),
                                                  ("All files", "*.*")])
        if p:
            self._convert_in_path = Path(p)
            self._convert_in_var.set(self._fmt_name(Path(p).name))
            if not self._convert_out_path:
                default = Path(p).parent / (Path(p).stem + "_ignore.xlsx")
                self._convert_out_path = default
                self._convert_out_var.set(self._fmt_name(default.name))
            self._save_app_paths()

    def _pick_convert_out(self):
        p = filedialog.asksaveasfilename(title="選取 Ignore Excel 輸出路徑",
                                          defaultextension=".xlsx",
                                          filetypes=[("Excel files", "*.xlsx")])
        if p:
            self._convert_out_path = Path(p)
            self._convert_out_var.set(self._fmt_name(Path(p).name))
            self._save_app_paths()

    def _run_convert(self):
        if not self._convert_in_path:
            messagebox.showwarning("提示", "請先選取語言掃描報告"); return

        out_path = self._convert_out_path
        if not out_path:
            out_path = self._convert_in_path.parent / (self._convert_in_path.stem + "_ignore.xlsx")
            self._convert_out_path = out_path
            self._convert_out_var.set(self._fmt_name(out_path.name))

        self._cancel_event.clear()
        self._convert_run_btn.configure(state="disabled")
        self._convert_run_btn.pack_forget()
        self._cancel_btn.pack(side="left")
        self._progress.configure(mode="indeterminate")
        self._progress.start(12)
        self._log_box.configure(state="normal")
        self._log_box.delete("1.0", "end")
        self._log_box.configure(state="disabled")

        in_path = self._convert_in_path

        def worker():
            try:
                n = convert_scan_to_ignore(in_path, out_path, self._log)
                self._log(f"\n🎉 完成！共 {n} 筆，已儲存至:\n   {out_path}")
                if messagebox.askyesno("完成", f"Ignore Excel 已產生！\n\n{out_path}\n\n是否立即開啟？"):
                    import os
                    if sys.platform == "win32":
                        os.startfile(str(out_path))
                    else:
                        import subprocess
                        subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", str(out_path)])
            except Exception as ex:
                import traceback
                self._log(f"❌ 發生錯誤: {ex}", "err")
                self._log(traceback.format_exc(), "err")
                messagebox.showerror("錯誤", str(ex))
            finally:
                self.after(0, self._finish_run, self._convert_run_btn)

        threading.Thread(target=worker, daemon=True).start()

    # ── Scan run ─────────────────────────────────────────────────────────────

    def _run_scan(self):
        if not self._zip_path:
            messagebox.showwarning("提示", "請先選取翻譯 Zip 檔"); return

        selected = [lang for lang, var in self._scan_lang_vars.items() if var.get()]
        if not selected:
            messagebox.showwarning("提示", "請至少選擇一個語言"); return

        out_path = getattr(self, "_scan_out_path", None)
        if not out_path:
            out_path = self._zip_path.parent / "scan_report.xlsx"
            self._scan_out_path = out_path
            self._scan_out_var.set(self._fmt_name(out_path.name))

        self._cancel_event.clear()
        self._scan_run_btn.configure(state="disabled")
        self._scan_run_btn.pack_forget()
        self._cancel_btn.pack(side="left")
        self._progress.configure(mode="indeterminate")
        self._progress.start(12)
        self._progress_label.configure(text="")
        self._log_box.configure(state="normal")
        self._log_box.delete("1.0", "end")
        self._log_box.configure(state="disabled")

        zip_path = self._zip_path

        def worker():
            try:
                index, _ = self._load_index(zip_path)
                if self._cancel_event.is_set():
                    self._log("⚠️  已取消"); return
                target_langs = APP_CONFIGS.get(self._app_var.get())

                ignore_set = set()
                if self._ignore_path and Path(self._ignore_path).exists():
                    ignore_set = load_ignore_list(Path(self._ignore_path), target_langs)
                    self._log(f"🚫 Ignore list: {len(ignore_set)} 筆")

                generate_scan_report(index, out_path, self._log, selected,
                                     target_langs, ignore_set, self._cancel_event)
                if self._cancel_event.is_set(): return
                self._log(f"\n🎉 完成！報告已儲存至:\n   {out_path}")

                if messagebox.askyesno("完成", f"掃描報告已產生！\n\n{out_path}\n\n是否立即開啟？"):
                    import subprocess, os
                    if sys.platform == "win32":   os.startfile(str(out_path))
                    elif sys.platform == "darwin": subprocess.Popen(["open", str(out_path)])
                    else:                          subprocess.Popen(["xdg-open", str(out_path)])

            except Exception as ex:
                import traceback
                self._log(f"❌ 發生錯誤: {ex}", "err")
                self._log(traceback.format_exc(), "err")
                messagebox.showerror("錯誤", str(ex))
            finally:
                self.after(0, self._finish_run, self._scan_run_btn)

        threading.Thread(target=worker, daemon=True).start()

    # ── File pickers ─────────────────────────────────────────────────────────

    def _pick_zip(self):
        p = filedialog.askopenfilename(title="選取翻譯 Zip 檔",
                                       filetypes=[("Zip files", "*.zip"), ("All files", "*.*")])
        if p:
            self._zip_path = Path(p)
            self._zip_var.set(self._fmt_name(Path(p).name))
            self._save_app_paths()

    def _pick_excel(self):
        p = filedialog.askopenfilename(title="選取輸入 Excel",
                                       filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
        if p:
            self._excel_path = Path(p)
            self._excel_var.set(self._fmt_name(Path(p).name))
            if not self._out_path:
                default = self._excel_path.parent / (self._excel_path.stem + "_translated.xlsx")
                self._out_path = default
                self._out_var.set(self._fmt_name(default.name))
            self._save_app_paths()

    def _pick_out(self):
        p = filedialog.asksaveasfilename(title="選取輸出 Excel 路徑",
                                         defaultextension=".xlsx",
                                         filetypes=[("Excel files", "*.xlsx")])
        if p:
            self._out_path = Path(p)
            self._out_var.set(self._fmt_name(Path(p).name))
            self._save_app_paths()

    def _pick_ignore(self):
        p = filedialog.askopenfilename(title="選取 Ignore Excel",
                                       filetypes=[("Excel files", "*.xlsx *.xls"),
                                                  ("All files", "*.*")])
        if p:
            self._ignore_path = Path(p)
            self._ignore_var.set(self._fmt_name(Path(p).name))
            self._save_app_paths()

    # ── Log ──────────────────────────────────────────────────────────────────

    def _cancel_run(self):
        self._cancel_event.set()
        self._cancel_btn.configure(state="disabled", text="■  取消中…")

    def _finish_run(self, run_btn):
        self._progress.stop()
        self._progress.configure(mode="indeterminate", value=0)
        self._progress_label.configure(text="")
        self._cancel_btn.pack_forget()
        self._cancel_btn.configure(state="normal", text="■  取消")
        run_btn.configure(state="normal")
        if run_btn is self._run_btn:
            run_btn.pack(side="left", padx=(0, 8))
        else:
            run_btn.pack(side="left")

    def _log(self, msg: str, tag: str = None):
        self.after(0, self._log_main, msg, tag)

    def _log_ql(self, msg: str, tag: str = None):
        self.after(0, self._log_ql_main, msg, tag)

    def _log_ql_main(self, msg: str, tag: str = None):
        if tag is None:
            tag = "ok" if msg.startswith("✅") else "err" if msg.startswith("❌") else "ql"
        self._ql_result_box.configure(state="normal")
        self._ql_result_box.insert("end", msg + "\n", tag)
        self._ql_result_box.see("end")
        self._ql_result_box.configure(state="disabled")

    def _log_main(self, msg: str, tag: str = None):
        if tag is None:
            tag = "ok" if msg.startswith("✅") else "err" if msg.startswith("❌") else "info"
        self._log_box.configure(state="normal")
        self._log_box.insert("end", msg + "\n", tag)
        self._log_box.see("end")
        self._log_box.configure(state="disabled")
        if "處理第" in msg or "掃描第" in msg:
            self._progress_label.configure(text=msg.strip())
            m = re.search(r'(\d+)\s*/\s*(\d+)', msg)
            if m:
                cur, tot = int(m.group(1)), int(m.group(2))
                if self._progress.cget("mode") == "indeterminate":
                    self._progress.stop()
                    self._progress.configure(mode="determinate")
                self._progress.configure(maximum=tot, value=cur)

    # ── Quick Lookup (#8) ────────────────────────────────────────────────────

    def _quick_lookup(self):
        queries = [l.strip() for l in self._ql_text.get("1.0", "end").splitlines() if l.strip()]
        if not queries:
            messagebox.showwarning("提示", "請輸入要查詢的英文字串"); return
        if not self._zip_path:
            messagebox.showwarning("提示", "請先選取翻譯 Zip 檔"); return

        self._ql_btn.configure(state="disabled")
        self._progress.start(12)
        self.after(0, lambda: (
            self._ql_result_box.configure(state="normal"),
            self._ql_result_box.delete("1.0", "end"),
            self._ql_result_box.configure(state="disabled")
        ))

        def worker():
            try:
                index, norm_index = self._load_index(self._zip_path)
                target_langs = APP_CONFIGS.get(self._app_var.get())
                lang_label   = {v: k for k, v in target_langs.items()} if target_langs else {}
                allowed      = set(target_langs.values()) if target_langs else None

                self._log_ql("─────────────────────────────────────")
                for q in queries:
                    rec = index.get(q)
                    matched = q
                    if not rec:
                        for k, v in index.items():
                            if k.lower() == q.lower(): rec = v; matched = k; break
                    if not rec:
                        orig = norm_index.get(_normalize(q))
                        if orig: rec = index.get(orig); matched = orig

                    if not rec:
                        self._log_ql(f"❌  查無字串: {q!r}", "err")
                        continue

                    note = f" → 模糊比對: {matched!r}" if matched != q else ""
                    self._log_ql(f"✅  {q!r}{note}")
                    for proj, pt in sorted(rec.items()):
                        for lang, val in sorted(pt.items()):
                            if allowed and lang not in allowed: continue
                            code   = lang_label.get(lang, lang)
                            status = "🔴" if (not val or val.strip() == q.strip()) else "✅"
                            self._log_ql(f"   {proj}  [{code}]  {status}  {val[:60]}", "info")
                self._log_ql("─────────────────────────────────────")

            except Exception as ex:
                self._log_ql(f"❌ 查詢錯誤: {ex}", "err")
            finally:
                self.after(0, lambda: (self._progress.stop(),
                                       self._ql_btn.configure(state="normal")))

        threading.Thread(target=worker, daemon=True).start()

    # ── Run ──────────────────────────────────────────────────────────────────

    def _run(self):
        if not self._zip_path:
            messagebox.showwarning("提示", "請先選取翻譯 Zip 檔"); return
        if not self._excel_path:
            messagebox.showwarning("提示", "請先選取輸入 Excel"); return
        if not self._out_path:
            self._out_path = self._excel_path.parent / (self._excel_path.stem + "_translated.xlsx")
            self._out_var.set(self._fmt_name(self._out_path.name))

        self._cancel_event.clear()
        self._run_btn.configure(state="disabled")
        self._run_btn.pack_forget()
        self._cancel_btn.pack(side="left", padx=(0, 8))
        self._progress.configure(mode="indeterminate")
        self._progress.start(12)
        self._progress_label.configure(text="")
        self._log_box.configure(state="normal")
        self._log_box.delete("1.0", "end")
        self._log_box.configure(state="disabled")

        zip_path = self._zip_path
        excel_path = self._excel_path
        out_path = self._out_path

        def worker():
            try:
                index, norm_index = self._load_index(zip_path)
                if self._cancel_event.is_set():
                    self._log("⚠️  已取消"); return
                selected_app = self._app_var.get()
                target_langs = APP_CONFIGS.get(selected_app)
                self._log(f"🎯 App: {selected_app}  ({len(target_langs) if target_langs else '全部'} 語言)")

                ignore_set = set()
                if self._ignore_path and Path(self._ignore_path).exists():
                    ignore_set = load_ignore_list(Path(self._ignore_path), target_langs)
                    self._log(f"🚫 Ignore list: {len(ignore_set)} 筆")

                generate_excel(index, norm_index, excel_path, out_path, self._log,
                               target_langs, ignore_set, self._cancel_event)
                if self._cancel_event.is_set(): return
                self._log(f"\n🎉 完成！檔案已儲存至:\n   {out_path}")

                if messagebox.askyesno("完成", f"Excel 已產生！\n\n{out_path}\n\n是否立即開啟？"):
                    import subprocess, os
                    if sys.platform == "win32":
                        os.startfile(str(out_path))
                    elif sys.platform == "darwin":
                        subprocess.Popen(["open", str(out_path)])
                    else:
                        subprocess.Popen(["xdg-open", str(out_path)])

            except Exception as ex:
                import traceback
                self._log(f"❌ 發生錯誤: {ex}", "err")
                self._log(traceback.format_exc(), "err")
                messagebox.showerror("錯誤", str(ex))
            finally:
                self.after(0, self._finish_run, self._run_btn)

        threading.Thread(target=worker, daemon=True).start()

    # ── Utils ─────────────────────────────────────────────────────────────────

    def _center(self):
        self.update_idletasks()
        w, h = 880, 750
        self.geometry(f"{w}x{h}")
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    # ── Changelog dialog ──────────────────────────────────────────────────────

    def _show_changelog(self):
        win = tk.Toplevel(self)
        win.title("更新記錄")
        win.configure(bg="#1e1b35")
        win.resizable(False, False)
        win.grab_set()

        tk.Label(win, text="更新記錄", font=("Arial", 13, "bold"),
                 fg="#ffffff", bg="#1e1b35").pack(pady=(16, 6))

        frame = tk.Frame(win, bg="#1e1b35")
        frame.pack(fill="both", expand=True, padx=16, pady=(0, 16))

        txt = tk.Text(frame, width=58, height=28,
                      bg="#0f0c29", fg="#cccccc",
                      font=("Courier", 10), relief="flat",
                      wrap="word", state="normal",
                      padx=12, pady=10)
        sb = ttk.Scrollbar(frame, command=txt.yview)
        txt.configure(yscrollcommand=sb.set)
        txt.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        # Insert with highlight for version headers
        for line in CHANGELOG.splitlines(keepends=True):
            if line.startswith("v1."):
                txt.insert("end", line, "ver")
            elif line.startswith("─"):
                txt.insert("end", line, "sep")
            elif line.strip().startswith("•"):
                txt.insert("end", line, "item")
            elif line.strip() in ("新功能", "Bug 修正", "效能 / 架構"):
                txt.insert("end", line, "section")
            else:
                txt.insert("end", line)

        txt.tag_config("ver",     foreground="#7c6ef5", font=("Courier", 11, "bold"))
        txt.tag_config("sep",     foreground="#333355")
        txt.tag_config("section", foreground="#aaaaaa", font=("Courier", 10, "bold"))
        txt.tag_config("item",    foreground="#cccccc")
        txt.configure(state="disabled")

        tk.Button(win, text="關閉", font=("Arial", 10),
                  bg="#302b63", fg="white", relief="flat",
                  padx=20, pady=4, cursor="hand2",
                  command=win.destroy).pack(pady=(0, 14))

        # Centre over main window
        self.update_idletasks()
        x = self.winfo_x() + (self.winfo_width()  - win.winfo_reqwidth())  // 2
        y = self.winfo_y() + (self.winfo_height() - win.winfo_reqheight()) // 2
        win.geometry(f"+{x}+{y}")

    # ── Auto-update check ─────────────────────────────────────────────────────

    @staticmethod
    def _parse_ver(s: str):
        """Parse 'v1.3.260508_1030' → (1, 3, '260508_1030'). Returns None if unparseable."""
        m = re.search(r'(\d+)\.(\d+)\.(\w+)', s)
        if m:
            return (int(m.group(1)), int(m.group(2)), m.group(3))
        return None

    def _check_for_update(self):
        """Background thread: silently fetch latest_version.txt and compare."""
        try:
            import urllib.request
            with urllib.request.urlopen(_UPDATE_CHECK_URL, timeout=6) as resp:
                latest_raw = resp.read().decode("utf-8").strip()
            cur = self._parse_ver(APP_VERSION)
            new = self._parse_ver(latest_raw)
            if cur and new and new != cur:
                # Trigger if major.minor is newer, OR same major.minor but newer build stamp
                if (new[0], new[1]) > (cur[0], cur[1]) or (
                    (new[0], new[1]) == (cur[0], cur[1]) and new[2] > cur[2]
                ):
                    self.after(0, self._show_update_dialog, latest_raw.strip())
        except Exception:
            pass  # network error or timeout — silently skip

    def _show_update_dialog(self, latest_ver: str):
        import webbrowser
        cur = self._parse_ver(APP_VERSION)
        cur_str = f"v{cur[0]}.{cur[1]}.{cur[2]}" if cur else APP_VERSION
        answer = messagebox.askyesno(
            "🆕 發現新版本",
            f"目前版本：{cur_str}\n最新版本：v{latest_ver}\n\n是否前往 Google Drive 下載新版？",
            icon="info"
        )
        if answer:
            webbrowser.open(_UPDATE_DOWNLOAD_URL)


if __name__ == "__main__":
    App().mainloop()
