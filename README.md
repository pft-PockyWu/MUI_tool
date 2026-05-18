# Translation Excel Generator

## 使用方式（直接跑 Python，不需打包）
```
pip install -r requirements.txt
python translation_gui.py
```

## 打包成執行檔（只需做一次）

### Mac
```bash
bash build_mac.sh
# → dist/TranslationTool  （雙擊執行）
```

### Windows
```
雙擊 build_windows.bat
# → dist\TranslationTool.exe  （雙擊執行）
```

## 新增語言
打開 translation_gui.py，找到 TARGET_LANGS，加一行：
```python
TARGET_LANGS = {
    ...
    "HEB": "he",   # 希伯來語
    "POL": "pl",   # 波蘭語
    "FIL": "fil",  # 菲律賓語
}
```
修改後重新打包一次即可。

## 色彩說明
- 🟡 黃色底 = 該語言是此列最長翻譯（唯一最長才標示）
- 🔴 紅色底 = 未翻譯或含亂碼
