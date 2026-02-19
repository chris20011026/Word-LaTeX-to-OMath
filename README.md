# Word LaTeX to OMath Converter (VBA) | Word LaTeX 方程式轉換器

這是一個專為 Microsoft Word 設計的 VBA 巨集工具，旨在將選取範圍內的 LaTeX 數學公式快速轉換為 Word 內建的方程式物件 (OMath)。

---

## 🇹🇼 繁體中文說明 (Traditional Chinese)

### ✨ 功能亮點
- **多格式支援**：支援 `$ ... $`, `$$...$$`, `\( ... \)` 及 `\[ ... \]` 等常見定界符。
- **智慧清理**：自動移除 `\tag{...}` 並轉換 `\text{...}`, `\mathrm{...}` 等指令為純文字。
- **巢狀處理**：能正確解析含有巢狀大括號 `{}` 的 LaTeX 指令。

### 🚀 安裝與使用教學
1. **匯入巨集**：下載 `ConvertLaTeXToOMath.bas`。在 Word 中按 `Alt + F11` 開啟編輯器，右鍵點擊左側選單選擇 `Import File...` 匯入。
2. **設定快捷鍵 (強烈建議)**：
   - 前往 `檔案` > `選項` > `自訂功能區`。
   - 點擊下方 `鍵盤快速鍵：自訂` 按鈕。
   - 在左側「類別」捲動到最下方選擇 `巨集`。
   - 在右側找到 `ConvertLaTeXToOMath_V1`。
   - 在「請按新設定的快速鍵」處按下 `Alt + Q` (或任何你喜歡的按鍵)，點擊 `指派`。
3. **執行轉換**：在 Word 中**反白選取** LaTeX 公式範圍，按下剛設定好的快捷鍵 (如 `Alt + Q`)，公式即刻完成轉換！

---

## 🇺🇸 English Description

### ✨ Key Features
- **Multi-format Support**: Supports `$ ... $`, `$$...$$`, `\( ... \)` and `\[ ... \]`.
- **Smart Cleaning**: Automatically removes `\tag{...}` and strips commands like `\text{...}`.
- **Nested Braces**: Correctly handles LaTeX commands with nested braces `{}`.

### 🚀 Installation & Usage
1. **Import Macro**: Download `ConvertLaTeXToOMath.bas`. Press `Alt + F11` in Word, right-click in the project pane, and select `Import File...`.
2. **Set Shortcut Key (Recommended)**:
   - Go to `File` > `Options` > `Customize Ribbon`.
   - Click the `Keyboard shortcuts: Customize` button at the bottom.
   - Scroll down to `Macros` in the "Categories" list.
   - Select `ConvertLaTeXToOMath_V1` from the "Macros" list.
   - Press `Alt + Q` (or your preferred key) in the "Press new shortcut key" box, then click `Assign`.
3. **Run Conversion**: **Highlight/Select** the LaTeX formulas in Word, press your shortcut key (e.g., `Alt + Q`), and the conversion is done!

---
*Developed for efficient academic writing and documentation.*
