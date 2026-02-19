# Word LaTeX to OMath Converter (VBA)

這是一個專為 Microsoft Word 設計的 VBA 巨集工具，旨在將選取範圍內的 LaTeX 數學公式快速轉換為 Word 內建的方程式物件（OMath）。

## ✨ 功能亮點
- **多格式支援**：支援 `$ ... $`, `$$...$$`, `\( ... \)` 及 `\[ ... \]` 等常見定界符。
- **智慧清理**：自動移除 `\tag{...}` 並轉換 `\text{...}`, `\mathrm{...}` 為純文字，避免轉換錯誤。
- **巢狀處理**：能正確解析含有巢狀大括號 `{}` 的 LaTeX 指令。
- **符號優化**：自動將 `\mu`, `\times`, `\cdot` 轉換為對應符號，提升 Word 解析成功率。

## 🚀 安裝與使用教學

### 1. 匯入巨集
1. 下載本專案的 `ConvertLaTeXToOMath.bas` 檔案。
2. 開啟 Word，按下 `Alt + F11` 開啟 VBA 編輯器。
3. 在左側選單點擊右鍵 -> `Import File...` (匯入檔案)，選擇下載的 `.bas` 檔。

### 2. 執行轉換
1. 在 Word 文件中，**反白選取**包含 LaTeX 公式的文字範圍。
2. 按下 `Alt + F8`，選擇 `ConvertLaTeXToOMath_V1` 並點擊「執行」。
3. 稍等片刻，公式即會自動轉換為漂亮的 Word 方程式！

## ⚠️ 注意事項
- 執行前請務必先選取範圍，否則會跳出提示。
- 複雜的環境（如 `align` 或 `matrix`）建議先簡化後再轉換。

---
*Developed for efficient academic writing and documentation.*