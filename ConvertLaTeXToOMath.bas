Attribute VB_Name = "Module1"
Option Explicit

' ConvertLaTeXToOMath_V1（修正版）
' 處理順序：先 $$...$$（display math），再 $...$，再 \(..\) 與 \[..\]
' 另外移除 \tag{...} 以免造成 OMath 解析錯誤

Sub ConvertLaTeXToOMath_V1()
    On Error GoTo EH

    Dim sel As Selection
    Set sel = Application.Selection

    If sel.Type = wdSelectionIP Then
        MsgBox "請先選取欲轉換的文字範圍。", vbExclamation, "需要選取範圍"
        Exit Sub
    End If

    Dim selRange As Range
    Set selRange = sel.Range

    ' 先處理 $$...$$（display math），再處理單一 $...$
    Call ProcessPattern(selRange, "\$\$([\s\S]+?)\$\$")
    Call ProcessPattern(selRange, "\$([\s\S]+?)\$")

    ' 處理 \( ... \) 與 \[ ... \]
    Call ProcessPattern(selRange, "\\\(([\s\S]+?)\\\)")
    Call ProcessPattern(selRange, "\\\[([\s\S]+?)\\\]")

    ' 讓 Word 對剩餘的 OMath 物件做 BuildUp（嘗試建立真正的方程式）
    On Error Resume Next
    selRange.OMaths.BuildUp
    On Error GoTo EH

    MsgBox "已嘗試轉換並建立方程式。", vbInformation, "完成"
    Exit Sub

EH:
    MsgBox "發生錯誤: " & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub

' 處理單一正則模式（從 selRange.Text 找到所有 match，再由後往前在文件中替換）
Private Sub ProcessPattern(ByRef selRange As Range, ByVal pattern As String)
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Multiline = True
    re.IgnoreCase = False
    re.pattern = pattern

    Dim allText As String
    allText = selRange.Text

    Dim matches As Object
    Set matches = re.Execute(allText)

    Dim i As Long
    For i = matches.count - 1 To 0 Step -1
        Dim m As Object
        Set m = matches.Item(i)
        Dim inner As String
        inner = m.SubMatches(0) ' 取出內部內容

        ' 清理 LaTeX 內容 → 轉成 Word 能合理解析的 linear text
        Dim cleaned As String
        cleaned = CleanLatexForWord(inner)

        ' 計算在文件中的絕對位置（selRange.Start + FirstIndex）
        Dim docStart As Long
        docStart = selRange.Start + m.FirstIndex

        Dim docEnd As Long
        docEnd = docStart + m.Length

        ' 建立一個範圍並替換文字（使用 Duplicate 避免改變原selRange）
        Dim r As Range
        Set r = selRange.Duplicate
        r.Start = docStart
        r.End = docEnd

        ' 將整個 match（含 delimiters）替換成 cleaned（不含 $ 或 \( \) 等）
        r.Text = cleaned

        ' 嘗試將該範圍轉成 OMath（容錯處理：若單一失敗則記錄但不終止）
        On Error Resume Next
        Dim om As OMath
        Set om = r.OMaths.Add(r)
        If Err.Number <> 0 Then
            Debug.Print "? 無法將範圍轉為 OMath: "; cleaned; " Err:"; Err.Number; Err.Description
            Err.Clear
        Else
            On Error Resume Next
            om.BuildUp
            Err.Clear
            On Error GoTo 0
        End If
        On Error GoTo 0
    Next i
End Sub

' 清理 LaTeX 內容（把 \text{...} 等轉成純文字、將常見命令替換為符號）
' 並移除 \tag{...}（完全移除，含內文）
Private Function CleanLatexForWord(ByVal s As String) As String
    If Len(s) = 0 Then
        CleanLatexForWord = ""
        Exit Function
    End If

    Dim t As String
    t = s

    ' 保留原有空格，但移除不必要的換行（在選取多段時可能有 vbCr）
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    t = Trim(t)

    ' 先移除 \tag{...}（整個移除）
    t = RemoveTagBlocks(t)

    ' 處理 \text{...}, \mathrm{...}, \operatorname{...} => 剝掉命令，只保留內文
    t = ReplaceCommandBlock(t, "\text{")
    t = ReplaceCommandBlock(t, "\mathrm{")
    t = ReplaceCommandBlock(t, "\operatorname{")
    t = ReplaceCommandBlock(t, "\textrm{")
    t = ReplaceCommandBlock(t, "\textbf{")
    t = ReplaceCommandBlock(t, "\emph{")

    ' 將 \mu -> μ， \times -> ×，\cdot -> ·，保留數學可讀性
    t = Replace(t, "\mu", "μ")
    t = Replace(t, "\times", "×")
    t = Replace(t, "\cdot", "·")
    t = Replace(t, "\,", " ")  ' 拉回空白
    t = Replace(t, "\;", " ")
    t = Replace(t, "~", " ")   ' non-breaking space -> space

    ' 把常見 escaped characters 處理
    t = Replace(t, "\%", "%")
    t = Replace(t, "\$", "$")  ' 如果使用者想要文字 $（雖然在 match 中已被拿掉）
    t = Replace(t, "\\", "\")  ' 兩個反斜線回成一個

    ' 若有 TeX style ^{...} 或 _{...} 保留原樣 (Word 會自行處理)
    ' 去除外層多餘大括號
    If Left(t, 1) = "{" And Right(t, 1) = "}" Then
        t = Mid(t, 2, Len(t) - 2)
    End If

    ' 最後再 Trim 一次，並回傳
    CleanLatexForWord = Trim(t)
End Function

' 輔助：把 \cmd{ ... } 裡的 { } 移除（僅移除最外層，支援巢狀會保守處理）
Private Function ReplaceCommandBlock(ByVal txt As String, ByVal openCmd As String) As String
    Dim res As String
    res = txt

    Dim pos As Long
    pos = InStr(1, res, openCmd, vbBinaryCompare)
    Do While pos > 0
        Dim startPos As Long
        startPos = pos + Len(openCmd) ' 內文開始位置（相對於 res）
        ' 找對應的右大括號（考慮巢狀）
        Dim idx As Long
        idx = startPos
        Dim depth As Long
        depth = 1
        Do While idx <= Len(res) And depth > 0
            Dim ch As String
            ch = Mid(res, idx, 1)
            If ch = "{" Then depth = depth + 1
            If ch = "}" Then depth = depth - 1
            idx = idx + 1
        Loop
        ' idx 現在指向右大括號後一位
        If depth = 0 Then
            Dim innerLen As Long
            innerLen = idx - startPos - 1
            Dim innerTxt As String
            If innerLen > 0 Then
                innerTxt = Mid(res, startPos, innerLen)
            Else
                innerTxt = ""
            End If
            ' 用內文替換整個命令(包含 \cmd{ and } )
            Dim beforeTxt As String
            beforeTxt = Left(res, pos - 1)
            Dim afterTxt As String
            afterTxt = Mid(res, idx)
            res = beforeTxt & innerTxt & afterTxt
            ' 從剛剛替換文字的下一個位置繼續找
            pos = InStr(pos + Len(innerTxt), res, openCmd, vbBinaryCompare)
        Else
            ' 找不到對應的 } 則跳出（避免死循環）
            Exit Do
        End If
    Loop

    ReplaceCommandBlock = res
End Function

' 專門移除 \tag{...}（包含內文），支援巢狀大括號偵測
Private Function RemoveTagBlocks(ByVal txt As String) As String
    Dim res As String
    res = txt

    Dim pos As Long
    pos = InStr(1, res, "\tag{", vbBinaryCompare)
    Do While pos > 0
        Dim startPos As Long
        startPos = pos + Len("\tag{")
        Dim idx As Long
        idx = startPos
        Dim depth As Long
        depth = 1
        Do While idx <= Len(res) And depth > 0
            Dim ch As String
            ch = Mid(res, idx, 1)
            If ch = "{" Then depth = depth + 1
            If ch = "}" Then depth = depth - 1
            idx = idx + 1
        Loop
        If depth = 0 Then
            ' idx 指向右大括號後一位，刪除從 pos 到 idx-1
            Dim beforeTxt As String
            beforeTxt = Left(res, pos - 1)
            Dim afterTxt As String
            afterTxt = Mid(res, idx)
            res = beforeTxt & afterTxt
            ' 繼續尋找下一個 \tag{
            pos = InStr(pos, res, "\tag{", vbBinaryCompare)
        Else
            ' 找不到對應 }，結束以避免死循環
            Exit Do
        End If
    Loop

    RemoveTagBlocks = res
End Function

