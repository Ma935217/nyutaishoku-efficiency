Attribute VB_Name = "modExport"
'==============================================================
' “Íoì¬ƒc[ƒ‹  modExport  v1.0_20261029
'  E‘ä’ i“üE^‘ŞEj‚ÆƒŒƒCƒAƒEƒg’è‹`‚©‚çA“Í‘CSV‚Ü‚½‚Í“]‹Lˆê——‚ğì‚é
'  E€–Ú‚Ì•À‚ÑEŒÅ’è’lEƒR[ƒh’l‚Í‚·‚×‚ÄƒV[ƒg‘¤‚É’u‚­B
'    ‚±‚Ìƒ‚ƒWƒ…[ƒ‹‚Íu“Ç‚ŞE®‚¦‚éEŠm‚©‚ß‚éE‘‚­v‚¾‚¯‚ğs‚¤
'  EŠù‘¶ƒtƒ@ƒCƒ‹‚Ìã‘‚«Eíœ‚Í‚µ‚È‚¢i–ˆ‰ñV‚µ‚¢ƒtƒ@ƒCƒ‹–¼‚Åì‚éj
'  E•ÏX‚µ‚½‚çu•ÏX—š—ğvƒV[ƒg‚É‹L˜^‚·‚é
'==============================================================
Option Explicit

Private Const HDR_ROW As Long = 4          ' ‘ä’ E’è‹`•\‚ÌŒ©o‚µs
Private Const DATA_ROW As Long = 5         ' ƒf[ƒ^‚ÌŠJns
Private Const ERR_ROW As Long = 11         ' o—ÍƒV[ƒg‚ÌƒGƒ‰[ˆê——‚ÌŠJns
Private Const SH_OUT As String = "o—Í"
Private Const SH_LAYOUT As String = "ƒŒƒCƒAƒEƒg’è‹`"
Private Const SH_CODE As String = "ƒR[ƒh•\"
Private Const SH_CONFIRM As String = "“à—eŠm”F"
Private Const SH_LOG As String = "ƒƒO"

Public gTestMode As Boolean                ' True ‚Ì‚Æ‚«ƒƒbƒZ[ƒWƒ{ƒbƒNƒX‚ğo‚³‚È‚¢iŒ±—pj

Private mErr As Long, mWarn As Long, mMsgRow As Long
Private mRow As Long, mStaffId As String, mStaffName As String, mItem As String   ' ƒGƒ‰[•\¦—p‚ÌŒ»İˆÊ’u
Private mCondFalse As Boolean               ' €–Ú‚ÌğŒ‚ª•s¬—§‚¾‚Á‚½‚©

'==================== ƒ{ƒ^ƒ“ ====================

' uƒ`ƒFƒbƒNvFƒtƒ@ƒCƒ‹‚Íì‚ç‚¸AƒGƒ‰[ˆê——‚Æ“à—eŠm”FƒV[ƒg‚¾‚¯‚ğì‚é
Public Sub btnCheck()
    Dim kind As String, info As Variant, lines As Collection, grid As Collection, n As Long
    On Error GoTo EH
    kind = CurrentKind(): If kind = "" Then Exit Sub
    info = KindInfo(kind)
    Application.ScreenUpdating = False
    BuildOutput kind, info, lines, grid, n
    Application.ScreenUpdating = True
    Say "ƒ`ƒFƒbƒNŠ®—¹i" & kind & "jF‘ÎÛ " & n & " Œ^ƒGƒ‰[ " & mErr & " Œ^Œx " & mWarn & " Œ"
    Exit Sub
EH:
    Application.ScreenUpdating = True
    Say "ˆÙíI—¹‚µ‚Ü‚µ‚½F" & Err.Description
End Sub

' uCSVo—ÍvFƒ`ƒFƒbƒN‚ğ’Ê‚Á‚½‚Æ‚«‚¾‚¯ƒtƒ@ƒCƒ‹i‚Ü‚½‚Í“]‹Lˆê——j‚ğì‚é
Public Sub btnExport()
    Dim kind As String, info As Variant, lines As Collection, grid As Collection
    Dim n As Long, nDone As Long, path As String, msg As String
    On Error GoTo EH
    kind = CurrentKind(): If kind = "" Then Exit Sub
    info = KindInfo(kind)
    Application.ScreenUpdating = False
    If Not BuildOutput(kind, info, lines, grid, n) Then
        Application.ScreenUpdating = True
        WriteLog kind, n, "", "ƒGƒ‰[ " & mErr & " Œ‚Ì‚½‚ß–¢o—Í"
        Say "ƒGƒ‰[‚ª " & mErr & " Œ‚ ‚è‚Ü‚·Bo—ÍƒV[ƒg‚Ìˆê——‚ğŒ©‚Ä’¼‚µ‚Ä‚©‚çA‚à‚¤ˆê“xÀs‚µ‚Ä‚­‚¾‚³‚¢Biƒtƒ@ƒCƒ‹‚Íì‚Á‚Ä‚¢‚Ü‚¹‚ñj"
        Exit Sub
    End If
    Application.ScreenUpdating = True
    If n = 0 Then Say "o—Í‘ÎÛi›j‚Å‘ÎÛğŒ‚É‡‚¤s‚ª‚ ‚è‚Ü‚¹‚ñB": Exit Sub

    nDone = CountExported(kind, info, grid)
    msg = kind & " ‚ğ " & n & " Œo—Í‚µ‚Ü‚·B"
    If mWarn > 0 Then msg = msg & vbCrLf & "Œx‚ª " & mWarn & " Œ‚ ‚è‚Ü‚·io—ÍƒV[ƒg‚Ìˆê——‚ğŠm”F‚µ‚Ä‚­‚¾‚³‚¢jB"
    If nDone > 0 Then msg = msg & vbCrLf & "‚¤‚¿ " & nDone & " Œ‚Ío—ÍÏ‚Å‚·i“ñd’ño‚É’ˆÓjB"
    If CStr(info(4)) <> "Ï" Then msg = msg & vbCrLf & "¦‚±‚Ì“Íí•Ê‚ÌƒŒƒCƒAƒEƒg‚Í–¢Šm’è‚Å‚·Bs—p‚Æ‚µ‚Äˆµ‚Á‚Ä‚­‚¾‚³‚¢B"
    If Not Ask(msg & vbCrLf & vbCrLf & "‚æ‚ë‚µ‚¢‚Å‚·‚©H") Then Say "o—Í‚ğæ‚è‚â‚ß‚Ü‚µ‚½B": Exit Sub

    If CStr(info(2)) = "CSV" Then
        path = NewFilePath(kind, CStr(info(3)))
        WriteText lines, path, CStr(Setting("•¶šƒR[ƒh"))
    Else
        path = "u" & SH_CONFIRM & "vƒV[ƒgi“]‹Lˆê——j"
    End If
    MarkExported kind, info, grid
    WriteLog kind, n, path, "¬Œ÷"
    Say n & " Œ‚ğo—Í‚µ‚Ü‚µ‚½F" & path
    Exit Sub
EH:
    Application.ScreenUpdating = True
    WriteLog kind, n, path, "ˆÙíI—¹F" & Err.Description
    Say "ˆÙíI—¹‚µ‚Ü‚µ‚½F" & Err.Description
End Sub

Public Sub btnClearNyushoku()
    ClearLedger "“üE‘ä’ "
End Sub

Public Sub btnClearTaishoku()
    ClearLedger "‘ŞE‘ä’ "
End Sub

' ‰‰ñ‚É1‰ñ‚¾‚¯ÀsFŠeƒV[ƒg‚Éƒ{ƒ^ƒ“‚ğ’u‚­iŠù‚É‚ ‚ê‚Î‰½‚à‚µ‚È‚¢j
Public Sub SetupButtons()
    AddButton SH_OUT, "D3", "ƒ`ƒFƒbƒN", "btnCheck"
    AddButton SH_OUT, "F3", "CSVo—Í", "btnExport"
    AddButton "“üE‘ä’ ", "F1", "‘ä’ ‚ğƒNƒŠƒA", "btnClearNyushoku"
    AddButton "‘ŞE‘ä’ ", "F1", "‘ä’ ‚ğƒNƒŠƒA", "btnClearTaishoku"
    Say "ƒ{ƒ^ƒ“‚ğİ’u‚µ‚Ü‚µ‚½B"
End Sub

'==================== ’†S‚Ìˆ— ====================

' ‘ä’ ‚Ì‘ÎÛs‚ğŒŸ¸‚µACSVsilinesj‚Æ“à—eŠm”F—p‚Ì•\igridj‚ğì‚éBƒGƒ‰[0Œ‚È‚ç True
Private Function BuildOutput(ByVal kind As String, info As Variant, lines As Collection, _
                             grid As Collection, n As Long) As Boolean
    Dim ws As Worksheet, hdr As Variant, lay As Variant
    Dim cMark As Long, cId As Long, cName As Long, r As Long, lastR As Long, isCsv As Boolean

    ResetMessages
    Set lines = New Collection: Set grid = New Collection: n = 0
    isCsv = (CStr(info(2)) = "CSV")
    Set ws = ThisWorkbook.Worksheets(CStr(info(1)))
    hdr = HeaderNames(ws)
    SetCtx 0, "", "", ""

    lay = LoadLayout(kind)
    If IsEmpty(lay) Then GoTo Done

    cMark = ColOf(hdr, "o—Í‘ÎÛ"): cId = ColOf(hdr, "Eˆõ”Ô†"): cName = ColOf(hdr, "–¼")
    If cMark <= 0 Then AddMsg "ƒGƒ‰[", "‘ä’ ‚ÌŒ©o‚µuo—Í‘ÎÛv‚ªŒ©‚Â‚©‚ç‚È‚¢‚©Ad•¡‚µ‚Ä‚¢‚Ü‚·": GoTo Done
    If cId <= 0 Then AddMsg "ƒGƒ‰[", "‘ä’ ‚ÌŒ©o‚µuEˆõ”Ô†v‚ªŒ©‚Â‚©‚ç‚È‚¢‚©Ad•¡‚µ‚Ä‚¢‚Ü‚·iDB“\‚è•t‚¯—“‚ÌŒ©o‚µ‚ğŠm”Fj": GoTo Done

    AddFixedLines lay, "ƒwƒbƒ_[", lines, isCsv
    lastR = ws.Cells(ws.Rows.Count, cId).End(xlUp).Row
    For r = DATA_ROW To lastR
        If IsMarked(ws.Cells(r, cMark).Value) And Trim$(CellText(ws.Cells(r, cId).Value)) <> "" Then
            SetCtx r, CellText(ws.Cells(r, cId).Value), IIf(cName > 0, CellText(ws.Cells(r, cName).Value), ""), "i‘ÎÛğŒj"
            If CondOK(CStr(info(5)), ws, r, hdr) Then
                n = n + 1
                AddDetail lay, ws, r, hdr, lines, grid, isCsv
            Else
                AddMsg "î•ñ", "‘ÎÛğŒi" & CStr(info(5)) & "j‚É‡‚í‚È‚¢‚½‚ßœŠO‚µ‚Ü‚µ‚½"
            End If
        End If
    Next r
    AddFixedLines lay, "ƒtƒbƒ^[", lines, isCsv
    WriteConfirm lay, grid, kind
Done:
    BuildOutput = (mErr = 0)
End Function

' –¾×1si1–¼•ªj‚ğì‚é
Private Sub AddDetail(lay As Variant, ws As Worksheet, ByVal r As Long, hdr As Variant, _
                      lines As Collection, grid As Collection, ByVal isCsv As Boolean)
    Dim i As Long, k As Long, m As Long, s As String, rowVals() As Variant, parts() As String
    For i = 1 To UBound(lay, 1)
        If lay(i, 2) = "–¾×" Then m = m + 1
    Next i
    If m = 0 Then Exit Sub
    ReDim rowVals(0 To m + 2): ReDim parts(1 To m)
    rowVals(0) = r: rowVals(1) = mStaffId: rowVals(2) = mStaffName
    For i = 1 To UBound(lay, 1)
        If lay(i, 2) = "–¾×" Then
            k = k + 1
            mItem = lay(i, 4)
            s = ItemValue(lay, i, ws, r, hdr)
            CheckValue s, lay, i, isCsv
            parts(k) = s: rowVals(k + 2) = s
        End If
    Next i
    lines.Add JoinCsv(parts)
    grid.Add rowVals
End Sub

' ƒwƒbƒ_[^ƒtƒbƒ^[si‘ä’ ‚ğQÆ‚µ‚È‚¢sj‚ğì‚é
Private Sub AddFixedLines(lay As Variant, ByVal grp As String, lines As Collection, ByVal isCsv As Boolean)
    Dim i As Long, cur As String, buf() As String, k As Long, s As String
    SetCtx 0, "", "", ""
    For i = 1 To UBound(lay, 1)
        If Left$(lay(i, 2), Len(grp)) = grp Then
            If lay(i, 2) <> cur Then
                If k > 0 Then lines.Add JoinCsv(buf)
                cur = lay(i, 2): k = 0
            End If
            k = k + 1
            ReDim Preserve buf(1 To k)
            mItem = lay(i, 4)
            s = ItemValue(lay, i, Nothing, 0, Empty)
            CheckValue s, lay, i, isCsv
            buf(k) = s
        End If
    Next i
    If k > 0 Then lines.Add JoinCsv(buf)
End Sub

' 1€–Ú‚Ì’l‚ğæ‚èo‚µA•ÏŠ·E‘®‚ğ“–‚Ä‚é
Private Function ItemValue(lay As Variant, ByVal i As Long, ws As Worksheet, ByVal r As Long, hdr As Variant) As String
    Dim src As String, v As Variant, c As Long, ok As Boolean
    mCondFalse = False
    src = Replace(lay(i, 5), "F", ":")
    If lay(i, 11) <> "" Then
        If r = 0 Then AddMsg "ƒGƒ‰[", "ƒwƒbƒ_[Eƒtƒbƒ^[s‚Å‚ÍuğŒv‚ğg‚¦‚Ü‚¹‚ñ": Exit Function
        If Not CondOK(lay(i, 11), ws, r, hdr) Then mCondFalse = True: Exit Function
    End If

    If src = "ŒÅ’è" Then
        v = lay(i, 6)
    ElseIf src = "‹ó—“" Or src = "" Then
        v = ""
    ElseIf Left$(src, 3) = "‘ä’ :" Then
        If r = 0 Then AddMsg "ƒGƒ‰[", "ƒwƒbƒ_[Eƒtƒbƒ^[s‚Å‚Í‘ä’ ‚Ì’l‚ğg‚¦‚Ü‚¹‚ñ": Exit Function
        c = ColOf(hdr, Mid$(src, 4))
        If c = 0 Then AddMsg "ƒGƒ‰[", "‘ä’ ‚É—ñu" & Mid$(src, 4) & "v‚ª‚ ‚è‚Ü‚¹‚ñiƒŒƒCƒAƒEƒg’è‹`‚Ìæ“¾Œ³‚ğŠm”Fj": Exit Function
        If c < 0 Then AddMsg "ƒGƒ‰[", "‘ä’ ‚Ì—ñu" & Mid$(src, 4) & "v‚ªd•¡‚µ‚Ä‚¢‚Ü‚·": Exit Function
        v = ws.Cells(r, c).Value
    ElseIf Left$(src, 3) = "İ’è:" Then
        v = Setting(Mid$(src, 4), ok)
        If Not ok Then AddMsg "ƒGƒ‰[", "İ’èƒV[ƒg‚É–¼‘Ou" & Mid$(src, 4) & "v‚ª‚ ‚è‚Ü‚¹‚ñ": Exit Function
    Else
        AddMsg "ƒGƒ‰[", "æ“¾Œ³‚Ì‘‚«•û‚ª³‚µ‚­‚ ‚è‚Ü‚¹‚ñF" & src: Exit Function
    End If
    If IsError(v) Then AddMsg "ƒGƒ‰[", "ƒZƒ‹‚ªƒGƒ‰[’li#N/A ‚È‚Çj‚Å‚·": Exit Function

    v = Transform(v, lay(i, 10))
    ItemValue = FormatAs(v, lay(i, 7))
End Function

' •K{EŒ…”Eg‚¦‚È‚¢•¶š‚ğŠm‚©‚ß‚é
Private Sub CheckValue(ByVal s As String, lay As Variant, ByVal i As Long, ByVal isCsv As Boolean)
    Dim lvl As Integer
    If mCondFalse Then Exit Sub
    If lay(i, 9) = "›" And s = "" Then AddMsg "ƒGƒ‰[", "•K{€–Ú‚ª‹ó‚Å‚·" & IIf(lay(i, 12) <> "", "i" & lay(i, 12) & "j", ""): Exit Sub
    If Val(lay(i, 8)) > 0 Then
        If Len(s) > Val(lay(i, 8)) Then AddMsg "ƒGƒ‰[", "Œ…”ƒI[ƒo[‚Å‚·i" & lay(i, 8) & "Œ…‚Ü‚Å^Œ»İ " & Len(s) & "Œ…j"
    End If
    If Not isCsv Or s = "" Then Exit Sub
    If CStr(Setting("ˆø—p•„")) = "‚·‚é" Then
        If InStr(s, """") > 0 Then AddMsg "ƒGƒ‰[", "ƒ_ƒuƒ‹ƒNƒH[ƒe[ƒVƒ‡ƒ“("")‚Íg‚¦‚Ü‚¹‚ñ"
    Else
        If InStr(s, ",") > 0 Or InStr(s, vbCr) > 0 Or InStr(s, vbLf) > 0 Then AddMsg "ƒGƒ‰[", "ƒJƒ“ƒ}E‰üs‚Íg‚¦‚Ü‚¹‚ñ"
    End If
    If CStr(Setting("•¶šƒR[ƒh")) <> "UTF-8" Then
        lvl = SjisLevel(s)
        If lvl = 2 Then
            AddMsg "ƒGƒ‰[", "Shift_JIS‚Å•\‚¹‚È‚¢•¶š‚ª‚ ‚è‚Ü‚·F" & s
        ElseIf lvl = 1 Then
            AddMsg "Œx", "‹@íˆË‘¶•¶ši—áFîàE‡@E‡Šj‚ª‚ ‚è‚Ü‚·B“Í‘ì¬ƒvƒƒOƒ‰ƒ€‚ÅæƒGƒ‰[‚É‚È‚éê‡‚ª‚ ‚è‚Ü‚·F" & s
        End If
    End If
End Sub

'==================== ƒŒƒCƒAƒEƒg’è‹` ====================

' w’è‚µ‚½“Íí•Ê‚Ìs‚ğ (1..n, 1..13) ‚Ì”z—ñ‚Å•Ô‚·B13—ñ–Ú‚Í•À‚×‘Ö‚¦ƒL[
Private Function LoadLayout(ByVal kind As String) As Variant
    Dim ws As Worksheet, v As Variant, r As Long, n As Long, i As Long, j As Long
    Dim lastR As Long, out() As Variant, tmp As Variant, g As Long, prevG As Long, expect As Long
    Set ws = ThisWorkbook.Worksheets(SH_LAYOUT)
    SetCtx 0, "", "", "iƒŒƒCƒAƒEƒg’è‹`j"
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR >= DATA_ROW Then
        v = ws.Range(ws.Cells(DATA_ROW, 1), ws.Cells(lastR, 12)).Value
        For r = 1 To UBound(v, 1)
            If Trim$(CellText(v(r, 1))) = kind Then n = n + 1
        Next r
    End If
    If n = 0 Then AddMsg "ƒGƒ‰[", "ƒŒƒCƒAƒEƒg’è‹`‚Éu" & kind & "v‚Ìs‚ª‚ ‚è‚Ü‚¹‚ñ": Exit Function

    ReDim out(1 To n, 1 To 13)
    For r = 1 To UBound(v, 1)
        If Trim$(CellText(v(r, 1))) = kind Then
            i = i + 1
            For j = 1 To 12: out(i, j) = Trim$(CellText(v(r, j))): Next j
            g = GroupRank(out(i, 2))
            If g < 0 Then AddMsg "ƒGƒ‰[", "sí•Êu" & out(i, 2) & "v‚Íg‚¦‚Ü‚¹‚ñiƒwƒbƒ_[1^–¾×^ƒtƒbƒ^[1 ‚È‚Çj"
            out(i, 13) = g * 10000 + Val(out(i, 3))
        End If
    Next r

    ' •À‚×‘Ö‚¦iŒ”‚ª­‚È‚¢‚Ì‚Å’Pƒ‚È‘}“üƒ\[ƒgj
    For i = 2 To n
        For j = i To 2 Step -1
            If out(j, 13) >= out(j - 1, 13) Then Exit For
            For r = 1 To 13
                tmp = out(j, r): out(j, r) = out(j - 1, r): out(j - 1, r) = tmp
            Next r
        Next j
    Next i

    ' ‡”Ô‚Ísí•Ê‚²‚Æ‚É1‚©‚çŒ‡”ÔEd•¡‚È‚­
    prevG = -999
    For i = 1 To n
        g = Int(out(i, 13) / 10000)
        If g <> prevG Then expect = 1: prevG = g
        If Val(out(i, 3)) <> expect Then
            AddMsg "ƒGƒ‰[", "u" & out(i, 2) & "v‚Ì‡”Ô‚ª1‚©‚ç˜A‘±‚µ‚Ä‚¢‚Ü‚¹‚ñi" & out(i, 4) & "F‡”Ô " & out(i, 3) & "A–{—ˆ " & expect & "j"
            expect = Val(out(i, 3))
        End If
        expect = expect + 1
    Next i
    LoadLayout = out
End Function

Private Function GroupRank(ByVal s As String) As Long
    If s = "–¾×" Then
        GroupRank = 100
    ElseIf Left$(s, 4) = "ƒwƒbƒ_[" And IsNumeric(Mid$(s, 5)) Then
        GroupRank = Val(Mid$(s, 5))
    ElseIf Left$(s, 4) = "ƒtƒbƒ^[" And IsNumeric(Mid$(s, 5)) Then
        GroupRank = 100 + Val(Mid$(s, 5))
    Else
        GroupRank = -1
    End If
End Function

'==================== ’l‚Ì•ÏŠ·E‘® ====================

Private Function Transform(ByVal v As Variant, ByVal conv As String) As Variant
    Dim t As Variant, nm As String
    For Each t In Split(conv, "+")
        t = Trim$(t)
        If t = "" Or CellText(v) = "" Then Exit For
        Select Case t
        Case "‘SŠpƒJƒi": v = StrConv(StrConv(CStr(v), vbKatakana), vbWide)
        Case "”¼Šp": v = StrConv(CStr(v), vbNarrow)
        Case "”š‚Ì‚İ": v = DigitsOnly(CStr(v))
        Case "©": v = NamePart(CStr(v), 1)
        Case "–¼": v = NamePart(CStr(v), 2)
        Case "—‚“ú"
            If IsDate(v) Then v = CDate(v) + 1 Else AddMsg "ƒGƒ‰[", "“ú•t‚Æ‚µ‚Ä“Ç‚ß‚Ü‚¹‚ñF" & CStr(v): v = ""
        Case "ç‰~"
            If IsNumeric(v) Then v = Int(CDbl(v) / 1000) Else AddMsg "ƒGƒ‰[", "”’l‚Å‚Í‚ ‚è‚Ü‚¹‚ñF" & CStr(v): v = ""
        Case Else
            If Left$(t, 5) = "ƒR[ƒh•\:" Then
                v = CodeOf(Mid$(t, 6), Trim$(CStr(v)))
            Else
                AddMsg "ƒGƒ‰[", "•ÏŠ·u" & t & "v‚Íg‚¦‚Ü‚¹‚ñ"
            End If
        End Select
    Next t
    Transform = v
End Function

Private Function FormatAs(ByVal v As Variant, ByVal fmt As String) As String
    Dim d As Date, era As String, code As String, yy As Long, x As Double, h As Long, m As Long
    If CellText(v) = "" Then Exit Function
    Select Case fmt
    Case "", "•¶š"
        FormatAs = CellText(v)
    Case "”’l"
        If IsNumeric(v) Then FormatAs = Format$(CDbl(v), "0") Else AddMsg "ƒGƒ‰[", "”’l‚Å‚Í‚ ‚è‚Ü‚¹‚ñF" & CStr(v)
    Case "ŠÔ•ª"
        If IsNumeric(v) Then
            x = CDbl(v): h = Int(x): m = Round((x - h) * 60, 0)
            FormatAs = h & "ŠÔ" & Format$(m, "00") & "•ª"
        Else
            AddMsg "ƒGƒ‰[", "ŠÔi”’lj‚Å‚Í‚ ‚è‚Ü‚¹‚ñF" & CStr(v)
        End If
    Case "“ú•t8", "“ú•t•\¦", "˜a—ï7", "˜a—ï6", "Œ³†ƒR[ƒh", "˜a—ï•\¦"
        If Not IsDate(v) Then AddMsg "ƒGƒ‰[", "“ú•t‚Æ‚µ‚Ä“Ç‚ß‚Ü‚¹‚ñF" & CStr(v): Exit Function
        d = CDate(v)
        Select Case fmt
        Case "“ú•t8": FormatAs = Format$(d, "yyyymmdd")
        Case "“ú•t•\¦": FormatAs = Format$(d, "yyyy/mm/dd")
        Case Else
            If Not EraOf(d, era, code, yy) Then AddMsg "ƒGƒ‰[", "Œ³†•\‚ÉŠY“–‚·‚éŒ³†‚ª‚ ‚è‚Ü‚¹‚ñF" & Format$(d, "yyyy/mm/dd"): Exit Function
            Select Case fmt
            Case "˜a—ï7": FormatAs = code & Format$(yy, "00") & Format$(d, "mmdd")
            Case "˜a—ï6": FormatAs = Format$(yy, "00") & Format$(d, "mmdd")
            Case "Œ³†ƒR[ƒh": FormatAs = code
            Case "˜a—ï•\¦": FormatAs = era & yy & "”N" & Month(d) & "Œ" & Day(d) & "“ú"
            End Select
        End Select
    Case Else
        AddMsg "ƒGƒ‰[", "‘®u" & fmt & "v‚Íg‚¦‚Ü‚¹‚ñ"
    End Select
End Function

' İ’èƒV[ƒg‚ÌŒ³†•\iŒ³†^ŠJn“ú^ƒR[ƒhj‚©‚çŒ³†‚ğ‹‚ß‚é
Private Function EraOf(ByVal d As Date, era As String, code As String, yy As Long) As Boolean
    Dim t As Variant, i As Long, best As Long
    t = ThisWorkbook.Names("Œ³†•\").RefersToRange.Value
    For i = 1 To UBound(t, 1)
        If IsDate(t(i, 2)) Then
            If d >= CDate(t(i, 2)) Then
                If best = 0 Then
                    best = i
                ElseIf CDate(t(i, 2)) > CDate(t(best, 2)) Then
                    best = i
                End If
            End If
        End If
    Next i
    If best = 0 Then Exit Function
    era = CStr(t(best, 1)): code = CellText(t(best, 3))
    yy = Year(d) - Year(CDate(t(best, 2))) + 1
    EraOf = True
End Function

' ƒR[ƒh•\i‹æ•ª^•\¦’l^CSVƒR[ƒhj‚ğˆø‚­
Private Function CodeOf(ByVal kubun As String, ByVal disp As String) As String
    Dim ws As Worksheet, v As Variant, i As Long, lastR As Long
    Set ws = ThisWorkbook.Worksheets(SH_CODE)
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR >= DATA_ROW Then
        v = ws.Range(ws.Cells(DATA_ROW, 1), ws.Cells(lastR, 3)).Value
        For i = 1 To UBound(v, 1)
            If Trim$(CellText(v(i, 1))) = kubun And Trim$(CellText(v(i, 2))) = disp Then
                CodeOf = Trim$(CellText(v(i, 3)))
                If CodeOf = "" Then AddMsg "ƒGƒ‰[", "ƒR[ƒh•\i" & kubun & "j‚Ìu" & disp & "v‚ÌCSVƒR[ƒh‚ª–¢“ü—Í‚Å‚·i—vŠm”Fj"
                Exit Function
            End If
        Next i
    End If
    AddMsg "ƒGƒ‰[", "ƒR[ƒh•\i" & kubun & "j‚Éu" & disp & "v‚ª‚ ‚è‚Ü‚¹‚ñ"
End Function

Private Function DigitsOnly(ByVal s As String) As String
    Dim i As Long, ch As String
    s = StrConv(s, vbNarrow)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then DigitsOnly = DigitsOnly & ch
    Next i
End Function

' –¼‚ğ©(1)E–¼(2)‚É•ª‚¯‚éB‹æØ‚è‚ÍÅ‰‚Ì‹ó”’i‘SŠpE”¼Špj
Private Function NamePart(ByVal s As String, ByVal part As Integer) As String
    Dim p As Long
    s = Trim$(Replace(s, "@", " "))
    p = InStr(s, " ")
    If p = 0 Then
        If part = 1 Then NamePart = s
        AddMsg "Œx", "–¼‚É‹ó”’‚ª‚È‚¢‚½‚ß©E–¼‚É•ª‚¯‚ç‚ê‚Ü‚¹‚ñF" & s
    ElseIf part = 1 Then
        NamePart = Left$(s, p - 1)
    Else
        NamePart = Trim$(Mid$(s, p + 1))
    End If
End Function

' 0=–â‘è‚È‚µ 1=‹@íˆË‘¶•¶š‚ ‚èiŒxj 2=Shift_JIS‚Å•\‚¹‚È‚¢iƒGƒ‰[j
Private Function SjisLevel(ByVal s As String) As Integer
    Dim i As Long, ch As String, b() As Byte, code As Long
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        b = StrConv(ch, vbFromUnicode)
        If UBound(b) = 0 Then
            If b(0) = 63 And ch <> "?" Then SjisLevel = 2: Exit Function
        ElseIf UBound(b) >= 1 Then
            code = CLng(b(0)) * 256 + b(1)
            If (code >= &H8740& And code <= &H879C&) Or (code >= &HED40& And code <= &HEEFC&) _
               Or (code >= &HFA40& And code <= &HFC4B&) Then SjisLevel = 1
        End If
    Next i
End Function

Private Function JoinCsv(parts() As String) As String
    Dim i As Long, q As String, s As String
    If CStr(Setting("ˆø—p•„")) = "‚·‚é" Then q = """"
    For i = LBound(parts) To UBound(parts)
        If i > LBound(parts) Then s = s & ","
        s = s & q & parts(i) & q
    Next i
    JoinCsv = s
End Function

'==================== ‘ä’ ‚Ì“Ç‚İæ‚è ====================

' Œ©o‚µs‚ğ³‹K‰»‚µ‚½”z—ñi1..ÅI—ñj‚Å•Ô‚·
Private Function HeaderNames(ws As Worksheet) As Variant
    Dim lastC As Long, c As Long, out() As String
    lastC = ws.Cells(HDR_ROW, ws.Columns.Count).End(xlToLeft).Column
    ReDim out(1 To lastC)
    For c = 1 To lastC
        out(c) = NormName(CellText(ws.Cells(HDR_ROW, c).Value))
    Next c
    HeaderNames = out
End Function

' Œ©o‚µ–¼‚©‚ç—ñ”Ô†‚ğ•Ô‚·B‚È‚¯‚ê‚Î0Ad•¡‚µ‚Ä‚¢‚ê‚Î-1
Private Function ColOf(hdr As Variant, ByVal nm As String) As Long
    Dim c As Long, hit As Long
    nm = NormName(nm)
    If nm = "" Then Exit Function
    For c = LBound(hdr) To UBound(hdr)
        If hdr(c) = nm Then
            If hit > 0 Then ColOf = -1: Exit Function
            hit = c
        End If
    Next c
    ColOf = hit
End Function

' ‰üsE‹ó”’Eš•K{‚È‚Ç‚ğœ‚¢‚½Œ©o‚µ–¼i—áFuEˆõ”Ô†(‰üs)š•K{v¨uEˆõ”Ô†vj
Private Function NormName(ByVal s As String) As String
    Dim t As Variant
    For Each t In Array(vbCr, vbLf, " ", "@", "š•K{", "š‘I‘ğ", "š“ü—Í", "š", "(©“®)", "i©“®j")
        s = Replace(s, t, "")
    Next t
    NormName = s
End Function

' u—ñ–¼=’lvu—ñ–¼<>’lvi’l‚Í‹ó‚Å‚à‚æ‚¢j‚ğ”»’è‚·‚éB‹ó‚È‚çí‚É True
Private Function CondOK(ByVal cond As String, ws As Worksheet, ByVal r As Long, hdr As Variant) As Boolean
    Dim p As Long, nm As String, want As String, neg As Boolean, c As Long, cur As String
    cond = Trim$(cond)
    If cond = "" Then CondOK = True: Exit Function
    p = InStr(cond, "<>")
    If p > 0 Then
        neg = True: nm = Trim$(Left$(cond, p - 1)): want = Trim$(Mid$(cond, p + 2))
    Else
        p = InStr(cond, "=")
        If p = 0 Then AddMsg "ƒGƒ‰[", "ğŒ‚Ì‘‚«•û‚ª³‚µ‚­‚ ‚è‚Ü‚¹‚ñF" & cond: Exit Function
        nm = Trim$(Left$(cond, p - 1)): want = Trim$(Mid$(cond, p + 1))
    End If
    c = ColOf(hdr, nm)
    If c <= 0 Then AddMsg "ƒGƒ‰[", "ğŒ‚Ì—ñu" & nm & "v‚ªŒ©‚Â‚©‚ç‚È‚¢‚©Ad•¡‚µ‚Ä‚¢‚Ü‚·": Exit Function
    cur = Trim$(CellText(ws.Cells(r, c).Value))
    CondOK = ((cur = want) Xor neg)
End Function

Private Function IsMarked(ByVal v As Variant) As Boolean
    Dim s As String
    s = Trim$(CellText(v))
    IsMarked = (s = "›" Or s = "Z" Or s = "ü" Or UCase$(s) = "O")
End Function

' ƒGƒ‰[’l‚Å‚à—‚¿‚È‚¢ CStr
Private Function CellText(ByVal v As Variant) As String
    If IsError(v) Then
        CellText = "#ERR"
    ElseIf IsEmpty(v) Or IsNull(v) Then
        CellText = ""
    Else
        CellText = CStr(v)
    End If
End Function

'==================== ‘‚«o‚µ ====================

' “à—eŠm”FƒV[ƒg‚É¡‰ñ‚Ì’l‚ğˆê——‚Å‘‚­i“]‹Lˆê——‚ğŒ“‚Ë‚éj
Private Sub WriteConfirm(lay As Variant, grid As Collection, ByVal kind As String)
    Dim ws As Worksheet, i As Long, k As Long, m As Long, hdr() As Variant, rowVals As Variant
    Set ws = ThisWorkbook.Worksheets(SH_CONFIRM)
    ws.Range("A4:ZZ2000").ClearContents
    ws.Range("A1").Value = kind & "@“à—eŠm”FE“]‹Lˆê——i" & Format$(Now, "yyyy/mm/dd hh:nn") & " ì¬j"
    For i = 1 To UBound(lay, 1)
        If lay(i, 2) = "–¾×" Then m = m + 1
    Next i
    ReDim hdr(1 To m + 3)
    hdr(1) = "‘ä’ s": hdr(2) = "Eˆõ”Ô†": hdr(3) = "–¼"
    k = 3
    For i = 1 To UBound(lay, 1)
        If lay(i, 2) = "–¾×" Then k = k + 1: hdr(k) = lay(i, 4)
    Next i
    ws.Range(ws.Cells(4, 1), ws.Cells(4, m + 3)).Value = hdr
    For i = 1 To grid.Count
        rowVals = grid(i)
        For k = 0 To UBound(rowVals)
            ws.Cells(4 + i, k + 1).NumberFormat = "@"
            ws.Cells(4 + i, k + 1).Value = rowVals(k)
        Next k
    Next i
End Sub

Private Function NewFilePath(ByVal kind As String, ByVal fixedName As String) As String
    Dim folder As String, sep As String, stem As String, path As String, k As Long
    sep = Application.PathSeparator
    folder = Trim$(CStr(Setting("o—ÍƒtƒHƒ‹ƒ_")))
    If folder = "" Then folder = ThisWorkbook.Path & sep & "o—Í"
    If Right$(folder, 1) = sep Then folder = Left$(folder, Len(folder) - 1)
    If Dir(folder, vbDirectory) = "" Then MkDir folder
    stem = folder & sep & kind & "_" & Format$(Now, "yyyymmdd_hhnn")
    If fixedName = "" Then
        path = stem & ".csv": k = 2
        Do While Dir(path) <> ""
            path = stem & "_" & k & ".csv": k = k + 1
        Loop
    Else
        ' d—l‚Åƒtƒ@ƒCƒ‹–¼‚ªŒˆ‚Ü‚Á‚Ä‚¢‚éê‡‚ÍA–ˆ‰ñV‚µ‚¢ƒtƒHƒ‹ƒ_‚ğì‚Á‚Ä‚»‚Ì’†‚É’u‚­
        path = stem: k = 2
        Do While Dir(path, vbDirectory) <> ""
            path = stem & "_" & k: k = k + 1
        Loop
        MkDir path
        path = path & sep & fixedName
    End If
    NewFilePath = path
End Function

Private Sub WriteText(lines As Collection, ByVal path As String, ByVal enc As String)
    Dim f As Integer, i As Long, st As Object, bin As Variant
    If Dir(path) <> "" Then Err.Raise vbObjectError + 1, , "“¯‚¶–¼‘O‚Ìƒtƒ@ƒCƒ‹‚ª‚ ‚é‚½‚ß’†~‚µ‚Ü‚µ‚½F" & path
    If enc = "UTF-8" Then
        Set st = CreateObject("ADODB.Stream")
        st.Type = 2: st.Charset = "UTF-8": st.Open
        For i = 1 To lines.Count: st.WriteText lines(i) & vbCrLf: Next i
        st.Position = 0: st.Type = 1: st.Position = 3       ' æ“ª‚ÌBOM(3ƒoƒCƒg)‚ğœ‚­
        bin = st.Read: st.Close
        Set st = CreateObject("ADODB.Stream")
        st.Type = 1: st.Open: st.Write bin
        st.SaveToFile path, 1                                ' 1=Šù‘¶ƒtƒ@ƒCƒ‹‚ª‚ ‚ê‚Î¸”siã‘‚«‚µ‚È‚¢j
        st.Close
    Else
        f = FreeFile
        Open path For Output As #f                           ' Windows‚ÌŠù’èiShift_JISjECRLF
        For i = 1 To lines.Count: Print #f, lines(i): Next i
        Close #f
    End If
End Sub

Private Function CountExported(ByVal kind As String, info As Variant, grid As Collection) As Long
    Dim ws As Worksheet, c As Long, i As Long
    Set ws = ThisWorkbook.Worksheets(CStr(info(1)))
    c = ColOf(HeaderNames(ws), "o—ÍÏ")
    If c <= 0 Then Exit Function
    For i = 1 To grid.Count
        If InStr(CellText(ws.Cells(grid(i)(0), c).Value), kind & ":") > 0 Then CountExported = CountExported + 1
    Next i
End Function

Private Sub MarkExported(ByVal kind As String, info As Variant, grid As Collection)
    Dim ws As Worksheet, c As Long, i As Long, s As String
    Set ws = ThisWorkbook.Worksheets(CStr(info(1)))
    c = ColOf(HeaderNames(ws), "o—ÍÏ")
    If c <= 0 Then Exit Sub
    For i = 1 To grid.Count
        s = CellText(ws.Cells(grid(i)(0), c).Value)
        If s <> "" Then s = s & " / "
        ws.Cells(grid(i)(0), c).Value = s & kind & ":" & Format$(Now, "yyyy/mm/dd hh:nn")
    Next i
End Sub

Private Sub WriteLog(ByVal kind As String, ByVal n As Long, ByVal path As String, ByVal result As String)
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_LOG)
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If r < DATA_ROW Then r = DATA_ROW
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = Environ("USERNAME")
    ws.Cells(r, 3).Value = kind
    ws.Cells(r, 4).Value = n
    ws.Cells(r, 5).Value = path
    ws.Cells(r, 6).Value = result
End Sub

Private Sub ClearLedger(ByVal sh As String)
    Dim ws As Worksheet, lastR As Long, lastC As Long, rng As Range
    Set ws = ThisWorkbook.Worksheets(sh)
    lastR = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    lastC = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    If lastR < DATA_ROW Then Say "ƒNƒŠƒA‚·‚és‚Í‚ ‚è‚Ü‚¹‚ñB": Exit Sub
    If Not Ask(sh & " ‚Ì " & DATA_ROW & " s–ÚˆÈ~‚Ì“ü—Í’l‚ğÁ‹‚µ‚Ü‚·i”®EŒ©o‚µ‚Íc‚è‚Ü‚·jB" & vbCrLf & _
               "o—Í‚Æ’ño‚ªÏ‚ñ‚Å‚¢‚é‚±‚Æ‚ğŠm”F‚µ‚Ü‚µ‚½‚©H") Then Exit Sub
    On Error Resume Next
    Set rng = ws.Range(ws.Cells(DATA_ROW, 1), ws.Cells(lastR, lastC)).SpecialCells(xlCellTypeConstants)
    On Error GoTo 0
    If Not rng Is Nothing Then rng.ClearContents
    Say sh & " ‚ğƒNƒŠƒA‚µ‚Ü‚µ‚½B"
End Sub

Private Sub AddButton(ByVal sh As String, ByVal addr As String, ByVal caption As String, ByVal macro As String)
    Dim ws As Worksheet, c As Range, b As Object
    Set ws = ThisWorkbook.Worksheets(sh)
    For Each b In ws.Buttons
        If b.Caption = caption Then Exit Sub
    Next b
    Set c = ws.Range(addr)
    Set b = ws.Buttons.Add(c.Left, c.Top, 110, 30)
    b.Caption = caption
    b.OnAction = macro
End Sub

'==================== ¬“¹‹ï ====================

Private Function CurrentKind() As String
    CurrentKind = Trim$(CellText(ThisWorkbook.Names("“Íí•Ê‘I‘ğ").RefersToRange.Value))
    If CurrentKind = "" Then Say "o—ÍƒV[ƒg‚Å“Íí•Ê‚ğ‘I‚ñ‚Å‚­‚¾‚³‚¢B"
End Function

' İ’èƒV[ƒg‚Ì“Íí•Êˆê——‚©‚ç (0)“Íí•Ê (1)‘ÎÛ‘ä’  (2)o—ÍŒ`® (3)ƒtƒ@ƒCƒ‹–¼ (4)Šm’è (5)‘ÎÛğŒ ‚ğ•Ô‚·
Private Function KindInfo(ByVal kind As String) As Variant
    Dim t As Variant, i As Long
    t = ThisWorkbook.Names("“Íí•Êˆê——").RefersToRange.Value
    For i = 1 To UBound(t, 1)
        If Trim$(CellText(t(i, 1))) = kind Then
            KindInfo = Array(kind, Trim$(CellText(t(i, 2))), Trim$(CellText(t(i, 3))), _
                             Trim$(CellText(t(i, 4))), Trim$(CellText(t(i, 5))), Trim$(CellText(t(i, 6))))
            Exit Function
        End If
    Next i
    Err.Raise vbObjectError + 2, , "İ’èƒV[ƒg‚Ì“Íí•Êˆê——‚Éu" & kind & "v‚ª‚ ‚è‚Ü‚¹‚ñ"
End Function

Private Function Setting(ByVal nm As String, Optional ok As Boolean) As Variant
    On Error GoTo NG
    Setting = ThisWorkbook.Names(nm).RefersToRange.Value
    ok = True
    Exit Function
NG:
    ok = False: Setting = ""
End Function

Private Sub SetCtx(ByVal r As Long, ByVal id As String, ByVal nm As String, ByVal itm As String)
    mRow = r: mStaffId = id: mStaffName = nm: mItem = itm
End Sub

Private Sub ResetMessages()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_OUT)
    ws.Range(ws.Cells(ERR_ROW, 1), ws.Cells(ERR_ROW + 2000, 6)).ClearContents
    mErr = 0: mWarn = 0: mMsgRow = ERR_ROW
End Sub

' ‹æ•ªiƒGƒ‰[^Œx^î•ñjE‘ä’ sEEˆõ”Ô†E–¼E€–ÚE“à—e ‚ğo—ÍƒV[ƒg‚É‘‚­
Private Sub AddMsg(ByVal level As String, ByVal msg As String)
    Dim ws As Worksheet
    If level = "ƒGƒ‰[" Then mErr = mErr + 1
    If level = "Œx" Then mWarn = mWarn + 1
    Set ws = ThisWorkbook.Worksheets(SH_OUT)
    ws.Cells(mMsgRow, 1).Value = level
    If mRow > 0 Then ws.Cells(mMsgRow, 2).Value = mRow
    ws.Cells(mMsgRow, 3).NumberFormat = "@"
    ws.Cells(mMsgRow, 3).Value = mStaffId
    ws.Cells(mMsgRow, 4).Value = mStaffName
    ws.Cells(mMsgRow, 5).Value = mItem
    ws.Cells(mMsgRow, 6).Value = msg
    mMsgRow = mMsgRow + 1
End Sub

Private Sub Say(ByVal msg As String)
    On Error Resume Next
    ThisWorkbook.Worksheets(SH_OUT).Range("B6").Value = Format$(Now, "hh:nn") & "@" & msg
    If Not gTestMode Then MsgBox msg, vbInformation, "“Íoì¬ƒc[ƒ‹"
End Sub

Private Function Ask(ByVal msg As String) As Boolean
    If gTestMode Then Ask = True: Exit Function
    Ask = (MsgBox(msg, vbYesNo + vbQuestion, "“Íoì¬ƒc[ƒ‹") = vbYes)
End Function
