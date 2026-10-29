Attribute VB_Name = "modExport"
'==============================================================
' 届出作成ツール  modExport  v1.0_20261029
'  ・台帳（入職／退職）とレイアウト定義から、届書CSVまたは転記一覧を作る
'  ・項目の並び・固定値・コード値はすべてシート側に置く。
'    このモジュールは「読む・整える・確かめる・書く」だけを行う
'  ・既存ファイルの上書き・削除はしない（毎回新しいファイル名で作る）
'  ・変更したら「変更履歴」シートに記録する
'==============================================================
Option Explicit

Private Const HDR_ROW As Long = 4          ' 台帳・定義表の見出し行
Private Const DATA_ROW As Long = 5         ' データの開始行
Private Const ERR_ROW As Long = 11         ' 出力シートのエラー一覧の開始行
Private Const SH_OUT As String = "出力"
Private Const SH_LAYOUT As String = "レイアウト定義"
Private Const SH_CODE As String = "コード表"
Private Const SH_CONFIRM As String = "内容確認"
Private Const SH_LOG As String = "ログ"

Public gTestMode As Boolean                ' True のときメッセージボックスを出さない（試験用）

Private mErr As Long, mWarn As Long, mMsgRow As Long
Private mRow As Long, mStaffId As String, mStaffName As String, mItem As String   ' エラー表示用の現在位置
Private mCondFalse As Boolean               ' 項目の条件が不成立だったか

'==================== ボタン ====================

' 「チェック」：ファイルは作らず、エラー一覧と内容確認シートだけを作る
Public Sub btnCheck()
    Dim kind As String, info As Variant, lines As Collection, grid As Collection, n As Long
    On Error GoTo EH
    kind = CurrentKind(): If kind = "" Then Exit Sub
    info = KindInfo(kind)
    Application.ScreenUpdating = False
    BuildOutput kind, info, lines, grid, n
    Application.ScreenUpdating = True
    Say "チェック完了（" & kind & "）：対象 " & n & " 件／エラー " & mErr & " 件／警告 " & mWarn & " 件"
    Exit Sub
EH:
    Application.ScreenUpdating = True
    Say "異常終了しました：" & Err.Description
End Sub

' 「CSV出力」：チェックを通ったときだけファイル（または転記一覧）を作る
Public Sub btnExport()
    Dim kind As String, info As Variant, lines As Collection, grid As Collection
    Dim n As Long, nDone As Long, path As String, msg As String
    On Error GoTo EH
    kind = CurrentKind(): If kind = "" Then Exit Sub
    info = KindInfo(kind)
    Application.ScreenUpdating = False
    If Not BuildOutput(kind, info, lines, grid, n) Then
        Application.ScreenUpdating = True
        WriteLog kind, n, "", "エラー " & mErr & " 件のため未出力"
        Say "エラーが " & mErr & " 件あります。出力シートの一覧を見て直してから、もう一度実行してください。（ファイルは作っていません）"
        Exit Sub
    End If
    Application.ScreenUpdating = True
    If n = 0 Then Say "出力対象（○）で対象条件に合う行がありません。": Exit Sub

    nDone = CountExported(kind, info, grid)
    msg = kind & " を " & n & " 件出力します。"
    If mWarn > 0 Then msg = msg & vbCrLf & "警告が " & mWarn & " 件あります（出力シートの一覧を確認してください）。"
    If nDone > 0 Then msg = msg & vbCrLf & "うち " & nDone & " 件は出力済です（二重提出に注意）。"
    If CStr(info(4)) <> "済" Then msg = msg & vbCrLf & "※この届種別のレイアウトは未確定です。試行用として扱ってください。"
    If Not Ask(msg & vbCrLf & vbCrLf & "よろしいですか？") Then Say "出力を取りやめました。": Exit Sub

    If CStr(info(2)) = "CSV" Then
        path = NewFilePath(kind, CStr(info(3)))
        WriteText lines, path, CStr(Setting("文字コード"))
    Else
        path = "「" & SH_CONFIRM & "」シート（転記一覧）"
    End If
    MarkExported kind, info, grid
    WriteLog kind, n, path, "成功"
    Say n & " 件を出力しました：" & path
    Exit Sub
EH:
    Application.ScreenUpdating = True
    WriteLog kind, n, path, "異常終了：" & Err.Description
    Say "異常終了しました：" & Err.Description
End Sub

Public Sub btnClearNyushoku()
    ClearLedger "入職台帳"
End Sub

Public Sub btnClearTaishoku()
    ClearLedger "退職台帳"
End Sub

' 初回に1回だけ実行：各シートにボタンを置く（既にあれば何もしない）
Public Sub SetupButtons()
    AddButton SH_OUT, "D3", "チェック", "btnCheck"
    AddButton SH_OUT, "F3", "CSV出力", "btnExport"
    AddButton "入職台帳", "F1", "台帳をクリア", "btnClearNyushoku"
    AddButton "退職台帳", "F1", "台帳をクリア", "btnClearTaishoku"
    Say "ボタンを設置しました。"
End Sub

'==================== 中心の処理 ====================

' 台帳の対象行を検査し、CSV行（lines）と内容確認用の表（grid）を作る。エラー0件なら True
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

    cMark = ColOf(hdr, "出力対象"): cId = ColOf(hdr, "職員番号"): cName = ColOf(hdr, "氏名")
    If cMark <= 0 Then AddMsg "エラー", "台帳の見出し「出力対象」が見つからないか、重複しています": GoTo Done
    If cId <= 0 Then AddMsg "エラー", "台帳の見出し「職員番号」が見つからないか、重複しています（DB貼り付け欄の見出しを確認）": GoTo Done

    AddFixedLines lay, "ヘッダー", lines, isCsv
    lastR = ws.Cells(ws.Rows.Count, cId).End(xlUp).Row
    For r = DATA_ROW To lastR
        If IsMarked(ws.Cells(r, cMark).Value) And Trim$(CellText(ws.Cells(r, cId).Value)) <> "" Then
            SetCtx r, CellText(ws.Cells(r, cId).Value), IIf(cName > 0, CellText(ws.Cells(r, cName).Value), ""), "（対象条件）"
            If CondOK(CStr(info(5)), ws, r, hdr) Then
                n = n + 1
                AddDetail lay, ws, r, hdr, lines, grid, isCsv
            Else
                AddMsg "情報", "対象条件（" & CStr(info(5)) & "）に合わないため除外しました"
            End If
        End If
    Next r
    AddFixedLines lay, "フッター", lines, isCsv
    WriteConfirm lay, grid, kind
Done:
    BuildOutput = (mErr = 0)
End Function

' 明細1行（1名分）を作る
Private Sub AddDetail(lay As Variant, ws As Worksheet, ByVal r As Long, hdr As Variant, _
                      lines As Collection, grid As Collection, ByVal isCsv As Boolean)
    Dim i As Long, k As Long, m As Long, s As String, rowVals() As Variant, parts() As String
    For i = 1 To UBound(lay, 1)
        If lay(i, 2) = "明細" Then m = m + 1
    Next i
    If m = 0 Then Exit Sub
    ReDim rowVals(0 To m + 2): ReDim parts(1 To m)
    rowVals(0) = r: rowVals(1) = mStaffId: rowVals(2) = mStaffName
    For i = 1 To UBound(lay, 1)
        If lay(i, 2) = "明細" Then
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

' ヘッダー／フッター行（台帳を参照しない行）を作る
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

' 1項目の値を取り出し、変換・書式を当てる
Private Function ItemValue(lay As Variant, ByVal i As Long, ws As Worksheet, ByVal r As Long, hdr As Variant) As String
    Dim src As String, v As Variant, c As Long, ok As Boolean
    mCondFalse = False
    src = Replace(lay(i, 5), "：", ":")
    If lay(i, 11) <> "" Then
        If r = 0 Then AddMsg "エラー", "ヘッダー・フッター行では「条件」を使えません": Exit Function
        If Not CondOK(lay(i, 11), ws, r, hdr) Then mCondFalse = True: Exit Function
    End If

    If src = "固定" Then
        v = lay(i, 6)
    ElseIf src = "空欄" Or src = "" Then
        v = ""
    ElseIf Left$(src, 3) = "台帳:" Then
        If r = 0 Then AddMsg "エラー", "ヘッダー・フッター行では台帳の値を使えません": Exit Function
        c = ColOf(hdr, Mid$(src, 4))
        If c = 0 Then AddMsg "エラー", "台帳に列「" & Mid$(src, 4) & "」がありません（レイアウト定義の取得元を確認）": Exit Function
        If c < 0 Then AddMsg "エラー", "台帳の列「" & Mid$(src, 4) & "」が重複しています": Exit Function
        v = ws.Cells(r, c).Value
    ElseIf Left$(src, 3) = "設定:" Then
        v = Setting(Mid$(src, 4), ok)
        If Not ok Then AddMsg "エラー", "設定シートに名前「" & Mid$(src, 4) & "」がありません": Exit Function
    Else
        AddMsg "エラー", "取得元の書き方が正しくありません：" & src: Exit Function
    End If
    If IsError(v) Then AddMsg "エラー", "セルがエラー値（#N/A など）です": Exit Function

    v = Transform(v, lay(i, 10))
    ItemValue = FormatAs(v, lay(i, 7))
End Function

' 必須・桁数・使えない文字を確かめる
Private Sub CheckValue(ByVal s As String, lay As Variant, ByVal i As Long, ByVal isCsv As Boolean)
    Dim lvl As Integer
    If mCondFalse Then Exit Sub
    If lay(i, 9) = "○" And s = "" Then AddMsg "エラー", "必須項目が空です" & IIf(lay(i, 12) <> "", "（" & lay(i, 12) & "）", ""): Exit Sub
    If Val(lay(i, 8)) > 0 Then
        If Len(s) > Val(lay(i, 8)) Then AddMsg "エラー", "桁数オーバーです（" & lay(i, 8) & "桁まで／現在 " & Len(s) & "桁）"
    End If
    If Not isCsv Or s = "" Then Exit Sub
    If CStr(Setting("引用符")) = "する" Then
        If InStr(s, """") > 0 Then AddMsg "エラー", "ダブルクォーテーション("")は使えません"
    Else
        If InStr(s, ",") > 0 Or InStr(s, vbCr) > 0 Or InStr(s, vbLf) > 0 Then AddMsg "エラー", "カンマ・改行は使えません"
    End If
    If CStr(Setting("文字コード")) <> "UTF-8" Then
        lvl = SjisLevel(s)
        If lvl = 2 Then
            AddMsg "エラー", "Shift_JISで表せない文字があります：" & s
        ElseIf lvl = 1 Then
            AddMsg "警告", "機種依存文字（例：髙・①・㈱）があります。届書作成プログラムで取込エラーになる場合があります：" & s
        End If
    End If
End Sub

'==================== レイアウト定義 ====================

' 指定した届種別の行を (1..n, 1..13) の配列で返す。13列目は並べ替えキー
Private Function LoadLayout(ByVal kind As String) As Variant
    Dim ws As Worksheet, v As Variant, r As Long, n As Long, i As Long, j As Long
    Dim lastR As Long, out() As Variant, tmp As Variant, g As Long, prevG As Long, expect As Long
    Set ws = ThisWorkbook.Worksheets(SH_LAYOUT)
    SetCtx 0, "", "", "（レイアウト定義）"
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR >= DATA_ROW Then
        v = ws.Range(ws.Cells(DATA_ROW, 1), ws.Cells(lastR, 12)).Value
        For r = 1 To UBound(v, 1)
            If Trim$(CellText(v(r, 1))) = kind Then n = n + 1
        Next r
    End If
    If n = 0 Then AddMsg "エラー", "レイアウト定義に「" & kind & "」の行がありません": Exit Function

    ReDim out(1 To n, 1 To 13)
    For r = 1 To UBound(v, 1)
        If Trim$(CellText(v(r, 1))) = kind Then
            i = i + 1
            For j = 1 To 12: out(i, j) = Trim$(CellText(v(r, j))): Next j
            g = GroupRank(out(i, 2))
            If g < 0 Then AddMsg "エラー", "行種別「" & out(i, 2) & "」は使えません（ヘッダー1／明細／フッター1 など）"
            out(i, 13) = g * 10000 + Val(out(i, 3))
        End If
    Next r

    ' 並べ替え（件数が少ないので単純な挿入ソート）
    For i = 2 To n
        For j = i To 2 Step -1
            If out(j, 13) >= out(j - 1, 13) Then Exit For
            For r = 1 To 13
                tmp = out(j, r): out(j, r) = out(j - 1, r): out(j - 1, r) = tmp
            Next r
        Next j
    Next i

    ' 順番は行種別ごとに1から欠番・重複なく
    prevG = -999
    For i = 1 To n
        g = Int(out(i, 13) / 10000)
        If g <> prevG Then expect = 1: prevG = g
        If Val(out(i, 3)) <> expect Then
            AddMsg "エラー", "「" & out(i, 2) & "」の順番が1から連続していません（" & out(i, 4) & "：順番 " & out(i, 3) & "、本来 " & expect & "）"
            expect = Val(out(i, 3))
        End If
        expect = expect + 1
    Next i
    LoadLayout = out
End Function

Private Function GroupRank(ByVal s As String) As Long
    If s = "明細" Then
        GroupRank = 100
    ElseIf Left$(s, 4) = "ヘッダー" And IsNumeric(Mid$(s, 5)) Then
        GroupRank = Val(Mid$(s, 5))
    ElseIf Left$(s, 4) = "フッター" And IsNumeric(Mid$(s, 5)) Then
        GroupRank = 100 + Val(Mid$(s, 5))
    Else
        GroupRank = -1
    End If
End Function

'==================== 値の変換・書式 ====================

Private Function Transform(ByVal v As Variant, ByVal conv As String) As Variant
    Dim t As Variant, nm As String
    For Each t In Split(conv, "+")
        t = Trim$(t)
        If t = "" Or CellText(v) = "" Then Exit For
        Select Case t
        Case "全角カナ": v = StrConv(StrConv(CStr(v), vbKatakana), vbWide)
        Case "半角": v = StrConv(CStr(v), vbNarrow)
        Case "数字のみ": v = DigitsOnly(CStr(v))
        Case "姓": v = NamePart(CStr(v), 1)
        Case "名": v = NamePart(CStr(v), 2)
        Case "翌日"
            If IsDate(v) Then v = CDate(v) + 1 Else AddMsg "エラー", "日付として読めません：" & CStr(v): v = ""
        Case "千円"
            If IsNumeric(v) Then v = Int(CDbl(v) / 1000) Else AddMsg "エラー", "数値ではありません：" & CStr(v): v = ""
        Case Else
            If Left$(t, 5) = "コード表:" Then
                v = CodeOf(Mid$(t, 6), Trim$(CStr(v)))
            Else
                AddMsg "エラー", "変換「" & t & "」は使えません"
            End If
        End Select
    Next t
    Transform = v
End Function

Private Function FormatAs(ByVal v As Variant, ByVal fmt As String) As String
    Dim d As Date, era As String, code As String, yy As Long, x As Double, h As Long, m As Long
    If CellText(v) = "" Then Exit Function
    Select Case fmt
    Case "", "文字"
        FormatAs = CellText(v)
    Case "数値"
        If IsNumeric(v) Then FormatAs = Format$(CDbl(v), "0") Else AddMsg "エラー", "数値ではありません：" & CStr(v)
    Case "時間分"
        If IsNumeric(v) Then
            x = CDbl(v): h = Int(x): m = Round((x - h) * 60, 0)
            FormatAs = h & "時間" & Format$(m, "00") & "分"
        Else
            AddMsg "エラー", "時間（数値）ではありません：" & CStr(v)
        End If
    Case "日付8", "日付表示", "和暦7", "和暦6", "元号コード", "和暦表示"
        If Not IsDate(v) Then AddMsg "エラー", "日付として読めません：" & CStr(v): Exit Function
        d = CDate(v)
        Select Case fmt
        Case "日付8": FormatAs = Format$(d, "yyyymmdd")
        Case "日付表示": FormatAs = Format$(d, "yyyy/mm/dd")
        Case Else
            If Not EraOf(d, era, code, yy) Then AddMsg "エラー", "元号表に該当する元号がありません：" & Format$(d, "yyyy/mm/dd"): Exit Function
            Select Case fmt
            Case "和暦7": FormatAs = code & Format$(yy, "00") & Format$(d, "mmdd")
            Case "和暦6": FormatAs = Format$(yy, "00") & Format$(d, "mmdd")
            Case "元号コード": FormatAs = code
            Case "和暦表示": FormatAs = era & yy & "年" & Month(d) & "月" & Day(d) & "日"
            End Select
        End Select
    Case Else
        AddMsg "エラー", "書式「" & fmt & "」は使えません"
    End Select
End Function

' 設定シートの元号表（元号／開始日／コード）から元号を求める
Private Function EraOf(ByVal d As Date, era As String, code As String, yy As Long) As Boolean
    Dim t As Variant, i As Long, best As Long
    t = ThisWorkbook.Names("元号表").RefersToRange.Value
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

' コード表（区分／表示値／CSVコード）を引く
Private Function CodeOf(ByVal kubun As String, ByVal disp As String) As String
    Dim ws As Worksheet, v As Variant, i As Long, lastR As Long
    Set ws = ThisWorkbook.Worksheets(SH_CODE)
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR >= DATA_ROW Then
        v = ws.Range(ws.Cells(DATA_ROW, 1), ws.Cells(lastR, 3)).Value
        For i = 1 To UBound(v, 1)
            If Trim$(CellText(v(i, 1))) = kubun And Trim$(CellText(v(i, 2))) = disp Then
                CodeOf = Trim$(CellText(v(i, 3)))
                If CodeOf = "" Then AddMsg "エラー", "コード表（" & kubun & "）の「" & disp & "」のCSVコードが未入力です（要確認）"
                Exit Function
            End If
        Next i
    End If
    AddMsg "エラー", "コード表（" & kubun & "）に「" & disp & "」がありません"
End Function

Private Function DigitsOnly(ByVal s As String) As String
    Dim i As Long, ch As String
    s = StrConv(s, vbNarrow)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then DigitsOnly = DigitsOnly & ch
    Next i
End Function

' 氏名を姓(1)・名(2)に分ける。区切りは最初の空白（全角・半角）
Private Function NamePart(ByVal s As String, ByVal part As Integer) As String
    Dim p As Long
    s = Trim$(Replace(s, "　", " "))
    p = InStr(s, " ")
    If p = 0 Then
        If part = 1 Then NamePart = s
        AddMsg "警告", "氏名に空白がないため姓・名に分けられません：" & s
    ElseIf part = 1 Then
        NamePart = Left$(s, p - 1)
    Else
        NamePart = Trim$(Mid$(s, p + 1))
    End If
End Function

' 0=問題なし 1=機種依存文字あり（警告） 2=Shift_JISで表せない（エラー）
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
    If CStr(Setting("引用符")) = "する" Then q = """"
    For i = LBound(parts) To UBound(parts)
        If i > LBound(parts) Then s = s & ","
        s = s & q & parts(i) & q
    Next i
    JoinCsv = s
End Function

'==================== 台帳の読み取り ====================

' 見出し行を正規化した配列（1..最終列）で返す
Private Function HeaderNames(ws As Worksheet) As Variant
    Dim lastC As Long, c As Long, out() As String
    lastC = ws.Cells(HDR_ROW, ws.Columns.Count).End(xlToLeft).Column
    ReDim out(1 To lastC)
    For c = 1 To lastC
        out(c) = NormName(CellText(ws.Cells(HDR_ROW, c).Value))
    Next c
    HeaderNames = out
End Function

' 見出し名から列番号を返す。なければ0、重複していれば-1
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

' 改行・空白・★必須などを除いた見出し名（例：「職員番号(改行)★必須」→「職員番号」）
Private Function NormName(ByVal s As String) As String
    Dim t As Variant
    For Each t In Array(vbCr, vbLf, " ", "　", "★必須", "★選択", "★入力", "★", "(自動)", "（自動）")
        s = Replace(s, t, "")
    Next t
    NormName = s
End Function

' 「列名=値」「列名<>値」（値は空でもよい）を判定する。空なら常に True
Private Function CondOK(ByVal cond As String, ws As Worksheet, ByVal r As Long, hdr As Variant) As Boolean
    Dim p As Long, nm As String, want As String, neg As Boolean, c As Long, cur As String
    cond = Trim$(cond)
    If cond = "" Then CondOK = True: Exit Function
    p = InStr(cond, "<>")
    If p > 0 Then
        neg = True: nm = Trim$(Left$(cond, p - 1)): want = Trim$(Mid$(cond, p + 2))
    Else
        p = InStr(cond, "=")
        If p = 0 Then AddMsg "エラー", "条件の書き方が正しくありません：" & cond: Exit Function
        nm = Trim$(Left$(cond, p - 1)): want = Trim$(Mid$(cond, p + 1))
    End If
    c = ColOf(hdr, nm)
    If c <= 0 Then AddMsg "エラー", "条件の列「" & nm & "」が見つからないか、重複しています": Exit Function
    cur = Trim$(CellText(ws.Cells(r, c).Value))
    CondOK = ((cur = want) Xor neg)
End Function

Private Function IsMarked(ByVal v As Variant) As Boolean
    Dim s As String
    s = Trim$(CellText(v))
    IsMarked = (s = "○" Or s = "〇" Or s = "◯" Or UCase$(s) = "O")
End Function

' エラー値でも落ちない CStr
Private Function CellText(ByVal v As Variant) As String
    If IsError(v) Then
        CellText = "#ERR"
    ElseIf IsEmpty(v) Or IsNull(v) Then
        CellText = ""
    Else
        CellText = CStr(v)
    End If
End Function

'==================== 書き出し ====================

' 内容確認シートに今回の値を一覧で書く（転記一覧を兼ねる）
Private Sub WriteConfirm(lay As Variant, grid As Collection, ByVal kind As String)
    Dim ws As Worksheet, i As Long, k As Long, m As Long, hdr() As Variant, rowVals As Variant
    Set ws = ThisWorkbook.Worksheets(SH_CONFIRM)
    ws.Range("A4:ZZ2000").ClearContents
    ws.Range("A1").Value = kind & "　内容確認・転記一覧（" & Format$(Now, "yyyy/mm/dd hh:nn") & " 作成）"
    For i = 1 To UBound(lay, 1)
        If lay(i, 2) = "明細" Then m = m + 1
    Next i
    ReDim hdr(1 To m + 3)
    hdr(1) = "台帳行": hdr(2) = "職員番号": hdr(3) = "氏名"
    k = 3
    For i = 1 To UBound(lay, 1)
        If lay(i, 2) = "明細" Then k = k + 1: hdr(k) = lay(i, 4)
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
    folder = Trim$(CStr(Setting("出力フォルダ")))
    If folder = "" Then folder = ThisWorkbook.Path & sep & "出力"
    If Right$(folder, 1) = sep Then folder = Left$(folder, Len(folder) - 1)
    If Dir(folder, vbDirectory) = "" Then MkDir folder
    stem = folder & sep & kind & "_" & Format$(Now, "yyyymmdd_hhnn")
    If fixedName = "" Then
        path = stem & ".csv": k = 2
        Do While Dir(path) <> ""
            path = stem & "_" & k & ".csv": k = k + 1
        Loop
    Else
        ' 仕様でファイル名が決まっている場合は、毎回新しいフォルダを作ってその中に置く
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
    If Dir(path) <> "" Then Err.Raise vbObjectError + 1, , "同じ名前のファイルがあるため中止しました：" & path
    If enc = "UTF-8" Then
        Set st = CreateObject("ADODB.Stream")
        st.Type = 2: st.Charset = "UTF-8": st.Open
        For i = 1 To lines.Count: st.WriteText lines(i) & vbCrLf: Next i
        st.Position = 0: st.Type = 1: st.Position = 3       ' 先頭のBOM(3バイト)を除く
        bin = st.Read: st.Close
        Set st = CreateObject("ADODB.Stream")
        st.Type = 1: st.Open: st.Write bin
        st.SaveToFile path, 1                                ' 1=既存ファイルがあれば失敗（上書きしない）
        st.Close
    Else
        f = FreeFile
        Open path For Output As #f                           ' Windowsの既定（Shift_JIS）・CRLF
        For i = 1 To lines.Count: Print #f, lines(i): Next i
        Close #f
    End If
End Sub

Private Function CountExported(ByVal kind As String, info As Variant, grid As Collection) As Long
    Dim ws As Worksheet, c As Long, i As Long
    Set ws = ThisWorkbook.Worksheets(CStr(info(1)))
    c = ColOf(HeaderNames(ws), "出力済")
    If c <= 0 Then Exit Function
    For i = 1 To grid.Count
        If InStr(CellText(ws.Cells(grid(i)(0), c).Value), kind & ":") > 0 Then CountExported = CountExported + 1
    Next i
End Function

Private Sub MarkExported(ByVal kind As String, info As Variant, grid As Collection)
    Dim ws As Worksheet, c As Long, i As Long, s As String
    Set ws = ThisWorkbook.Worksheets(CStr(info(1)))
    c = ColOf(HeaderNames(ws), "出力済")
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
    If lastR < DATA_ROW Then Say "クリアする行はありません。": Exit Sub
    If Not Ask(sh & " の " & DATA_ROW & " 行目以降の入力値を消去します（数式・見出しは残ります）。" & vbCrLf & _
               "出力と提出が済んでいることを確認しましたか？") Then Exit Sub
    On Error Resume Next
    Set rng = ws.Range(ws.Cells(DATA_ROW, 1), ws.Cells(lastR, lastC)).SpecialCells(xlCellTypeConstants)
    On Error GoTo 0
    If Not rng Is Nothing Then rng.ClearContents
    Say sh & " をクリアしました。"
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

'==================== 小道具 ====================

Private Function CurrentKind() As String
    CurrentKind = Trim$(CellText(ThisWorkbook.Names("届種別選択").RefersToRange.Value))
    If CurrentKind = "" Then Say "出力シートで届種別を選んでください。"
End Function

' 設定シートの届種別一覧から (0)届種別 (1)対象台帳 (2)出力形式 (3)ファイル名 (4)確定 (5)対象条件 を返す
Private Function KindInfo(ByVal kind As String) As Variant
    Dim t As Variant, i As Long
    t = ThisWorkbook.Names("届種別一覧").RefersToRange.Value
    For i = 1 To UBound(t, 1)
        If Trim$(CellText(t(i, 1))) = kind Then
            KindInfo = Array(kind, Trim$(CellText(t(i, 2))), Trim$(CellText(t(i, 3))), _
                             Trim$(CellText(t(i, 4))), Trim$(CellText(t(i, 5))), Trim$(CellText(t(i, 6))))
            Exit Function
        End If
    Next i
    Err.Raise vbObjectError + 2, , "設定シートの届種別一覧に「" & kind & "」がありません"
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

' 区分（エラー／警告／情報）・台帳行・職員番号・氏名・項目・内容 を出力シートに書く
Private Sub AddMsg(ByVal level As String, ByVal msg As String)
    Dim ws As Worksheet
    If level = "エラー" Then mErr = mErr + 1
    If level = "警告" Then mWarn = mWarn + 1
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
    ThisWorkbook.Worksheets(SH_OUT).Range("B6").Value = Format$(Now, "hh:nn") & "　" & msg
    If Not gTestMode Then MsgBox msg, vbInformation, "届出作成ツール"
End Sub

Private Function Ask(ByVal msg As String) As Boolean
    If gTestMode Then Ask = True: Exit Function
    Ask = (MsgBox(msg, vbYesNo + vbQuestion, "届出作成ツール") = vbYes)
End Function
