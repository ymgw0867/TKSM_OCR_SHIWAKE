Imports 仕訳伝票.PCData

Public Class frmInfo
    Dim fr As New Font("ＭＳゴシック", PRINTFONTSIZE, FontStyle.Regular)
    Dim fb As New Font("ＭＳゴシック", PRINTFONTSIZE + 3, FontStyle.Bold)
    Dim fl As New Font("ＭＳゴシック", PRINTFONTSIZE + 3, FontStyle.Underline)

    Public Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument.PrintPage
        '----------------------------------------------------------------
        '   伝票印刷
        '----------------------------------------------------------------
        'Public Sub PrintDen(ByVal Den As Integer, ByVal Mode As Integer)
        Dim wrkWord As String
        Dim wrkXBase As Integer
        Dim wrkYBase As Integer
        Dim wrkNowX As Integer
        Dim wrkNowY As Integer
        Dim Gyou As Integer
        Dim KariSum As Decimal
        Dim KashiSum As Decimal
        Dim KariMinSum As Decimal
        Dim KashiMinSum As Decimal
        Dim Loopcnt As Integer
        Dim Cnt As Integer
        Dim wrkFirstDen As Integer
        Dim EndFlg As Boolean
        Dim wrkErrNo As Integer
        Dim ErrFlg As Boolean
        Dim wrkPrintDen As Integer

        Dim WritePen As New Pen(Color.Black, PRINTFONTSIZE)
        Dim WritePoint As Point()
        Dim pf As New PCfunc

        On Error GoTo ErrPrc

        '印刷設定
        e.PageSettings.PaperSize.PaperName = Printing.PaperKind.A4
        e.PageSettings.Landscape = True
        wrkXBase = PRINTXBASE
        wrkYBase = PRINTYBASE
        wrkNowX = wrkXBase
        wrkNowY = wrkYBase

        '----------------------------------------------------
        ' 複数チェック付きの伝票を印刷
        '----------------------------------------------------

        '複写チェックがなくなるか、先頭の伝票まで前に戻る
        For Cnt = pblNowden To 1 Step -1
            If ((pblDenRec(Cnt).Head.FukusuChk = "0") Or (Cnt = 1)) Then
                wrkFirstDen = Cnt
                Exit For
            End If
        Next

        Loopcnt = 0

        KariSum = 0
        KashiSum = 0

        ' 複写チェックがなくなるまでループ
        Do While (1)
            '5枚に一度ヘッダ書き込み
            If ((Loopcnt Mod PRINTMAXGYOU) = 0) Then
                '伝票ヘッダ出力
                Call WritePrintHead("伝票認識内容", pblNowden, wrkNowX, wrkNowY, e)

                'ページ番号
                wrkNowX = wrkXBase + 160
                'Call SetXY(wrkNowX, wrkNowY)
                e.Graphics.DrawString("Page " & CStr(Fix(Loopcnt / PRINTMAXGYOU) + 1), fr, Brushes.Black, wrkNowX, wrkNowY)
                'printer.Print("Page " & CStr(Fix(Loopcnt / PRINTMAXGYOU) + 1))

                wrkNowY = wrkNowY + 1

                'ライン
                wrkNowY = wrkNowY + 1

                '伝票行ヘッダ
                wrkNowY = wrkNowY + 1
                wrkNowX = wrkXBase
                Call WriteGyouHead(wrkNowX, wrkNowY, e)

                'ライン
                wrkNowY = wrkNowY + 3
                'Call SetXY(wrkXBase, wrkNowY)

                WritePoint = {New Point(wrkNowX, wrkNowY), New Point(wrkNowX + 15500, wrkNowY)}
                e.Graphics.DrawLines(WritePen, WritePoint)
                'Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX + 15500, Printer.CurrentY)

            End If

            '１枚印刷なら、引数の伝票のみ印刷
            If (pblPrintMode = PRINTMODEONE) Then
                wrkPrintDen = pblNowden
            Else
                wrkPrintDen = wrkFirstDen + Loopcnt
            End If


            '伝票ヘッダデータ書込み
            wrkNowY = wrkNowY + 2
            wrkNowX = wrkXBase
            Call WriteDenHead(wrkPrintDen, wrkNowX, wrkNowY, e)

            '小計金額のクリア
            KariMinSum = 0
            KashiMinSum = 0

            '伝票行データ
            For Gyou = 1 To MAXGYOU
                wrkNowX = wrkXBase
                wrkNowY = wrkNowY + 1
                If pblDenRec(wrkPrintDen).Gyou(Gyou).Torikeshi = "0" Then

                    Call WriteGyouData(wrkPrintDen, Gyou, wrkNowX, wrkNowY, e)

                    KariSum = KariSum + CDec(Val(pblDenRec(wrkPrintDen).Gyou(Gyou).Kari.Kin))
                    KashiSum = KashiSum + CDec(Val(pblDenRec(wrkPrintDen).Gyou(Gyou).Kashi.Kin))
                    KariMinSum = KariMinSum + CDec(Val(pblDenRec(wrkPrintDen).Gyou(Gyou).Kari.Kin))
                    KashiMinSum = KashiMinSum + CDec(Val(pblDenRec(wrkPrintDen).Gyou(Gyou).Kashi.Kin))
                End If
            Next

            'ライン
            wrkNowY = wrkNowY + 1
            WritePoint = {New Point(wrkXBase, wrkNowY), New Point(wrkXBase + 15500, wrkNowY)}
            e.Graphics.DrawLines(WritePen, WritePoint)
            'Call SetXY(wrkXBase, wrkNowY)
            'Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX + 15500, Printer.CurrentY)

            wrkNowX = wrkXBase
            '小計金額記入
            Call WriteMinSum(KariMinSum, KashiMinSum, wrkNowX, wrkNowY, e)

            '貸借ライン
            WritePoint = {New Point(76, wrkNowY - 7 + 50), New Point(76 + 15500, wrkNowY - 7 + 50 + 1105)}
            e.Graphics.DrawLines(WritePen, WritePoint)
            'Call SetXY(76, wrkNowY - 7)
            'Printer.Line (Printer.CurrentX, Printer.CurrentY + 50)-(Printer.CurrentX, Printer.CurrentY + 1105)

            '摘要ライン
            WritePoint = {New Point(148, wrkNowY - 7 + 50), New Point(148, wrkNowY - 7 + 50 + 1105)}
            e.Graphics.DrawLines(WritePen, WritePoint)
            'Call SetXY(148, wrkNowY - 7)
            'Printer.Line (Printer.CurrentX, Printer.CurrentY + 50)-(Printer.CurrentX, Printer.CurrentY + 1105)

            Loopcnt = Loopcnt + 1

            '印刷終了判定
            EndFlg = False
            '１枚印刷モードのとき
            If (pblPrintMode = PRINTMODEONE) Then
                EndFlg = True

                '最終伝票に達したとき
            ElseIf (wrkFirstDen + Loopcnt > pblDenNum) Then
                EndFlg = True

                '複数チェックがなくなったとき
            ElseIf (pblDenRec(wrkFirstDen + Loopcnt).Head.FukusuChk = "0") Then
                EndFlg = True
            End If

            If (EndFlg = True) Then
                '合計金額
                wrkNowY = wrkNowY + 2
                wrkNowX = wrkXBase

                wrkNowX = wrkNowX + 4
                'Call SetXY(wrkNowX, wrkNowY)
                e.Graphics.DrawString("金額合計：", fr, Brushes.Black, wrkNowX, wrkNowY)
                'printer.Print("金額合計：")

                wrkNowX = wrkNowX + 45
                'Call SetXY(wrkNowX, wrkNowY)
                wrkWord = pf.AddKanma(CStr(KariSum), 3)
                wrkWord = pf.AddSpace(wrkWord, 14)
                e.Graphics.DrawString(wrkWord, fr, Brushes.Black, wrkNowX, wrkNowY)
                'printer.Print(wrkWord)

                wrkNowX = wrkNowX + 73
                'Call SetXY(wrkNowX, wrkNowY)
                wrkWord = pf.AddKanma(CStr(KashiSum), 3)
                wrkWord = pf.AddSpace(wrkWord, 14)
                e.Graphics.DrawString(wrkWord, fr, Brushes.Black, wrkNowX, wrkNowY)
                'printer.Print(wrkWord)

                wrkNowX = wrkNowX + 20
                'Call SetXY(wrkNowX, wrkNowY)
                e.Graphics.DrawString("貸借差額：" & CStr(Math.Abs(KariSum - KashiSum)), fr, Brushes.Black, wrkNowX, wrkNowY)
                'printer.Print("貸借差額：" & CStr(Abs(KariSum - KashiSum)))
                Exit Do
            End If

            '5枚に1回改ページ
            If ((Loopcnt Mod PRINTMAXGYOU) = 0) Then
                e.HasMorePages = True
                'printer.NewPage()
                wrkNowY = wrkYBase
            End If
        Loop
        'printer.EndDoc()
        e.HasMorePages = False

        On Error GoTo 0

        Exit Sub

        ' エラー処理
ErrPrc:
        MessageBox.Show("印刷中に不具合が発生したため印刷を中断します", Title)
        Exit Sub

    End Sub


    '----------------------------------------------------------------
    '   X座標Y座標設定
    '----------------------------------------------------------------
    'Private Sub SetXY(ByVal X As Integer, ByVal Y As Integer)
    '    e.
    '    Printer.CurrentX = Printer.TextWidth("-") * X
    '    Printer.CurrentY = Printer.TextHeight("-") * Y
    'End Sub

    '----------------------------------------------------------------
    '   印刷ヘッダ出力
    '----------------------------------------------------------------
    Private Sub WritePrintHead(ByVal Title As String, ByVal Den As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal e As System.Drawing.Printing.PrintPageEventArgs)
        Dim wrkNowX As Integer
        Dim wrkNowY As Integer

        wrkNowX = X + 80 - Int(Len(Title))
        wrkNowY = Y
        'BOLDで印刷したのちアンダーラインを引く
        e.Graphics.DrawString(Title, fb, Brushes.Black, wrkNowX, wrkNowY)
        e.Graphics.DrawString(Title, fl, Brushes.Black, wrkNowX, wrkNowY)
        'Call SetXY(wrkNowX, wrkNowY)
        'Printer.FontBold = True
        'Printer.FontUnderline = True
        'Printer.FontSize = PRINTFONTSIZE + 3
        'Printer.Print(Title)
        'Printer.FontBold = False
        'Printer.FontUnderline = False
        'Printer.FontSize = PRINTFONTSIZE

    End Sub

    '----------------------------------------------------------------
    '   行ヘッダ部出力
    '----------------------------------------------------------------
    Private Sub WriteGyouHead(ByVal X As Integer, ByVal Y As Integer, ByVal e As System.Drawing.Printing.PrintPageEventArgs)
        Dim wrkNowX As Integer
        Dim wrkNowY As Integer
        
        wrkNowY = Y
        wrkNowX = X

        wrkNowX = wrkNowX + 3
        e.Graphics.DrawString("［借　　方］", fr, Brushes.Black, wrkNowX, wrkNowY)
        'Call SetXY(wrkNowX, wrkNowY)
        'Printer.Print("［借　　方］")

        wrkNowX = wrkNowX + 74
        e.Graphics.DrawString("［貸　　方］", fr, Brushes.Black, wrkNowX, wrkNowY)
        'Call SetXY(wrkNowX, wrkNowY)
        'Printer.Print("［貸　　方］")

        wrkNowY = Y + 2
        wrkNowX = X

        wrkNowX = wrkNowX + 1
        e.Graphics.DrawString("部門", fr, Brushes.Black, wrkNowX, wrkNowY)
        'Call SetXY(wrkNowX, wrkNowY)
        'Printer.Print("部門")

        wrkNowX = wrkNowX + 6
        e.Graphics.DrawString("科目", fr, Brushes.Black, wrkNowX, wrkNowY)
        'Call SetXY(wrkNowX, wrkNowY)
        'Printer.Print("科目")

        wrkNowX = wrkNowX + 21
        e.Graphics.DrawString("補助", fr, Brushes.Black, wrkNowX, wrkNowY)
        'Call SetXY(wrkNowX, wrkNowY)
        'Printer.Print("補助")

        wrkNowX = wrkNowX + 31
        e.Graphics.DrawString("金額", fr, Brushes.Black, wrkNowX, wrkNowY)
        'Call SetXY(wrkNowX, wrkNowY)
        'Printer.Print("金額")

        wrkNowX = wrkNowX + 6
        e.Graphics.DrawString("処", fr, Brushes.Black, wrkNowX, wrkNowY)
        'Call SetXY(wrkNowX, wrkNowY)
        'Printer.Print("処")

        wrkNowX = wrkNowX + 3
        e.Graphics.DrawString("区", fr, Brushes.Black, wrkNowX, wrkNowY)
        'Call SetXY(wrkNowX, wrkNowY)
        'Printer.Print("区")

        wrkNowX = wrkNowX + 6
        e.Graphics.DrawString("部門", fr, Brushes.Black, wrkNowX, wrkNowY)
        'Call SetXY(wrkNowX, wrkNowY)
        'Printer.Print("部門")

        wrkNowX = wrkNowX + 6
        e.Graphics.DrawString("科目", fr, Brushes.Black, wrkNowX, wrkNowY)
        'Call SetXY(wrkNowX, wrkNowY)
        'Printer.Print("科目")

        wrkNowX = wrkNowX + 21
        e.Graphics.DrawString("補助", fr, Brushes.Black, wrkNowX, wrkNowY)
        'Call SetXY(wrkNowX, wrkNowY)
        'Printer.Print("補助")

        wrkNowX = wrkNowX + 31
        e.Graphics.DrawString("金額", fr, Brushes.Black, wrkNowX, wrkNowY)
        'Call SetXY(wrkNowX, wrkNowY)
        'Printer.Print("金額")

        wrkNowX = wrkNowX + 6
        e.Graphics.DrawString("処", fr, Brushes.Black, wrkNowX, wrkNowY)
        'Call SetXY(wrkNowX, wrkNowY)
        'Printer.Print("処")

        wrkNowX = wrkNowX + 3
        e.Graphics.DrawString("区", fr, Brushes.Black, wrkNowX, wrkNowY)
        'Call SetXY(wrkNowX, wrkNowY)
        'Printer.Print("区")

        wrkNowX = wrkNowX + 4
        e.Graphics.DrawString("複写", fr, Brushes.Black, wrkNowX, wrkNowY)
        'Call SetXY(wrkNowX, wrkNowY)
        'Printer.Print("複写")

        wrkNowX = wrkNowX + 6
        e.Graphics.DrawString("摘要", fr, Brushes.Black, wrkNowX, wrkNowY)
        'Call SetXY(wrkNowX, wrkNowY)
        'Printer.Print("摘要")

    End Sub

    '----------------------------------------------------------------
    '   伝票ヘッダデータ出力
    '----------------------------------------------------------------
    Private Sub WriteDenHead(ByVal Den As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal e As System.Drawing.Printing.PrintPageEventArgs)
        Dim wrkNowX As Integer
        Dim wrkNowY As Integer
        Dim wrkXBase As Integer
        Dim wrkDenNo As String
        Dim wrkKessan As String
        Dim wrkFukusu As String
        Dim pf As New PCfunc
        Dim printData As String
        Dim WritePoint As Point()
        Dim WritePen As Pen

        wrkNowX = X
        wrkNowY = Y
        wrkXBase = PRINTXBASE

        If pblDenRec(Den).Head.Kessan = "1" Then
            wrkKessan = "●"
        Else
            wrkKessan = "　"
        End If

        If pblDenRec(Den).Head.FukusuChk = "1" Then
            wrkFukusu = "●"
        Else
            wrkFukusu = "　"
        End If

        'Call SetXY(wrkNowX, wrkNowY)
        printData = "日付：" & pf.AddSpace(pblDenRec(Den).Head.Year, 2) & "年 " & _
                                 pf.AddSpace(pblDenRec(Den).Head.Month, 2) & "月 " & _
                                 pf.AddSpace(pblDenRec(Den).Head.Day, 2) & "日　　" & _
                   "伝票No.：" & pf.AddSpace(pblDenRec(Den).Head.DenNo, 6) & "　　" & _
                   "決算：" & wrkKessan & "　" & "伝票結合：" & wrkFukusu
        e.Graphics.DrawString(printData, fr, Brushes.Black, wrkNowX, wrkNowY)
        'Printer.Print("日付：" & pf.AddSpace(pblDenRec(Den).Head.Year, 2) & "年 " & _
        '                         pf.AddSpace(pblDenRec(Den).Head.Month, 2) & "月 " & _
        '                         pf.AddSpace(pblDenRec(Den).Head.Day, 2) & "日　　" & _
        '           "伝票No.：" & pf.AddSpace(pblDenRec(Den).Head.DenNo, 6) & "　　" & _
        '              "決算：" & wrkKessan & "　" & _
        '          "伝票結合：" & wrkFukusu)

        'ライン
        wrkNowY = wrkNowY + 1
        WritePoint = {New Point(wrkXBase, wrkNowY), New Point(wrkXBase + 15500, wrkNowY)}
        WritePen = New Pen(Color.Black, PRINTFONTSIZE)
        e.Graphics.DrawLines(WritePen, WritePoint)

        'Call SetXY(wrkXBase, wrkNowY)
        'Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX + 15500, Printer.CurrentY)

    End Sub


    '----------------------------------------------------------------
    '   行データ出力
    '----------------------------------------------------------------
    Private Sub WriteGyouData(ByVal Den As Integer, ByVal Gyou As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal e As System.Drawing.Printing.PrintPageEventArgs)
        Dim wrkWord As String
        Dim wrkNowX As Integer
        Dim wrkNowY As Integer
        Dim wrkName As String
        Dim pf As New PCfunc

        wrkNowX = X
        wrkNowY = Y

        '借方部門
        wrkNowX = wrkNowX + 1
        'Call SetXY(wrkNowX, wrkNowY)
        wrkWord = pblDenRec(Den).Gyou(Gyou).Kari.Bumon
        If (wrkWord <> "") Then
            wrkWord = pf.AddSpace(wrkWord, 4)
            e.Graphics.DrawString(wrkWord, fr, Brushes.Black, wrkNowX, wrkNowY)
            'Printer.Print(wrkWord)
        End If

        '借方科目
        wrkNowX = wrkNowX + 6
        'Call SetXY(wrkNowX, wrkNowY)
        wrkWord = pblDenRec(Den).Gyou(Gyou).Kari.Kamoku
        If (wrkWord <> "") Then
            wrkWord = pf.AddSpace(wrkWord, 4)
            e.Graphics.DrawString(wrkWord, fr, Brushes.Black, wrkNowX, wrkNowY)
            'Printer.Print(wrkWord)
        End If

        '借方科目名
        wrkNowX = wrkNowX + 5
        'Call SetXY(wrkNowX, wrkNowY)
        wrkWord = pblDenRec(Den).Gyou(Gyou).Kari.Kamoku
        If (wrkWord <> "") Then
            wrkName = KamokuCodeToName(wrkWord)
            If Len(wrkName) > 7 Then
                wrkName = Strings.Left(wrkName, 7)
            End If
            e.Graphics.DrawString(wrkName, fr, Brushes.Black, wrkNowX, wrkNowY)
            'Printer.Print(wrkName)
        End If

        '借方補助
        wrkNowX = wrkNowX + 16
        'Call SetXY(wrkNowX, wrkNowY)
        wrkWord = pblDenRec(Den).Gyou(Gyou).Kari.Hojo
        If (wrkWord <> "") Then
            wrkWord = pf.AddSpace(wrkWord, 4)
            e.Graphics.DrawString(wrkWord, fr, Brushes.Black, wrkNowX, wrkNowY)
            'Printer.Print(wrkWord)
        End If

        '借方補助名
        wrkNowX = wrkNowX + 5
        'Call SetXY(wrkNowX, wrkNowY)
        wrkWord = pblDenRec(Den).Gyou(Gyou).Kari.Hojo
        If (wrkWord <> "") Then
            wrkName = HojoCodeToName(wrkWord, pblDenRec(Den).Gyou(Gyou).Kari.Kamoku)
            If Len(wrkName) > 7 Then
                wrkName = Strings.Left(wrkName, 7)
            End If
            e.Graphics.DrawString(wrkName, fr, Brushes.Black, wrkNowX, wrkNowY)
            'Printer.Print(wrkName)
        End If

        '借方金額
        wrkNowX = wrkNowX + 16
        'Call SetXY(wrkNowX, wrkNowY)
        wrkWord = pblDenRec(Den).Gyou(Gyou).Kari.Kin
        If (wrkWord <> "") Then
            wrkWord = pf.AddKanma(wrkWord, 3)
            wrkWord = pf.AddSpace(wrkWord, 14)
            e.Graphics.DrawString(wrkWord, fr, Brushes.Black, wrkNowX, wrkNowY)
            'Printer.Print(wrkWord)
        End If

        '借方消費税計算区分
        wrkNowX = wrkNowX + 17
        'Call SetXY(wrkNowX, wrkNowY)
        wrkWord = pblDenRec(Den).Gyou(Gyou).Kari.TaxMas
        If (wrkWord <> "") Then
            e.Graphics.DrawString(wrkWord, fr, Brushes.Black, wrkNowX, wrkNowY)
            'Printer.Print(wrkWord)
        End If

        '貸方消費税区分
        wrkNowX = wrkNowX + 2
        'Call SetXY(wrkNowX, wrkNowY)
        wrkWord = pblDenRec(Den).Gyou(Gyou).Kari.TaxKbn
        If (wrkWord <> "") Then
            wrkWord = pf.AddSpace(wrkWord, 2)
            e.Graphics.DrawString(wrkWord, fr, Brushes.Black, wrkNowX, wrkNowY)
            'Printer.Print(wrkWord)
        End If

        '貸方部門
        wrkNowX = wrkNowX + 6
        'Call SetXY(wrkNowX, wrkNowY)
        wrkWord = pblDenRec(Den).Gyou(Gyou).Kashi.Bumon
        If (wrkWord <> "") Then
            wrkWord = pf.AddSpace(wrkWord, 4)
            e.Graphics.DrawString(wrkWord, fr, Brushes.Black, wrkNowX, wrkNowY)
            'Printer.Print(wrkWord)
        End If

        '貸方科目
        wrkNowX = wrkNowX + 6
        'Call SetXY(wrkNowX, wrkNowY)
        wrkWord = pblDenRec(Den).Gyou(Gyou).Kashi.Kamoku
        wrkWord = pf.AddSpace(wrkWord, 4)
        If (wrkWord <> "") Then
            e.Graphics.DrawString(wrkWord, fr, Brushes.Black, wrkNowX, wrkNowY)
            'Printer.Print(wrkWord)
        End If

        '貸方科目名
        wrkNowX = wrkNowX + 5
        'Call SetXY(wrkNowX, wrkNowY)
        wrkWord = pblDenRec(Den).Gyou(Gyou).Kashi.Kamoku
        If (wrkWord <> "") Then
            wrkName = KamokuCodeToName(wrkWord)
            If Len(wrkName) > 7 Then
                wrkName = Strings.Left(wrkName, 7)
            End If
            e.Graphics.DrawString(wrkName, fr, Brushes.Black, wrkNowX, wrkNowY)
            'Printer.Print(wrkName)
        End If

        '貸方補助
        wrkNowX = wrkNowX + 16
        'Call SetXY(wrkNowX, wrkNowY)
        wrkWord = pblDenRec(Den).Gyou(Gyou).Kashi.Hojo
        If (wrkWord <> "") Then
            wrkWord = pf.AddSpace(wrkWord, 4)
            e.Graphics.DrawString(wrkWord, fr, Brushes.Black, wrkNowX, wrkNowY)
            'Printer.Print(wrkWord)
        End If

        '貸方補助名
        wrkNowX = wrkNowX + 5
        'Call SetXY(wrkNowX, wrkNowY)
        wrkWord = pblDenRec(Den).Gyou(Gyou).Kashi.Hojo
        If (wrkWord <> "") Then
            wrkName = HojoCodeToName(wrkWord, pblDenRec(Den).Gyou(Gyou).Kashi.Kamoku)
            If Len(wrkName) > 7 Then
                wrkName = Strings.Left(wrkName, 7)
            End If
            e.Graphics.DrawString(wrkName, fr, Brushes.Black, wrkNowX, wrkNowY)
            'Printer.Print(wrkName)
        End If

        '貸方金額
        wrkNowX = wrkNowX + 16
        'Call SetXY(wrkNowX, wrkNowY)
        wrkWord = pblDenRec(Den).Gyou(Gyou).Kashi.Kin
        If (wrkWord <> "") Then
            wrkWord = pf.AddKanma(wrkWord, 3)
            wrkWord = pf.AddSpace(wrkWord, 14)
            e.Graphics.DrawString(wrkWord, fr, Brushes.Black, wrkNowX, wrkNowY)
            'Printer.Print(wrkWord)
        End If

        '貸方消費税計算区分
        wrkNowX = wrkNowX + 17
        'Call SetXY(wrkNowX, wrkNowY)
        wrkWord = pblDenRec(Den).Gyou(Gyou).Kashi.TaxMas
        If (wrkWord <> "") Then
            e.Graphics.DrawString(wrkWord, fr, Brushes.Black, wrkNowX, wrkNowY)
            'Printer.Print(wrkWord)
        End If

        '貸方消費税区分
        wrkNowX = wrkNowX + 2
        'Call SetXY(wrkNowX, wrkNowY)
        wrkWord = pblDenRec(Den).Gyou(Gyou).Kashi.TaxKbn
        If (wrkWord <> "") Then
            wrkWord = pf.AddSpace(wrkWord, 2)
            e.Graphics.DrawString(wrkWord, fr, Brushes.Black, wrkNowX, wrkNowY)
            'Printer.Print(wrkWord)
        End If

        '複写
        wrkNowX = wrkNowX + 5
        'Call SetXY(wrkNowX, wrkNowY)
        If (pblDenRec(Den).Gyou(Gyou).CopyChk <> "0") Then
            e.Graphics.DrawString("●", fr, Brushes.Black, wrkNowX, wrkNowY)
            'Printer.Print("●")
        End If

        '摘要
        wrkNowX = wrkNowX + 3
        'Call SetXY(wrkNowX, wrkNowY)
        wrkWord = pblDenRec(Den).Gyou(Gyou).Tekiyou
        If (wrkWord <> "") Then
            e.Graphics.DrawString(wrkWord, fr, Brushes.Black, wrkNowX, wrkNowY)
            'Printer.Print(wrkWord)
        End If

    End Sub

    '----------------------------------------------------------------
    '   科目名取得
    '----------------------------------------------------------------
    Private Function KamokuCodeToName(ByVal Code As String) As String
        Dim wrkName As String
        Dim Cnt As Integer

        wrkName = ""

        '科目名取得
        For Cnt = 0 To UBound(pblKamokuData)
            'コードが一致したとき
            If (pblKamokuData(Cnt).Code = Code) Then
                '科目名取得
                wrkName = Strings.Left(pblKamokuData(Cnt).Name, 8)
                Exit For
            End If
        Next

        KamokuCodeToName = wrkName

    End Function

    '----------------------------------------------------------------
    '   補助名取得
    '----------------------------------------------------------------
    Private Function HojoCodeToName(ByVal Code As String, ByVal Kamoku As String) As String
        Dim wrkName As String
        Dim wrkExistFlg As Boolean
        Dim Cnt As Integer
        Dim Loopcnt As Integer

        wrkName = ""
        wrkExistFlg = False

        '勘定科目の記入が無ければ補助名はなし
        If Kamoku = "" Then
            HojoCodeToName = ""
            Exit Function
        End If

        '科目名取得
        For Cnt = 0 To UBound(pblKamokuData)
            'コードが一致したとき
            If (pblKamokuData(Cnt).Code = Kamoku) Then
                '補助設定が有りのとき
                If pblKamokuData(Cnt).HojoExist = True Then
                    wrkExistFlg = True
                End If
                Exit For
            End If
        Next

        '勘定科目の記入があり、補助設定がある時
        If wrkExistFlg = True Then
            '補助科目リストループ
            For Loopcnt = 0 To UBound(pblKamokuData(Cnt).HojoData)
                '補助コード一致
                If (pblKamokuData(Cnt).HojoData(Loopcnt).Code = Code) Then
                    '補助名取得
                    wrkName = pblKamokuData(Cnt).HojoData(Loopcnt).Name
                    Exit For
                End If
            Next

        End If

        HojoCodeToName = wrkName

    End Function

    '----------------------------------------------------------------
    '   行ヘッダ部出力
    '----------------------------------------------------------------
    Private Sub WriteMinSum(ByVal Kari As Decimal, ByVal Kashi As Decimal, ByVal X As Integer, ByVal Y As Integer, ByVal e As System.Drawing.Printing.PrintPageEventArgs)
        Dim wrkNowX As Integer
        Dim wrkNowY As Integer
        Dim wrkKari As String
        Dim wrkKashi As String
        Dim wrkWord As String
        Dim pf As New PCfunc

        wrkNowX = X
        wrkNowY = Y

        wrkKari = CStr(Kari)
        wrkKashi = CStr(Kashi)

        wrkNowX = wrkNowX + 4
        'Call SetXY(wrkNowX, wrkNowY)
        e.Graphics.DrawString("　小　計：", fr, Brushes.Black, wrkNowX, wrkNowY)
        'Printer.Print("　小　計：")

        wrkNowX = wrkNowX + 44
        'Call SetXY(wrkNowX, wrkNowY)
        wrkWord = pf.AddKanma(CStr(wrkKari), 3)
        wrkWord = "(" & pf.AddSpace(wrkWord, 14) & ")"
        e.Graphics.DrawString(wrkWord, fr, Brushes.Black, wrkNowX, wrkNowY)
        'Printer.Print(wrkWord)

        wrkNowX = wrkNowX + 73
        'Call SetXY(wrkNowX, wrkNowY)
        wrkWord = pf.AddKanma(CStr(wrkKashi), 3)
        wrkWord = "(" & pf.AddSpace(wrkWord, 14) & ")"
        e.Graphics.DrawString(wrkWord, fr, Brushes.Black, wrkNowX, wrkNowY)
        'Printer.Print(wrkWord)

    End Sub

    '----------------------------------------------------------------
    '   X座標Y座標設定
    '----------------------------------------------------------------
    'Private Sub SetXY(ByRef X As Integer, ByRef Y As Integer)
    'Printer.CurrentX = Printer.TextWidth("-") * X
    'Printer.CurrentY = Printer.TextHeight("-") * Y
    'End Sub


    Private Sub PrintDImage_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDImage.PrintPage
        '----------------------------------------------------------------
        '   伝票印刷
        '----------------------------------------------------------------
        'Public Sub PrintImage(ByVal FileName As String, ByVal OutMode As Integer)
        ' OutMode 0:仕訳伝票 1:入出金
        Dim MyFile As String
        Dim wrkXBase As Integer
        Dim wrkYBase As Integer
        Dim wrkNowX As Integer
        Dim wrkNowY As Integer
        Dim WritePen As New Pen(Color.Black, PRINTFONTSIZE)
        Dim WritePoint As Point()
        Dim pf As New PCfunc

        On Error GoTo ErrPrc

        '印刷設定
        e.PageSettings.PaperSize.PaperName = Printing.PaperKind.A4
        e.PageSettings.Landscape = False

        'Printer.Orientation = vbPRORPortrait
        'Printer.FontName = "MS 明朝"
        'Printer.FontSize = PRINTFONTSIZE
        wrkXBase = PRINTXBASE
        wrkYBase = PRINTYBASE
        wrkNowX = wrkXBase
        wrkNowY = wrkYBase
        If pblPrintMode = 0 Then
            'Printer.PaperSize = vbPRPSA4
            'Printer.Orientation = vbPRORPortrait
            e.PageSettings.PaperSize.PaperName = Printing.PaperKind.A4
            e.PageSettings.Landscape = False
        Else
            e.PageSettings.PaperSize.PaperName = Printing.PaperKind.A5
            e.PageSettings.Landscape = True
            'Printer.PaperSize = vbPRPSA5
            'Printer.Orientation = vbPRORLandscape
        End If

        'タイトル
        wrkNowX = wrkXBase + 42
        'Call SetXY(wrkNowX, wrkNowY)
        'Printer.Print(App.Title)
        e.Graphics.DrawString(Title, fr, Brushes.Black, wrkNowX, wrkNowY)

        '日付
        wrkNowX = wrkXBase + 75
        wrkNowY = wrkNowY + 1
        'Call SetXY(wrkNowX, wrkNowY)
        'Printer.Print Date & " " & Time
        e.Graphics.DrawString(System.DateTime.Today.ToString & " " & Microsoft.VisualBasic.DateAndTime.TimeString, fr, Brushes.Black, wrkNowX, wrkNowY)

        'ライン
        wrkNowY = wrkNowY + 2
        'Call SetXY(wrkXBase, wrkNowY)
        'Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX + 10500, Printer.CurrentY)
        WritePoint = {New Point(wrkXBase, wrkNowY), New Point(wrkXBase + 10500, wrkNowY)}
        e.Graphics.DrawLines(WritePen, WritePoint)


        '画像表示
        If pblPrintFile <> "" Then
            wrkNowY = wrkNowY + 2
            wrkNowX = wrkXBase
            'Call SetXY(wrkNowX, wrkNowY)
            If pblPrintMode = 0 Then
                e.Graphics.DrawImageUnscaledAndClipped(New Bitmap(pblPrintFile), New Rectangle(600, 1500, 10300, 1400))
                'Printer.PaintPicture(LoadPicture(FileName), 600, 1500, 10300, 14000)
            Else
                e.Graphics.DrawImageUnscaledAndClipped(New Bitmap(pblPrintFile), New Rectangle(400, 700, 9300, 7300))
                'Printer.PaintPicture(LoadPicture(FileName), 400, 700, 9300, 7300)
            End If
        End If

        'Printer.EndDoc()
        e.HasMorePages = False

        Exit Sub

        ' エラー処理
ErrPrc:
        MessageBox.Show("印刷中に不具合が発生したため印刷を中断します", Title)
        Exit Sub
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdImgPrn.Click

    End Sub
End Class