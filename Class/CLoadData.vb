Imports 仕訳伝票.PCData
Imports 仕訳伝票.CErr
Imports 仕訳伝票.frmProg

Public Class CLoadData
    '----------------------------------------------------------------
    '   伝票データロード
    '----------------------------------------------------------------
    Public Sub LoadData()
        On Error GoTo ErrPrc

        'frmProg.Caption = "データロード中・・・伝票"
        'frmProg.prgBar.Value = 15

        Call LoadDataFurikae()

        Exit Sub
        On Error GoTo 0
        'エラー処理
ErrPrc:
        Dim em As CErr
        em.ErrMessage("伝票データ取得中")
    End Sub

    '----------------------------------------------------------------
    '   振替伝票データロード
    '----------------------------------------------------------------
    Private Sub LoadDataFurikae()
        'Dim inTmp As Integer          '入力ファイル番号
        Dim readbuf As String           '1行読込みバッファ
        Dim Gyou As Integer          '行カウンタ
        Dim wrkHead As New strInHead        'Tempヘッダデータ
        Dim wrkGyou As New strInGyou        'Temp行データ
        Dim ret As String
        Dim firstFlg As String
        Dim wrkDenRec As strInputRecord
        Dim ix As Integer
        Dim iz As Integer
        Dim sFileName As String
        Dim iGyoCnt As Integer
        Dim inFileR As System.IO.StreamReader

        Dim fp As New frmProg
        fp.Show()

        '------------------------------------------------
        ' ファイルオープン
        '------------------------------------------------
        On Error GoTo ErrPrc

        pblDenNum = 0
        firstFlg = FLGON

        '伝票合計メモリクリア
        Call InitDenRec_Total(wrkDenRec)

        '分割後のＣＳＶファイルを読み込む
        If Dir(pblInstPath & DIR_INCSV & "*.csv") <> "" Then
            sFileName = Dir(pblInstPath & DIR_INCSV & "*.csv")

            Do While sFileName <> ""
                'CSVファイルを開く
                inFileR = My.Computer.FileSystem.OpenTextFileReader(pblInstPath & DIR_INCSV & sFileName)

                '------------------------------------------
                ' 全データロード
                '------------------------------------------
                Do While Not EOF(inFileR.EndOfStream <> True)
                    System.Windows.Forms.Application.DoEvents()

                    ' 1行リード
                    readbuf = inFileR.ReadLine

                    '先頭に「*」か「#」があったら新たな伝票なのでヘッダ格納
                    If Strings.Left(readbuf, 1) = "*" Then
                        firstFlg = FLGOFF

                        'メモリクリア処理
                        Call InitDenRec(wrkDenRec)

                        'ヘッダ取得
                        Call DataGetHead(wrkDenRec.Head, readbuf, sFileName)
                    Else
                        '行データ格納
                        Call DataGetGyou(wrkGyou, readbuf)
                        '空白行は無視
                        If ((Val(wrkGyou.GyouNum) >= 1) And (Val(wrkGyou.GyouNum) <= MAXGYOU)) Then
                            wrkDenRec.Gyou(Val(wrkGyou.GyouNum)) = wrkGyou
                        End If
                    End If
                Loop

                '伝票データあり
                If (firstFlg = FLGOFF) Then
                    '伝票枚数＋１
                    pblDenNum = pblDenNum + 1
                    '配列数再宣言
                    ReDim Preserve pblDenRec(0 To pblDenNum)
                    'メモリクリア
                    Call InitDenRec(pblDenRec(pblDenNum))
                    'データ格納
                    Call GetDenRec(wrkDenRec, pblDenRec(pblDenNum))

                End If

                On Error GoTo 0
                '------------------------------------------------------
                ' ファイルクローズ
                '------------------------------------------------------
                inFileR.Close()

                sFileName = Dir()
            Loop

        End If

        '伝票なし
        If (pblDenNum = 0) Then
            MessageBox.Show("不正な伝票ファイルです。伝票データが存在しません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End
        End If

        If Trim(pblDenRec(1).Head.Image) = "" Then
            MessageBox.Show("WinReaderHand S から画像が出力されていないため" & vbNewLine & "画像の表示はされません", Title, MessageBoxButtons.OK)
        End If

        fp.Close()
        Exit Sub

ErrPrc:
        'エラー処理
        inFileR.Close()
        fp.Close()
        Dim em As New CErr
        em.ErrMessage("伝票データ取得中")

    End Sub

    '----------------------------------------------------------------
    '   伝票データ初期化
    '----------------------------------------------------------------
    Private Sub InitDenRec(ByRef DenRec As strInputRecord)
        Dim Cnt As Integer

        With DenRec.Head
            .Image = ""
            .CsvFile = ""
            .Year = ""
            .Month = ""
            .Day = ""
            .Kessan = ""
            .FukusuChk = ""
            .DenNo = ""
        End With

        ReDim Preserve DenRec.Gyou(0 To MAXGYOU)
        For Cnt = 1 To MAXGYOU
            With DenRec.Gyou(Cnt)
                .GyouNum = ""
                .Kari.Bumon = ""
                .Kari.Kamoku = ""
                .Kari.Hojo = ""
                .Kari.Kin = ""
                .Kari.TaxMas = ""
                .Kari.TaxKbn = ""
                .Kashi.Bumon = ""
                .Kashi.Kamoku = ""
                .Kashi.Hojo = ""
                .Kashi.Kin = ""
                .Kashi.TaxMas = ""
                .Kashi.TaxKbn = ""
                .CopyChk = ""
                .Tekiyou = ""
            End With
        Next

        With DenRec
            .KariTotal = 0
            .KashiTotal = 0
        End With

    End Sub

    '----------------------------------------------------------------
    '   伝票データ初期化
    '----------------------------------------------------------------
    Private Sub InitDenRec_Total(ByVal DenRec As strInputRecord)

        With DenRec.Head
            .Kari_T = 0
            .Kashi_T = 0
            .FukuMai = 0
        End With

    End Sub

    '----------------------------------------------------------------
    '   伝票データ格納
    '----------------------------------------------------------------
    Private Sub GetDenRec(ByVal srcDenRec As strInputRecord, ByVal dstDenRec As strInputRecord)
        Dim Cnt As Integer

        'ヘッダはそのまま格納
        dstDenRec.Head = srcDenRec.Head

        '全行ループ
        For Cnt = 1 To MAXGYOU
            '行もそのまま格納
            If srcDenRec.Gyou(Cnt).GyouNum = "" Then
                srcDenRec.Gyou(Cnt).CopyChk = "0"
                srcDenRec.Gyou(Cnt).Torikeshi = "1"
            End If
            dstDenRec.Gyou(Cnt) = srcDenRec.Gyou(Cnt)
        Next

        '合計金額
        dstDenRec.KariTotal = srcDenRec.KariTotal
        dstDenRec.KashiTotal = srcDenRec.KashiTotal

        dstDenRec.Head.Kari_T = srcDenRec.Head.Kari_T
        dstDenRec.Head.Kashi_T = srcDenRec.Head.Kashi_T

    End Sub

    '----------------------------------------------------------------
    '   伝票ヘッダ部格納
    '----------------------------------------------------------------
    Private Sub DataGetHead(ByVal Head As strInHead, ByVal readbuf As String, ByVal csvf As String)
        Dim Cnt As Integer          '文字列展開用カウンタ
        Dim hcnt As Integer         '文字列展開用カウンタ

        'アスタリスクは読み捨て
        Cnt = 1
        hcnt = InStr(Cnt, readbuf, ",", vbBinaryCompare)

        '画像ファイル名
        Cnt = hcnt + 1
        hcnt = InStr(Cnt, readbuf, ",", vbBinaryCompare)
        Head.Image = Trim(Strings.Left(csvf, csvf.Length - 4) & ".bmp")

        'CSVファイル名
        Head.CsvFile = csvf
        '年
        Cnt = hcnt + 1
        hcnt = InStr(Cnt, readbuf, ",", vbBinaryCompare)
        Head.Year = StrConv(Mid(readbuf, Cnt, hcnt - Cnt), vbNarrow)
        '--データ変換処理--Yamamoto
        Head.Year = Trim(Replace(Head.Year, "-", ""))
        '月
        Cnt = hcnt + 1
        hcnt = InStr(Cnt, readbuf, ",", vbBinaryCompare)
        Head.Month = StrConv(Mid(readbuf, Cnt, hcnt - Cnt), vbNarrow)
        '--データ変換処理--Yamamoto
        Head.Month = Trim(Replace(Head.Month, "-", ""))
        '日
        Cnt = hcnt + 1
        hcnt = InStr(Cnt, readbuf, ",", vbBinaryCompare)
        Head.Day = StrConv(Mid(readbuf, Cnt, hcnt - Cnt), vbNarrow)
        '--データ変換処理--Yamamoto
        Head.Day = Trim(Replace(Head.Day, "-", ""))
        '伝票No.
        Cnt = hcnt + 1
        hcnt = InStr(Cnt, readbuf, ",", vbBinaryCompare)
        Head.DenNo = StrConv(Mid(readbuf, Cnt, hcnt - Cnt), vbNarrow)
        '--データ変換処理--Yamamoto
        Head.DenNo = Trim(Replace(Head.DenNo, "-", ""))
        '決算処理グラフ
        Cnt = hcnt + 1
        hcnt = InStr(Cnt, readbuf, ",", vbBinaryCompare)
        Head.Kessan = Trim(StrConv(Mid(readbuf, Cnt, hcnt - Cnt), vbNarrow))
        '複数枚チェック
        Cnt = hcnt + 1
        Head.FukusuChk = Trim(StrConv(Mid(readbuf, Cnt), vbNarrow))

    End Sub

    '----------------------------------------------------------------
    '   伝票行データ格納
    '----------------------------------------------------------------
    Private Sub DataGetGyou(ByVal GyouData As strInGyou, ByVal readbuf As String)

        Dim Cnt As Integer          '文字列展開用カウンタ
        Dim hcnt As Integer         '文字列展開用カウンタ
        Dim wrkGyou As strInGyou
        Dim GyoCnt As Integer

        Cnt = 1
        hcnt = InStr(Cnt, readbuf, ",", vbBinaryCompare)

        '*** 2006.08.29 追加 ***
        wrkGyou.Torikeshi = Trim(Mid(readbuf, Cnt, hcnt - Cnt))

        Cnt = hcnt + 1
        hcnt = InStr(Cnt, readbuf, ",", vbBinaryCompare)
        '*** 2006.08.29 変更 ***
        wrkGyou.GyouNum = Mid(readbuf, Cnt, hcnt - Cnt)

        '借方明細
        Cnt = hcnt + 1
        hcnt = InStr(Cnt, readbuf, ",", vbBinaryCompare)
        wrkGyou.Kari.Bumon = StrConv(Mid(readbuf, Cnt, hcnt - Cnt), vbNarrow)
        '--データ変換処理--Yamamoto
        If Trim(wrkGyou.Kari.Bumon) = "" Then
            wrkGyou.Kari.Bumon = ""
        ElseIf Not InStr(1, wrkGyou.Kari.Bumon, "-", vbBinaryCompare) = 0 Then
            wrkGyou.Kari.Bumon = Trim(Replace(wrkGyou.Kari.Bumon, " ", ""))
            wrkGyou.Kari.Bumon = Trim(Replace(wrkGyou.Kari.Bumon, "-", ""))
            If IsNumeric(wrkGyou.Kari.Bumon) = True Then
                wrkGyou.Kari.Bumon = Format(wrkGyou.Kari.Bumon, "-##0")
            Else
                wrkGyou.Kari.Bumon = Trim("-" & wrkGyou.Kari.Bumon)
            End If
        Else
            wrkGyou.Kari.Bumon = Trim(Replace(wrkGyou.Kari.Bumon, " ", ""))
            If IsNumeric(wrkGyou.Kari.Bumon) = True Then
                wrkGyou.Kari.Bumon = Format(wrkGyou.Kari.Bumon, "###0")
            End If
        End If
        Cnt = hcnt + 1
        hcnt = InStr(Cnt, readbuf, ",", vbBinaryCompare)
        wrkGyou.Kari.Kamoku = StrConv(Mid(readbuf, Cnt, hcnt - Cnt), vbNarrow)
        '--データ変換処理--Yamamoto
        If Trim(wrkGyou.Kari.Kamoku) = "" Then
            wrkGyou.Kari.Kamoku = ""
        ElseIf Not InStr(1, wrkGyou.Kari.Kamoku, "-", vbBinaryCompare) = 0 Then
            wrkGyou.Kari.Kamoku = Trim(Replace(wrkGyou.Kari.Kamoku, " ", ""))
            wrkGyou.Kari.Kamoku = Trim(Replace(wrkGyou.Kari.Kamoku, "-", ""))
            If IsNumeric(wrkGyou.Kari.Kamoku) = True Then
                wrkGyou.Kari.Kamoku = Format(wrkGyou.Kari.Kamoku, "-##0")
            Else
                wrkGyou.Kari.Kamoku = Trim("-" & wrkGyou.Kari.Kamoku)
            End If
        Else
            wrkGyou.Kari.Kamoku = Trim(Replace(wrkGyou.Kari.Kamoku, " ", ""))
            If IsNumeric(wrkGyou.Kari.Kamoku) = True Then
                wrkGyou.Kari.Kamoku = Format(wrkGyou.Kari.Kamoku, "###0")
            End If
        End If

        '*** 2006.08.29 追加 ***
        '科目が設定されている場合、基本情報で「部門あり」時は、部門が設定されていない場合は、０を設定
        If (wrkGyou.Kari.Kamoku <> "") And (pblBumonFlg = True) And (wrkGyou.Kari.Bumon = "") Then
            wrkGyou.Kari.Bumon = "0"
        End If
        '*** 2006.08.29 追加 ***

        Cnt = hcnt + 1
        hcnt = InStr(Cnt, readbuf, ",", vbBinaryCompare)
        wrkGyou.Kari.Hojo = StrConv(Mid(readbuf, Cnt, hcnt - Cnt), vbNarrow)
        '--データ変換処理--Yamamoto
        If Trim(wrkGyou.Kari.Hojo) = "" Then
            wrkGyou.Kari.Hojo = ""
        ElseIf Not InStr(1, wrkGyou.Kari.Hojo, "-", vbBinaryCompare) = 0 Then
            wrkGyou.Kari.Hojo = Trim(Replace(wrkGyou.Kari.Hojo, " ", ""))
            wrkGyou.Kari.Hojo = Trim(Replace(wrkGyou.Kari.Hojo, "-", ""))
            If IsNumeric(wrkGyou.Kari.Hojo) = True Then
                wrkGyou.Kari.Hojo = Format(wrkGyou.Kari.Hojo, "-##0")
            Else
                wrkGyou.Kari.Hojo = Trim("-" & wrkGyou.Kari.Hojo)
            End If
        Else
            wrkGyou.Kari.Hojo = Trim(Replace(wrkGyou.Kari.Hojo, " ", ""))
            If IsNumeric(wrkGyou.Kari.Hojo) = True Then
                wrkGyou.Kari.Hojo = Format(wrkGyou.Kari.Hojo, "###0")
            End If
        End If
        Cnt = hcnt + 1
        hcnt = InStr(Cnt, readbuf, ",", vbBinaryCompare)
        wrkGyou.Kari.Kin = StrConv(Mid(readbuf, Cnt, hcnt - Cnt), vbNarrow)

        '--データ変換処理--Yamamoto
        If Trim(wrkGyou.Kari.Kin) = "" Then
            wrkGyou.Kari.Kin = ""
        ElseIf Not InStr(1, wrkGyou.Kari.Kin, "-", vbBinaryCompare) = 0 Then
            wrkGyou.Kari.Kin = Trim(Replace(wrkGyou.Kari.Kin, " ", ""))
            wrkGyou.Kari.Kin = Trim(Replace(wrkGyou.Kari.Kin, "-", ""))
            If IsNumeric(wrkGyou.Kari.Kin) = True Then
                wrkGyou.Kari.Kin = Format(wrkGyou.Kari.Kin, "-##########")
            Else
                wrkGyou.Kari.Kin = "-9"
            End If
        Else
            wrkGyou.Kari.Kin = Trim(Replace(wrkGyou.Kari.Kin, " ", ""))
            If IsNumeric(wrkGyou.Kari.Kin) = True Then
                wrkGyou.Kari.Kin = Format(wrkGyou.Kari.Kin, "###########")
            Else
                wrkGyou.Kari.Kin = "-9"
            End If
        End If

        Cnt = hcnt + 1
        hcnt = InStr(Cnt, readbuf, ",", vbBinaryCompare)
        wrkGyou.Kari.TaxMas = StrConv(Mid(readbuf, Cnt, hcnt - Cnt), vbNarrow)
        Cnt = hcnt + 1
        hcnt = InStr(Cnt, readbuf, ",", vbBinaryCompare)
        wrkGyou.Kari.TaxKbn = StrConv(Mid(readbuf, Cnt, hcnt - Cnt), vbNarrow)
        '--データ変換処理--Yamamoto
        If Trim(wrkGyou.Kari.TaxKbn) = "" Then
            wrkGyou.Kari.TaxKbn = ""
        ElseIf Not InStr(1, wrkGyou.Kari.TaxKbn, "-", vbBinaryCompare) = 0 Then
            wrkGyou.Kari.TaxKbn = Trim(Replace(wrkGyou.Kari.TaxKbn, " ", ""))
            wrkGyou.Kari.TaxKbn = Trim(Replace(wrkGyou.Kari.TaxKbn, "-", ""))
            If IsNumeric(wrkGyou.Kari.TaxKbn) = True Then
                wrkGyou.Kari.TaxKbn = Format(wrkGyou.Kari.TaxKbn, "-0")
            Else
                wrkGyou.Kari.TaxKbn = Trim("-" & wrkGyou.Kari.TaxKbn)
            End If
        Else
            wrkGyou.Kari.TaxKbn = Trim(Replace(wrkGyou.Kari.TaxKbn, " ", ""))
            If IsNumeric(wrkGyou.Kari.TaxKbn) = True Then
                wrkGyou.Kari.TaxKbn = Format(wrkGyou.Kari.TaxKbn, "#0")
            End If
        End If

        '貸方明細
        Cnt = hcnt + 1
        hcnt = InStr(Cnt, readbuf, ",", vbBinaryCompare)
        wrkGyou.Kashi.Bumon = StrConv(Mid(readbuf, Cnt, hcnt - Cnt), vbNarrow)
        '--データ変換処理--Yamamoto
        If Trim(wrkGyou.Kashi.Bumon) = "" Then
            wrkGyou.Kashi.Bumon = ""
        ElseIf Not InStr(1, wrkGyou.Kashi.Bumon, "-", vbBinaryCompare) = 0 Then
            wrkGyou.Kashi.Bumon = Trim(Replace(wrkGyou.Kashi.Bumon, " ", ""))
            wrkGyou.Kashi.Bumon = Trim(Replace(wrkGyou.Kashi.Bumon, "-", ""))
            If IsNumeric(wrkGyou.Kashi.Bumon) = True Then
                wrkGyou.Kashi.Bumon = Format(wrkGyou.Kashi.Bumon, "-##0")
            Else
                wrkGyou.Kashi.Bumon = Trim("-" & wrkGyou.Kashi.Bumon)
            End If
        Else
            wrkGyou.Kashi.Bumon = Trim(Replace(wrkGyou.Kashi.Bumon, " ", ""))
            If IsNumeric(wrkGyou.Kashi.Bumon) = True Then
                wrkGyou.Kashi.Bumon = Format(wrkGyou.Kashi.Bumon, "###0")
            End If
        End If
        Cnt = hcnt + 1
        hcnt = InStr(Cnt, readbuf, ",", vbBinaryCompare)
        wrkGyou.Kashi.Kamoku = StrConv(Mid(readbuf, Cnt, hcnt - Cnt), vbNarrow)
        '--データ変換処理--Yamamoto
        If Trim(wrkGyou.Kashi.Kamoku) = "" Then
            wrkGyou.Kashi.Kamoku = ""
        ElseIf Not InStr(1, wrkGyou.Kashi.Kamoku, "-", vbBinaryCompare) = 0 Then
            wrkGyou.Kashi.Kamoku = Trim(Replace(wrkGyou.Kashi.Kamoku, " ", ""))
            wrkGyou.Kashi.Kamoku = Trim(Replace(wrkGyou.Kashi.Kamoku, "-", ""))
            If IsNumeric(wrkGyou.Kashi.Kamoku) = True Then
                wrkGyou.Kashi.Kamoku = Format(wrkGyou.Kashi.Kamoku, "-##0")
            Else
                wrkGyou.Kashi.Kamoku = Trim("-" & wrkGyou.Kashi.Kamoku)
            End If
        Else
            wrkGyou.Kashi.Kamoku = Trim(Replace(wrkGyou.Kashi.Kamoku, " ", ""))
            If IsNumeric(wrkGyou.Kashi.Kamoku) = True Then
                wrkGyou.Kashi.Kamoku = Format(wrkGyou.Kashi.Kamoku, "###0")
            End If
        End If

        '科目が設定されている場合、基本情報で「部門あり」時は、部門が設定されていない場合は、０を設定
        If (wrkGyou.Kashi.Kamoku <> "") And (pblBumonFlg = True) And (wrkGyou.Kashi.Bumon = "") Then
            wrkGyou.Kashi.Bumon = "0"
        End If

        Cnt = hcnt + 1
        hcnt = InStr(Cnt, readbuf, ",", vbBinaryCompare)
        wrkGyou.Kashi.Hojo = StrConv(Mid(readbuf, Cnt, hcnt - Cnt), vbNarrow)
        '--データ変換処理--
        If Trim(wrkGyou.Kashi.Hojo) = "" Then
            wrkGyou.Kashi.Hojo = ""
        ElseIf Not InStr(1, wrkGyou.Kashi.Hojo, "-", vbBinaryCompare) = 0 Then
            wrkGyou.Kashi.Hojo = Trim(Replace(wrkGyou.Kashi.Hojo, " ", ""))
            wrkGyou.Kashi.Hojo = Trim(Replace(wrkGyou.Kashi.Hojo, "-", ""))
            If IsNumeric(wrkGyou.Kashi.Hojo) = True Then
                wrkGyou.Kashi.Hojo = Format(wrkGyou.Kashi.Hojo, "-##0")
            Else
                wrkGyou.Kashi.Hojo = Trim("-" & wrkGyou.Kashi.Hojo)
            End If
        Else
            wrkGyou.Kashi.Hojo = Trim(Replace(wrkGyou.Kashi.Hojo, " ", ""))
            If IsNumeric(wrkGyou.Kashi.Hojo) = True Then
                wrkGyou.Kashi.Hojo = Format(wrkGyou.Kashi.Hojo, "###0")
            End If
        End If
        Cnt = hcnt + 1
        hcnt = InStr(Cnt, readbuf, ",", vbBinaryCompare)
        wrkGyou.Kashi.Kin = StrConv(Mid(readbuf, Cnt, hcnt - Cnt), vbNarrow)

        '--データ変換処理--
        If Trim(wrkGyou.Kashi.Kin) = "" Then
            wrkGyou.Kashi.Kin = ""
        ElseIf Not InStr(1, wrkGyou.Kashi.Kin, "-", vbBinaryCompare) = 0 Then
            wrkGyou.Kashi.Kin = Trim(Replace(wrkGyou.Kashi.Kin, " ", ""))
            wrkGyou.Kashi.Kin = Trim(Replace(wrkGyou.Kashi.Kin, "-", ""))
            If IsNumeric(wrkGyou.Kashi.Kin) = True Then
                wrkGyou.Kashi.Kin = Format(wrkGyou.Kashi.Kin, "-##########")
            Else
                wrkGyou.Kashi.Kin = "-9"
            End If
        Else
            wrkGyou.Kashi.Kin = Trim(Replace(wrkGyou.Kashi.Kin, " ", ""))
            If IsNumeric(wrkGyou.Kashi.Kin) = True Then
                wrkGyou.Kashi.Kin = Format(wrkGyou.Kashi.Kin, "###########")
            Else
                wrkGyou.Kashi.Kin = "-9"
            End If
        End If

        Cnt = hcnt + 1
        hcnt = InStr(Cnt, readbuf, ",", vbBinaryCompare)
        wrkGyou.Kashi.TaxMas = StrConv(Mid(readbuf, Cnt, hcnt - Cnt), vbNarrow)
        Cnt = hcnt + 1
        hcnt = InStr(Cnt, readbuf, ",", vbBinaryCompare)
        wrkGyou.Kashi.TaxKbn = StrConv(Mid(readbuf, Cnt, hcnt - Cnt), vbNarrow)
        '--データ変換処理--
        If Trim(wrkGyou.Kashi.TaxKbn) = "" Then
            wrkGyou.Kashi.TaxKbn = ""
        ElseIf Not InStr(1, wrkGyou.Kashi.TaxKbn, "-", vbBinaryCompare) = 0 Then
            wrkGyou.Kashi.TaxKbn = Trim(Replace(wrkGyou.Kashi.TaxKbn, " ", ""))
            wrkGyou.Kashi.TaxKbn = Trim(Replace(wrkGyou.Kashi.TaxKbn, "-", ""))
            If IsNumeric(wrkGyou.Kashi.TaxKbn) = True Then
                wrkGyou.Kashi.TaxKbn = Format(wrkGyou.Kashi.TaxKbn, "-0")
            Else
                wrkGyou.Kashi.TaxKbn = Trim("-" & wrkGyou.Kashi.TaxKbn)
            End If
        Else
            wrkGyou.Kashi.TaxKbn = Trim(Replace(wrkGyou.Kashi.TaxKbn, " ", ""))
            If IsNumeric(wrkGyou.Kashi.TaxKbn) = True Then
                wrkGyou.Kashi.TaxKbn = Format(wrkGyou.Kashi.TaxKbn, "#0")
            End If
        End If

        '摘要複写
        ' １行目の摘要複写は存在しない
        If GyoCnt <> 1 Then
            Cnt = hcnt + 1
            hcnt = InStr(Cnt, readbuf, ",", vbBinaryCompare)
            wrkGyou.CopyChk = Trim(StrConv(Mid(readbuf, Cnt, hcnt - Cnt), vbNarrow))
        Else
            wrkGyou.CopyChk = "0"
        End If

        '摘要
        Cnt = hcnt + 1
        wrkGyou.Tekiyou = RTrim(StrConv(Mid(readbuf, Cnt), vbWide))
        GyouData = wrkGyou
        Exit Sub

    End Sub

End Class
