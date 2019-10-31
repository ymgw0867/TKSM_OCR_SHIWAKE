Imports 仕訳伝票.PCData
Imports 仕訳伝票.CErr


Public Class CLoadCSV
    '----------------------------------------------------------------
    '   伝票データロード
    '----------------------------------------------------------------
    Public Sub LoadCSV()
        On Error GoTo ErrPrc

        '入力ファイルがなければ終わる
        If pblFlgINFILE = False Then
            Exit Sub
        End If

        'frmProg.Caption = "データロード中・・・伝票"
        'frmProg.prgBar.Value = 15

        Call LoadCsvDivide()

        Exit Sub
        On Error GoTo 0
        'エラー処理
ErrPrc:
        Dim em As New CErr
        em.ErrMessage("伝票データ分割中")

    End Sub

    '----------------------------------------------------------------
    '   伝票ＣＳＶデータを一枚ごとに分割する
    '----------------------------------------------------------------
    Private Sub LoadCsvDivide()
        Dim sInbuf As String           '1行読込みバッファ
        Dim firstFlg As String
        Dim Cnt As Integer
        Dim hcnt As Integer
        Dim ImgNAME As String           '画像ファイル名

        Dim fp As New frmProg
        fp.Show()

        '--------------------------------------------------
        ' ファイルオープン
        '--------------------------------------------------
        '入力ファイル
        Dim inFileR = My.Computer.FileSystem.OpenTextFileReader(pblInstPath & DIR_HENKAN & TMPREAD)
        'Dim oFileR As System.IO.StreamWriter
        Dim oFileR = My.Computer.FileSystem.OpenTextFileWriter(pblInstPath & DIR_HENKAN & DIVFILE, False)

        On Error GoTo ErrPrc
        firstFlg = FLGON
        pblDenNum = 0

        Do While (inFileR.EndOfStream <> True)
            '----------------------------------------------------
            ' 全データロード
            '----------------------------------------------------
            sInbuf = inFileR.ReadLine()
            '先頭に「*」か「#」があったら新たな伝票なのでCSVファイル作成
            If Left(sInbuf, 1) = "*" Then
                '最初の伝票以外
                If (firstFlg = FLGOFF) Then
                    'ファイルクローズ
                    oFileR.Close()
                    oFileR.Dispose()
                    oFileR = Nothing

                    '伝票ＣＳＶファイルコピー
                    Call CSV_FileCopy(ImgNAME)

                    '出力ファイルオープン
                    oFileR = My.Computer.FileSystem.OpenTextFileWriter(pblInstPath & DIR_HENKAN & DIVFILE, False)

                End If

                pblDenNum = pblDenNum + 1
                firstFlg = FLGOFF

                'アスタリスクは読み捨て
                Cnt = 1
                hcnt = InStr(Cnt, sInbuf, ",", vbBinaryCompare)

                '画像ファイル名を取得
                Cnt = hcnt + 1
                hcnt = InStr(Cnt, sInbuf, ".", vbBinaryCompare)
                ImgNAME = StrConv(Mid(sInbuf, Cnt, hcnt - Cnt), vbNarrow)

            End If
            oFileR.WriteLine(sInbuf)
        Loop

        '伝票データあり
        If (firstFlg = FLGOFF) Then
            oFileR.Close()
        End If

        On Error GoTo 0

        '------------------
        ' ファイルクローズ
        '------------------
        inFileR.Close()

        '伝票なし
        Dim pf As New PCfunc
        If (pblDenNum = 0) Then
            MessageBox.Show("不正な伝票ファイルです。伝票データが存在しません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

            '分割一時ファイルの削除
            pf.FileDelete(pblInstPath & DIR_HENKAN & DIVFILE)

            End
        Else
            '伝票ＣＳＶファイルコピー
            Call CSV_FileCopy(ImgNAME)

            '分割一時ファイルの削除
            pf.FileDelete(pblInstPath & DIR_HENKAN & DIVFILE)

        End If

        '-------------------------
        '   入力ファイルを削除する　
        '-------------------------
        '一時ファイル削除
        pf.FileDelete(pblInstPath & DIR_HENKAN & TMPREAD)

        '入力ファイル削除
        pf.FileDelete(pblInstPath & DIR_HENKAN & INFILE)

        '画像ファイル削除 "FileDelete"に変更
        pf.FileDelete(pblInstPath & DIR_HENKAN & "WRH*.bmp")

        fp.Close()
        Exit Sub

ErrPrc:
        'エラー処理
        fp.Close()
        Dim em As New CErr
        em.ErrMessage("伝票データ分割中")

    End Sub

    Sub CSV_FileCopy(ByVal ImageName)
        Dim sfile As String

        '-------------------------------------------------
        '   分割したＣＳＶファイルのコピー
        '-------------------------------------------------
        sfile = System.DateTime.Now
        sfile = Format(sfile, "yyMMddHHmmss")
        sfile = Strings.Right(sfile, 10)

        FileCopy(pblInstPath & DIR_HENKAN & DIVFILE, _
                 pblInstPath & DIR_INCSV & _
                 sfile & CStr(Format(pblDenNum, "000")) & ImageName & ".csv")
        '-------------------------------------------------
        '   画像ファイルのコピー
        '-------------------------------------------------

        FileCopy(pblInstPath & DIR_HENKAN & ImageName & ".bmp", _
                 pblInstPath & DIR_INCSV & _
                 sfile & CStr(Format(pblDenNum, "000")) & ImageName & ".bmp")

    End Sub

    Public Sub LoadChudan()
        '-------------------------------------------------
        '   中断データリカバリー
        '-------------------------------------------------
        Dim sFileName As String
        Dim gFileName As String
        Dim fFileName As String
        Dim i As Integer

        '分割フォルダのデータを削除
        If (Dir(pblInstPath & DIR_INCSV & "*.*") <> "") Then
            Kill(pblInstPath & DIR_INCSV & "*.*")
        End If

        For i = 1 To UBound(pblRecov)
            If pblRecov(i).recFlg = 1 Then
                If Dir(pblInstPath & DIR_BREAK & Format(pblComNo, "000") & "\*.*") <> "" Then
                    sFileName = Dir(pblInstPath & DIR_BREAK & Format(pblComNo, "000") & "\*.*")
                    Do While sFileName <> ""
                        '日付時間を取得
                        gFileName = Left(sFileName, 10)
                        'ファイル名を取得
                        fFileName = Right(sFileName, 25)
                        'ファイル名
                        sFileName = pblInstPath & DIR_BREAK & CStr(Format(pblComNo, "000")) & "\" _
                                  & sFileName
                        If pblRecov(i).recName = gFileName Then
                            Microsoft.VisualBasic.FileSystem.Rename(sFileName, pblInstPath & DIR_INCSV & fFileName)
                        End If
                        '次のファイル
                        sFileName = Dir()
                    Loop
                End If
            End If
        Next i

        '全てのファイルをリカバリーしたときはフォルダを削除
        If Dir(pblInstPath & DIR_BREAK & CStr(Format(pblComNo, "000")) & "\*.*") = "" Then
            RmDir(pblInstPath & DIR_BREAK & CStr(Format(pblComNo, "000")))
        End If
    End Sub

End Class
