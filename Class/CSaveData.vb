Imports fp = 仕訳伝票.frmProg
Imports 仕訳伝票.PCData

Public Class CSaveData
    '----------------------------------------------------------------
    '   データ出力
    '----------------------------------------------------------------
    Public Sub SaveData()
        Dim OutTmp As Integer       'ファイル番号
        Dim ret As Boolean
        Dim DenCnt As Integer
        Dim GyouCnt As Integer
        Dim Cnt As Integer
        Dim wrkOutRec As strOutRecord
        Dim wrkOutputData As String

        On Error GoTo ErrPrc

        Dim fp As New frmProg
        fp.Show()

        '一時ファイルオープン
        Dim fileWrite = My.Computer.FileSystem.OpenTextFileWriter(pblInstPath & DIR_OK & tmpFile, False)

        For DenCnt = 1 To pblDenNum
            'プログレスバー表示
            fp.lblsyori.Text = "データ変換中・・・(" & CStr(DenCnt) & "/" & CStr(pblDenNum) & ")"

            '前行摘要クリア
            pblBeforeTekiyou = ""

            pblFirstGyouFlg = True

            For GyouCnt = 1 To MAXGYOU
                '取消行は対象外
                If pblDenRec(DenCnt).Gyou(GyouCnt).Torikeshi = "0" Then
                    '空白行は出力しない（借方貸方両方の科目コードがないとき、又は摘要チェック、摘要がないとき）
                    If ((pblDenRec(DenCnt).Gyou(GyouCnt).Kari.Kamoku <> "") Or _
                        (pblDenRec(DenCnt).Gyou(GyouCnt).Kashi.Kamoku <> "")) Or _
                        (pblDenRec(DenCnt).Gyou(GyouCnt).CopyChk <> "") Or _
                        (Trim(pblDenRec(DenCnt).Gyou(GyouCnt).Tekiyou) <> "") Then

                        '出力データ初期化
                        Call InitOutRec(wrkOutRec)

                        '出力データ設定
                        Call SetData(pblDenRec(DenCnt).Head, pblDenRec(DenCnt).Gyou(GyouCnt), wrkOutRec)

                        '出力形式に変換
                        wrkOutputData = ChgOutRec(wrkOutRec)

                        '出力
                        fileWrite.WriteLine(wrkOutputData)

                    End If
                End If
            Next
        Next

        'ファイルクローズ
        fileWrite.Close()

        On Error GoTo 0

        Dim pf As New PCfunc

        '出力ファイル削除
        pf.FileDelete(pblInstPath & DIR_OK & OUTFILE)

        '一時ファイルを出力ファイルにコピー
        FileCopy(pblInstPath & DIR_OK & tmpFile, pblInstPath & DIR_OK & OUTFILE)

        '一時ファイル削除
        pf.FileDelete(pblInstPath & DIR_OK & tmpFile)

        Exit Sub

        ' エラー処理
ErrPrc:
        Dim em As CErr
        em.ErrMessage("データ変換中")

    End Sub


    '----------------------------------------------------------------
    '   出力データ初期化
    '----------------------------------------------------------------
    Private Sub InitOutRec(ByVal OutRec As strOutRecord)
        With OutRec
            .Kugiri = ""
            .Kessan = ""
            .DenDate = ""
            .DenNo = ""
            .Kari.Bumon = ""
            .Kari.Kamoku = ""
            .Kari.Hojo = ""
            .Kari.Kin = ""
            .Kari.Tax = ""
            .Kari.TaxMas = ""
            .Kari.TaxKbn = ""
            .Kari.JigyoKbn = ""
            .Kashi.Bumon = ""
            .Kashi.Kamoku = ""
            .Kashi.Hojo = ""
            .Kashi.Kin = ""
            .Kashi.Tax = ""
            .Kashi.TaxMas = ""
            .Kashi.TaxKbn = ""
            .Kashi.JigyoKbn = ""
            .Tekiyou = ""
        End With
    End Sub


    '----------------------------------------------------------------
    '   出力データ設定
    '----------------------------------------------------------------
    Private Sub SetData(ByVal DenHead As strInHead, ByVal DenGyou As strInGyou, ByVal OutRec As strOutRecord)
        Dim wrkYear As String
        Dim wrkMonth As String
        Dim wrkDay As String
        Dim wrkGenbaFst As String
        Dim wrkGenbaScd As String
        Dim wrkGenba As String
        Dim wrkGyousya As String
        Dim wrkHacchu As String

        '----------
        ' 伝票区切
        '----------
        '複数チェックなし　かつ　伝票最初の行のみ
        If ((DenHead.FukusuChk = "0") And _
            (pblFirstGyouFlg = True)) Then
            OutRec.Kugiri = "*"
            pblFirstGyouFlg = False
        Else
            OutRec.Kugiri = ""
        End If

        '---------------
        ' 決算処理フラグ
        '---------------
        'チェックがあれば"1"、無ければ""
        If DenHead.Kessan = "1" Then
            OutRec.Kessan = "1"
        Else
            OutRec.Kessan = ""
        End If

        '----------
        ' 伝票日付
        '----------
        wrkYear = DenHead.Year
        wrkMonth = DenHead.Month
        wrkDay = DenHead.Day

        '0詰め
        Dim pf As New PCfunc
        wrkYear = pf.AddZero(wrkYear, 2)
        wrkMonth = pf.AddZero(wrkMonth, 2)
        wrkDay = pf.AddZero(wrkDay, 2)

        '年月日結合
        OutRec.DenDate = wrkYear & wrkMonth & wrkDay

        '----------
        ' 伝票番号
        '----------
        OutRec.DenNo = DenHead.DenNo

        '----------------
        ' 借方部門
        '----------------
        OutRec.Kari.Bumon = DenGyou.Kari.Bumon

        '----------------
        ' 借方科目
        '----------------
        OutRec.Kari.Kamoku = DenGyou.Kari.Kamoku

        '----------------
        ' 借方補助
        '----------------
        OutRec.Kari.Hojo = DenGyou.Kari.Hojo

        '----------------
        ' 借方金額
        '----------------
        OutRec.Kari.Kin = DenGyou.Kari.Kin

        '----------------
        ' 借方消費税額
        '----------------
        OutRec.Kari.Tax = ""

        '--------------------
        ' 借方消費税額計算区分
        '--------------------
        If (DenGyou.Kari.TaxMas = "") Then
            OutRec.Kari.TaxMas = fncGetZeiFlag(OutRec.Kari.Kamoku)
        ElseIf (DenGyou.Kari.TaxMas = "0") Then
            OutRec.Kari.TaxMas = ""
        Else
            OutRec.Kari.TaxMas = DenGyou.Kari.TaxMas
        End If

        '----------------
        ' 借方消費税区分
        '----------------
        If DenGyou.Kari.TaxKbn = "" Then
            OutRec.Kari.TaxKbn = ""
        Else
            OutRec.Kari.TaxKbn = DenGyou.Kari.TaxKbn
        End If

        '----------------
        ' 借方事業区分
        '----------------
        OutRec.Kari.JigyoKbn = ""


        '----------------
        ' 貸方部門
        '----------------
        OutRec.Kashi.Bumon = DenGyou.Kashi.Bumon

        '----------------
        ' 貸方科目
        '----------------
        OutRec.Kashi.Kamoku = DenGyou.Kashi.Kamoku

        '----------------
        ' 貸方補助
        '----------------
        OutRec.Kashi.Hojo = DenGyou.Kashi.Hojo

        '----------------
        ' 貸方金額
        '----------------
        OutRec.Kashi.Kin = DenGyou.Kashi.Kin

        '----------------
        ' 貸方消費税額
        '----------------
        OutRec.Kashi.Tax = ""

        '--------------------
        ' 貸方消費税額計算区分
        '--------------------
        If (DenGyou.Kashi.TaxMas = "") Then
            OutRec.Kashi.TaxMas = fncGetZeiFlag(OutRec.Kashi.Kamoku)
        ElseIf (DenGyou.Kashi.TaxMas = "0") Then
            OutRec.Kashi.TaxMas = ""
        Else
            OutRec.Kashi.TaxMas = DenGyou.Kashi.TaxMas
        End If

        '----------------
        ' 貸方消費税区分
        '----------------
        If DenGyou.Kashi.TaxKbn = "" Then
            OutRec.Kashi.TaxKbn = ""
        Else
            OutRec.Kashi.TaxKbn = DenGyou.Kashi.TaxKbn
        End If

        '----------------
        ' 貸方事業区分
        '----------------
        OutRec.Kashi.JigyoKbn = ""

        OutRec.Tekiyou = RTrim(DenGyou.Tekiyou)
        pblBeforeTekiyou = DenGyou.Tekiyou

    End Sub

    '----------------------------------------------------------------
    '   出力データ変換
    '----------------------------------------------------------------
    Private Function ChgOutRec(ByVal OutRec As strOutRecord)
        Dim wrkRetValue As String

        wrkRetValue = OutRec.Kugiri & KANMA & _
                      OutRec.Kessan & KANMA & _
                      OutRec.DenDate & KANMA & _
                      OutRec.DenNo & KANMA & _
                      OutRec.Kari.Bumon & KANMA & _
                      OutRec.Kari.Kamoku & KANMA & _
                      OutRec.Kari.Hojo & KANMA & _
                      OutRec.Kari.Kin & KANMA & _
                      OutRec.Kari.Tax & KANMA & _
                      OutRec.Kari.TaxMas & KANMA & _
                      OutRec.Kari.TaxKbn & KANMA & _
                      OutRec.Kari.JigyoKbn & KANMA & _
                      OutRec.Kashi.Bumon & KANMA & _
                      OutRec.Kashi.Kamoku & KANMA & _
                      OutRec.Kashi.Hojo & KANMA & _
                      OutRec.Kashi.Kin & KANMA & _
                      OutRec.Kashi.Tax & KANMA & _
                      OutRec.Kashi.TaxMas & KANMA & _
                      OutRec.Kashi.TaxKbn & KANMA & _
                      OutRec.Kashi.JigyoKbn & KANMA & _
                      OutRec.Tekiyou

        ChgOutRec = wrkRetValue

    End Function


    '------------------------------------------
    ' 総勘定科目税処理フラグ取得
    ' 戻り値：税処理フラグ
    ' 引　数：勘定科目コード
    '------------------------------------------
    Private Function fncGetZeiFlag(ByVal Code As String) As String

        Dim lLoopCnt As Long
        Dim sRetCode As String

        lLoopCnt = -1 : sRetCode = ""

        For lLoopCnt = 0 To UBound(pblKamokuData)
            With pblKamokuData(lLoopCnt)
                If (.Code = Code) Then
                    If (.IsZei = 0) Or (.IsZei = 1) Then
                        sRetCode = ""
                    ElseIf (.IsZei = 2) Then
                        sRetCode = "1"
                    End If
                    Exit For
                End If
            End With
        Next

        fncGetZeiFlag = sRetCode

    End Function

End Class
