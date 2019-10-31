Imports 仕訳伝票.PCData
Public Class CChkDataPoor
    '---------------------
    'マッチングエラー
    '---------------------
    Public Function ChkNoData(ByVal Loopcnt) As Boolean
        Dim wrkRetValue As Boolean
        Dim ret As String
        Dim Gyou As Integer

        wrkRetValue = True

        With pblDenRec(Loopcnt)
            If (.Head.Year = "") And (.Head.Month = "") And (.Head.Day = "") And _
                (.Head.DenNo = "") And (.Head.Kessan <> "1") And (.Head.FukusuChk <> "1") Then


                '行ループ
                For Gyou = 1 To MAXGYOU

                    '取消行は対象外とする
                    If .Gyou(Gyou).Torikeshi <> "1" Then
                        '借方科目が空欄のとき
                        If (.Gyou(Gyou).Kari.Kamoku = "") And _
                                (.Gyou(Gyou).Kari.Bumon = "") And _
                                (.Gyou(Gyou).Kari.Hojo = "") And _
                                (.Gyou(Gyou).Kari.Kin = "") And _
                                (Trim(.Gyou(Gyou).Kari.TaxMas) = "") And _
                                (Trim(.Gyou(Gyou).Kari.TaxKbn) = "") And _
                                (.Gyou(Gyou).Kashi.Kamoku = "") And _
                                (.Gyou(Gyou).Kashi.Bumon = "") And _
                                (.Gyou(Gyou).Kashi.Hojo = "") And _
                                (.Gyou(Gyou).Kashi.Kin = "") And _
                                (Trim(.Gyou(Gyou).Kashi.TaxMas) = "") And _
                                (Trim(.Gyou(Gyou).Kashi.TaxKbn) = "") And _
                                (Trim(.Gyou(Gyou).Tekiyou) = "") And _
                                (.Gyou(Gyou).CopyChk <> "1") Then

                            wrkRetValue = False
                        Else
                            Exit For
                        End If

                    End If
                Next
            End If
        End With
        If wrkRetValue = False Then
            '明細がない場合
            'エラーテーブル
            pblErrCnt = pblErrCnt + 1
            ReDim Preserve pblErrTBL(pblErrCnt)
            With pblErrTBL(pblErrCnt)
                .ErrDenNo = Loopcnt
                .ErrLINE = 0
                .ErrField = ""
                .ErrData = ""
                .ErrNotes = "認識できない帳票です。"
                .ErrDpPos = DP_DENYEAR
            End With
        End If
        ChkNoData = wrkRetValue

    End Function
    '----------------------------------------------------------------
    '   有効明細チェック　2006.08.29　新規作成
    '----------------------------------------------------------------
    Public Function ChkYukoMeisai(ByVal Loopcnt) As Boolean
        Dim wrkRetValue As Boolean
        Dim ret As String
        Dim Gyou As Integer

        wrkRetValue = True

        '行ループ
        For Gyou = 1 To MAXGYOU

            '取消行は対象外とする
            If pblDenRec(Loopcnt).Gyou(Gyou).Torikeshi = "0" Then
                '借方科目が空欄のとき
                If (pblDenRec(Loopcnt).Gyou(Gyou).Kari.Kamoku = "") And _
                        (pblDenRec(Loopcnt).Gyou(Gyou).Kari.Bumon = "") And _
                        (pblDenRec(Loopcnt).Gyou(Gyou).Kari.Hojo = "") And _
                        (pblDenRec(Loopcnt).Gyou(Gyou).Kari.Kin = "") And _
                        (Trim(pblDenRec(Loopcnt).Gyou(Gyou).Kari.TaxMas) = "") And _
                        (Trim(pblDenRec(Loopcnt).Gyou(Gyou).Kari.TaxKbn) = "") And _
                        (pblDenRec(Loopcnt).Gyou(Gyou).Kashi.Kamoku = "") And _
                        (pblDenRec(Loopcnt).Gyou(Gyou).Kashi.Bumon = "") And _
                        (pblDenRec(Loopcnt).Gyou(Gyou).Kashi.Hojo = "") And _
                        (pblDenRec(Loopcnt).Gyou(Gyou).Kashi.Kin = "") And _
                        (Trim(pblDenRec(Loopcnt).Gyou(Gyou).Kashi.TaxMas) = "") And _
                        (Trim(pblDenRec(Loopcnt).Gyou(Gyou).Kashi.TaxKbn) = "") And _
                        (Trim(pblDenRec(Loopcnt).Gyou(Gyou).Tekiyou) = "") Then

                    wrkRetValue = False
                Else
                    Exit For
                End If

            End If
        Next
        If wrkRetValue = False Then
            '明細がない場合
            'エラーテーブル
            pblErrCnt = pblErrCnt + 1
            ReDim Preserve pblErrTBL(pblErrCnt)
            With pblErrTBL(pblErrCnt)
                .ErrDenNo = Loopcnt
                .ErrLINE = 1
                .ErrField = "借"
                .ErrData = "明細なし"
                .ErrNotes = "有効な明細がありません。"
                .ErrDpPos = DP_KARI_CODE
            End With
        End If
        ChkYukoMeisai = wrkRetValue

    End Function

    '----------------------------------------------------------------
    '   入力不備チェック -NEW- -Sofybit k.yamagiwa- 2004/06/18
    '----------------------------------------------------------------
    Public Function ChkDataPoorNEW(ByVal Loopcnt) As Boolean
        Dim wrkRetValue As Boolean
        Dim ret As String
        Dim Gyou As Integer

        wrkRetValue = True

        '行ループ
        For Gyou = 1 To MAXGYOU

            '取消行は対象外とする
            If pblDenRec(Loopcnt).Gyou(Gyou).Torikeshi = "0" Then

                '部門科目が空欄のとき
                If (pblBumonFlg = True) And (pblDenRec(Loopcnt).Gyou(Gyou).Kari.Bumon = "") And _
                    (pblDenRec(Loopcnt).Gyou(Gyou).Kari.Kamoku <> "") Then
                    wrkRetValue = False

                    'エラーテーブル
                    Call ErrTblPoor(Loopcnt, Gyou, "借", "部門未登録", DP_KARI_CODEB)

                End If

                '借方科目が空欄のとき
                If (pblDenRec(Loopcnt).Gyou(Gyou).Kari.Kamoku = "") Then
                    '他の借方欄に何か記入されていたらNG
                    If ((pblDenRec(Loopcnt).Gyou(Gyou).Kari.Bumon <> "") Or _
                        (pblDenRec(Loopcnt).Gyou(Gyou).Kari.Hojo <> "") Or _
                        (pblDenRec(Loopcnt).Gyou(Gyou).Kari.Kin <> "") Or _
                        (pblDenRec(Loopcnt).Gyou(Gyou).Kari.TaxMas <> "") Or _
                        (pblDenRec(Loopcnt).Gyou(Gyou).Kari.TaxKbn <> "")) Then
                        wrkRetValue = False

                        'エラーテーブル
                        Call ErrTblPoor(Loopcnt, Gyou, "借", "勘定科目未登録", DP_KARI_CODE)

                    End If

                    '借方科目が記入されているとき
                Else
                    '金額欄が空欄のときNG
                    If pblDenRec(Loopcnt).Gyou(Gyou).Kari.Kin = "" Then
                        wrkRetValue = False

                        'エラーテーブル
                        Call ErrTblPoor(Loopcnt, Gyou, "借", "金額未登録", DP_KARI_KIN)

                    End If
                End If

                '部門科目が空欄のとき
                If (pblBumonFlg = True) And (pblDenRec(Loopcnt).Gyou(Gyou).Kashi.Bumon = "") And _
                    (pblDenRec(Loopcnt).Gyou(Gyou).Kashi.Kamoku <> "") Then
                    wrkRetValue = False

                    'エラーテーブル
                    Call ErrTblPoor(Loopcnt, Gyou, "貸", "部門未登録", DP_KASHI_CODEB)

                End If

                '貸方科目が空欄のとき
                If (pblDenRec(Loopcnt).Gyou(Gyou).Kashi.Kamoku = "") Then
                    '他の貸方欄（摘要以外）に何か記入されていたらNG
                    If ((pblDenRec(Loopcnt).Gyou(Gyou).Kashi.Bumon <> "") Or _
                        (pblDenRec(Loopcnt).Gyou(Gyou).Kashi.Hojo <> "") Or _
                        (pblDenRec(Loopcnt).Gyou(Gyou).Kashi.Kin <> "") Or _
                        (pblDenRec(Loopcnt).Gyou(Gyou).Kashi.TaxMas <> "") Or _
                        (pblDenRec(Loopcnt).Gyou(Gyou).Kashi.TaxKbn <> "")) Then
                        wrkRetValue = False

                        'エラーテーブル
                        Call ErrTblPoor(Loopcnt, Gyou, "貸", "勘定科目未登録", DP_KASHI_CODE)

                    End If

                    '貸方科目が記入されているとき
                Else
                    '金額欄が空欄のときNG
                    If pblDenRec(Loopcnt).Gyou(Gyou).Kashi.Kin = "" Then
                        wrkRetValue = False

                        'エラーテーブル
                        Call ErrTblPoor(Loopcnt, Gyou, "貸", "金額未登録", DP_KASHI_KIN)

                    End If
                End If
            End If
        Next

        ChkDataPoorNEW = wrkRetValue

    End Function

    '------------------------------------------------
    '   エラーテーブル
    '------------------------------------------------
    Sub ErrTblPoor(ByVal p_Cnt, ByVal l_CNT, ByVal Taishaku, ByVal Fie, ByVal stDP)

        'エラーテーブルに値を確保
        pblErrCnt = pblErrCnt + 1
        ReDim Preserve pblErrTBL(pblErrCnt)
        With pblErrTBL(pblErrCnt)
            .ErrDenNo = p_Cnt
            .ErrLINE = l_CNT
            .ErrField = Taishaku
            .ErrData = Fie
            .ErrNotes = "データに不備があります。"
            .ErrDpPos = stDP
        End With
    End Sub

End Class
