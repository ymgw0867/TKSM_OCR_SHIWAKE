Imports 仕訳伝票.PCData

Public Class CChkCombine
    '--------------------
    '   結合枚数チェック  
    '--------------------
    Public Function ChkCombine(ByVal Loopcnt) As Boolean
        Dim wrkRetValue As Boolean
        Dim ret As String
        Dim wrkDen As Integer

        wrkRetValue = True

        pblMaisu = 1
        wrkDen = 1

        '--------------------------------------------
        '   1行目に複数枚チェックが入っていたらNG
        '--------------------------------------------
        If (Loopcnt = 1) And (pblDenRec(Loopcnt).Head.FukusuChk = "1") Then

            'エラーテーブルに値を確保
            pblErrCnt = pblErrCnt + 1
            ReDim Preserve pblErrTBL(pblErrCnt)
            With pblErrTBL(pblErrCnt)
                .ErrDenNo = Loopcnt
                .ErrLINE = 0
                .ErrField = "結合"
                .ErrData = ""
                .ErrNotes = "先頭伝票に複数チェックが入っています。"
                .ErrDpPos = DP_FUKU
            End With

            wrkRetValue = False

        Else
            '複数チェックなし
            If (pblDenRec(Loopcnt).Head.FukusuChk = "0") Then
                pblMaisu = 1
                wrkDen = Loopcnt
                '複数チェックあり
            Else
                pblMaisu = pblMaisu + 1
            End If

            '結合可能枚数を越えた時
            If (pblMaisu > pblCombineMax) Then

                'エラーテーブルに値を確保
                pblErrCnt = pblErrCnt + 1
                ReDim Preserve pblErrTBL(pblErrCnt)
                With pblErrTBL(pblErrCnt)
                    .ErrDenNo = Loopcnt
                    .ErrLINE = 0
                    .ErrField = "結合"
                    .ErrData = pblMaisu
                    .ErrNotes = "最大結合枚数を超えています"
                    .ErrDpPos = DP_FUKU
                End With

                frmInfo.tabData.Tab = TAB_ERR
                wrkRetValue = False
            End If

        End If

        ChkCombine = wrkRetValue

    End Function

    '-------------------------
    '   結合チェック
    '-------------------------
    Public Function ChkCombineDate(ByVal Loopcnt) As Boolean
        Dim wrkRetValue As Boolean
        Dim ret As String

        wrkRetValue = True

        '先頭伝票はネグる
        If Loopcnt > 1 Then

            '複数チェックあり
            If (pblDenRec(Loopcnt).Head.FukusuChk <> "0") Then
                '前伝票と日付が異なっていた場合エラー
                If (Val(pblDenRec(Loopcnt).Head.Year) <> Val(pblDenRec(Loopcnt - 1).Head.Year)) Or _
                    (Val(pblDenRec(Loopcnt).Head.Month) <> Val(pblDenRec(Loopcnt - 1).Head.Month)) Or _
                    (Val(pblDenRec(Loopcnt).Head.Day) <> Val(pblDenRec(Loopcnt - 1).Head.Day)) Then

                    'エラーテーブルに値を確保
                    pblErrCnt = pblErrCnt + 1
                    ReDim Preserve pblErrTBL(pblErrCnt)
                    With pblErrTBL(pblErrCnt)
                        .ErrDenNo = Loopcnt
                        .ErrLINE = 0
                        .ErrField = "結合"
                        .ErrData = pblComData.Reki & pblDenRec(Loopcnt).Head.Year & "年" & pblDenRec(Loopcnt).Head.Month & "月" _
                                & pblDenRec(Loopcnt).Head.Day & "日"
                        .ErrNotes = "結合伝票で日付が異なっています。"
                        .ErrDpPos = DP_DENYEAR
                    End With

                    frmInfo.tabData.Tab = TAB_ERR
                    wrkRetValue = False
                End If
            End If
        End If

        ChkCombineDate = wrkRetValue

    End Function

    '-----------------------
    '   結合 伝票No.チェック 
    '-----------------------
    Public Function ChkCombineDenNo(ByVal Loopcnt) As Boolean
        Dim wrkRetValue As Boolean
        Dim ret As String
        Dim wrkBeforeDenNo As String
        Dim wrkDenNo As String

        wrkRetValue = True

        wrkDenNo = pblDenRec(Loopcnt).Head.DenNo

        ' 不正チェック　　数字以外又はマイナスはNG
        If (wrkDenNo <> "" And IsNumeric(wrkDenNo) = False) Or _
            Val(wrkDenNo) < 0 Then

            'エラーテーブルに値を確保
            pblErrCnt = pblErrCnt + 1
            ReDim Preserve pblErrTBL(pblErrCnt)
            With pblErrTBL(pblErrCnt)
                .ErrDenNo = Loopcnt
                .ErrLINE = 0
                .ErrField = "№"
                .ErrData = wrkDenNo
                .ErrNotes = "伝票№が不正です。"
            End With

            wrkRetValue = False
        End If

        '先頭伝票はネグる
        If Loopcnt > 1 Then
            wrkBeforeDenNo = pblDenRec(Loopcnt - 1).Head.DenNo

            '複数チェックあり
            If (pblDenRec(Loopcnt).Head.FukusuChk <> "0") Then

                '前伝票と伝票No.が異なっていた場合エラー
                If wrkDenNo <> wrkBeforeDenNo Then

                    'エラーテーブルに値を確保
                    pblErrCnt = pblErrCnt + 1
                    ReDim Preserve pblErrTBL(pblErrCnt)
                    With pblErrTBL(pblErrCnt)
                        .ErrDenNo = Loopcnt
                        .ErrLINE = 0
                        .ErrField = "№"
                        .ErrData = wrkDenNo
                        .ErrNotes = "結合伝票で伝票№が異なっています。"
                        .ErrDpPos = DP_DENNO
                    End With

                    wrkRetValue = False
                End If

            End If

        End If

        ChkCombineDenNo = wrkRetValue

    End Function


    '--------------------
    '   結合枚数チェック  
    '--------------------
    Public Function ChkCombineItem(ByVal Loopcnt) As Boolean
        Dim wrkRetValue As Integer
        Dim ret As String
        Dim wrkDen As Integer
        Dim ItemLimit As Integer

        wrkRetValue = 0
        ItemLimit = 0

        '--------------------------------------------
        '   1行目に複数枚チェックが入っていたらNG
        '--------------------------------------------
        ItemLimit = MAX21

        '最大行数を越えた時
        If ChkVersion(Loopcnt) > ItemLimit Then

            'エラーテーブルに値を確保
            pblErrCnt = pblErrCnt + 1
            ReDim Preserve pblErrTBL(pblErrCnt)
            With pblErrTBL(pblErrCnt)
                .ErrDenNo = Loopcnt
                .ErrLINE = 0
                .ErrField = "行数"
                .ErrData = pblItem
                .ErrNotes = "最大処理行数を超えています"
                .ErrDpPos = DP_FUKU
            End With

            frmInfo.tabData.Tab = TAB_ERR
        End If

    End Function

    '----------------------------------------------------
    '   明細行カウント
    '----------------------------------------------------
    Public Function ChkVersion(ByVal Cnt) As Integer

        Dim LoopCntB As Integer

        '複数チェックなし
        If (pblDenRec(Cnt).Head.FukusuChk = "0") Then
            'カウント初期化
            pblItem = 0
        End If

        For LoopCntB = 1 To MAXGYOU
            '取消行はカウントしません
            If pblDenRec(Cnt).Gyou(LoopCntB).Torikeshi = "0" Then

                '空白行でなければ・・・
                If ((pblDenRec(Cnt).Gyou(LoopCntB).Kari.Kamoku <> "") Or _
                    (pblDenRec(Cnt).Gyou(LoopCntB).Kashi.Kamoku <> "")) Or _
                    (pblDenRec(Cnt).Gyou(LoopCntB).CopyChk <> "") Or _
                    (Trim(pblDenRec(Cnt).Gyou(LoopCntB).Tekiyou) <> "") Then

                    'カウントを足す
                    pblItem = pblItem + 1
                End If
            End If
        Next

        ChkVersion = pblItem

    End Function


    '----------------------
    '   結合 決算チェック   
    '----------------------
    Public Function ChkCombineKessan(ByVal Loopcnt) As Boolean
        Dim wrkRetValue As Boolean
        Dim ret As String
        Dim wrkBeforeKessan As String

        wrkRetValue = True

        '先頭伝票はネグる
        If Loopcnt > 1 Then
            wrkBeforeKessan = pblDenRec(Loopcnt - 1).Head.Kessan

            '複数チェックあり
            If (pblDenRec(Loopcnt).Head.FukusuChk <> "0") Then
                '前伝票と伝票No.が異なっていた場合エラー
                If pblDenRec(Loopcnt).Head.Kessan <> wrkBeforeKessan Then

                    'エラーテーブルに値を確保
                    pblErrCnt = pblErrCnt + 1
                    ReDim Preserve pblErrTBL(pblErrCnt)
                    With pblErrTBL(pblErrCnt)
                        .ErrDenNo = Loopcnt
                        .ErrLINE = 0
                        .ErrField = "決算"
                        .ErrData = pblDenRec(Loopcnt).Head.Kessan
                        .ErrNotes = "結合伝票に通常月と決算月整理仕訳が混在しています。"
                        .ErrDpPos = DP_KESSAN
                    End With

                    frmInfo.tabData.Tab = TAB_ERR
                    wrkRetValue = False
                End If

            End If
        End If

        ChkCombineKessan = wrkRetValue

    End Function

End Class
