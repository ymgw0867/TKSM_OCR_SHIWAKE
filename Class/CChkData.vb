Public Class CChkData
    '--------------------------
    '   相手科目未記入チェック  
    '--------------------------
    Public Function ChkAite(ByVal Cnt) As Boolean
        Dim Gyou As Integer                             '行毎ループカウンタ
        Dim wrkRetValue As Boolean                       '戻り値
        Dim wrkNowDen As Integer
        Dim wrkNowDenNo As String
        Dim wrkNowYear As String
        Dim wrkNowMonth As String
        Dim wrkNowDay As String

        wrkRetValue = True

        '現在の伝票番号取得
        wrkNowDenNo = pblDenRec(Cnt).Head.DenNo
        wrkNowYear = pblDenRec(Cnt).Head.Year
        wrkNowMonth = pblDenRec(Cnt).Head.Month
        wrkNowDay = pblDenRec(Cnt).Head.Day
        wrkNowDen = Cnt

        '先頭レコードはフラグ初期化
        If Cnt = 1 Then
            Call FLGClr()
        Else

            '複数チェックなし
            If (pblDenRec(Cnt).Head.FukusuChk = "0") Then

                '-----------------------------------------------
                ' 相手科目未記入エラー
                '-----------------------------------------------
                If (pblFlgKariKamoku = False) And (pblFlgKashiKamoku = True) Then

                    'エラーテーブルに値を確保
                    pblErrCnt = pblErrCnt + 1
                    ReDim Preserve pblErrTBL(pblErrCnt)
                    With pblErrTBL(pblErrCnt)
                        .ErrDenNo = Cnt
                        .ErrLINE = 1
                        .ErrField = "借方"
                        .ErrData = ""
                        .ErrNotes = "勘定科目が未記入です。"
                        .ErrDpPos = DP_KARI_CODE
                    End With

                    wrkRetValue = False

                End If

                If (pblFlgKashiKamoku = False And pblFlgKariKamoku = True) Then

                    'エラーテーブルに値を確保
                    pblErrCnt = pblErrCnt + 1
                    ReDim Preserve pblErrTBL(pblErrCnt)
                    With pblErrTBL(pblErrCnt)
                        .ErrDenNo = Cnt - 1
                        .ErrLINE = 1
                        .ErrField = "貸方"
                        .ErrData = ""
                        .ErrNotes = "勘定科目が未記入です。"
                        .ErrDpPos = DP_KASHI_CODE
                    End With

                End If

                Call FLGClr()

            End If
        End If

        '勘定科目状態を調べる行毎ループ
        For Gyou = 1 To MAXGYOU

            '取消行は対象外とする
            If pblDenRec(Cnt).Gyou(Gyou).Torikeshi = "0" Then

                '相手科目未記入チェック
                If pblDenRec(Cnt).Gyou(Gyou).Kari.Kamoku <> "" Then
                    pblFlgKariKamoku = True
                End If
                If pblDenRec(Cnt).Gyou(Gyou).Kashi.Kamoku <> "" Then
                    pblFlgKashiKamoku = True
                End If

            End If
        Next

        '伝票数まで達したら終了
        If (Cnt = pblDenNum) Then

            '-----------------------------------------------
            ' 相手科目未記入エラー
            '-----------------------------------------------
            If (pblFlgKariKamoku = False And pblFlgKashiKamoku = True) Then

                'エラーテーブルに値を確保
                pblErrCnt = pblErrCnt + 1
                ReDim Preserve pblErrTBL(pblErrCnt)
                With pblErrTBL(pblErrCnt)
                    .ErrDenNo = Cnt
                    .ErrLINE = 1
                    .ErrField = "借方"
                    .ErrData = ""
                    .ErrNotes = "勘定科目が未記入です。"
                    .ErrDpPos = DP_KARI_CODE
                End With

                wrkRetValue = False

            End If

            If (pblFlgKashiKamoku = False And pblFlgKariKamoku = True) Then

                'エラーテーブルに値を確保
                pblErrCnt = pblErrCnt + 1
                ReDim Preserve pblErrTBL(pblErrCnt)
                With pblErrTBL(pblErrCnt)
                    .ErrDenNo = Cnt
                    .ErrLINE = 1
                    .ErrField = "貸方"
                    .ErrData = ""
                    .ErrNotes = "勘定科目が未記入です。"
                    .ErrDpPos = DP_KASHI_CODE
                End With

            End If

        End If

        ChkAite = wrkRetValue
    End Function

    Sub FLGClr()
        '-----------------------------------------
        '       借方貸方科目ステータス初期化
        '-----------------------------------------
        pblFlgKariKamoku = False
        pblFlgKashiKamoku = False
    End Sub


    '---------------------
    '   部門コードチェック 
    '---------------------
    Public Function ChkBumon(ByVal Loopcnt) As Boolean
        Dim wrkRetValue As Boolean
        Dim ret As String
        Dim Gyou As Integer

        wrkRetValue = True

        '---------------
        ' 全データループ
        '---------------

        '行ループ
        For Gyou = 1 To MAXGYOU

            '取消行は対象外とする
            If pblDenRec(Loopcnt).Gyou(Gyou).Torikeshi = "0" Then

                '借方
                If ((ChkBumonIndi(pblDenRec(Loopcnt).Gyou(Gyou).Kari.Bumon) = False)) Then

                    'エラーテーブルに値を確保
                    pblErrCnt = pblErrCnt + 1
                    ReDim Preserve pblErrTBL(pblErrCnt)
                    With pblErrTBL(pblErrCnt)
                        .ErrDenNo = Loopcnt
                        .ErrLINE = Gyou
                        .ErrField = "借"
                        .ErrData = pblDenRec(Loopcnt).Gyou(Gyou).Kari.Bumon
                        .ErrNotes = "不正な部門コードです。"
                        .ErrDpPos = DP_KARI_CODEB
                    End With

                    wrkRetValue = False
                End If

                '貸方
                If ((ChkBumonIndi(pblDenRec(Loopcnt).Gyou(Gyou).Kashi.Bumon) = False)) Then

                    'エラーテーブルに値を確保
                    pblErrCnt = pblErrCnt + 1
                    ReDim Preserve pblErrTBL(pblErrCnt)
                    With pblErrTBL(pblErrCnt)
                        .ErrDenNo = Loopcnt
                        .ErrLINE = Gyou
                        .ErrField = "貸"
                        .ErrData = pblDenRec(Loopcnt).Gyou(Gyou).Kashi.Bumon
                        .ErrNotes = "不正な部門コードです。"
                        .ErrDpPos = DP_KASHI_CODEB
                    End With

                    wrkRetValue = False
                End If

            End If

        Next

        ChkBumon = wrkRetValue

    End Function


    '-----------------------
    '   部門コードチェック
    '-----------------------
    Private Function ChkBumonIndi(ByVal Bumon As String) As Boolean
        Dim Cnt As Integer
        Dim wrkRetValue As Boolean

        '部門なしのときはOK
        If Bumon = "" Then
            wrkRetValue = True

            '部門登録が無しで、部門記入がある時NG
        ElseIf (pblBumonFlg = False) And (Bumon <> "") Then
            wrkRetValue = False

            '数字以外、4桁以上は×
        ElseIf ((IsNumeric(Bumon) = False) Or (Len(Bumon) > 4)) Then
            wrkRetValue = False

            '部門存在チェック
        Else
            wrkRetValue = False
            '部門コード数分ループ
            For Cnt = 0 To UBound(pblBumonData)
                '部門コードがあったらOK
                '--データ変換処理--Yamamoto
                If (pblBumonData(Cnt).Code = Bumon) Then
                    wrkRetValue = True
                    Exit For
                End If
            Next
        End If

        ChkBumonIndi = wrkRetValue

    End Function


    '----------------------------
    '   日付チェック(存在する日付) 
    '----------------------------
    Public Function ChkDate(ByVal Loopcnt) As Boolean
        Dim wrkRetValue As Boolean
        Dim ret As String
        Dim Gyou As Integer
        Dim wrkYear As String
        Dim wrkMonth As String
        Dim wrkDay As String
        Dim wrkHosei As String

        wrkRetValue = True
        wrkHosei = pblComData.Hosei

        wrkYear = pblDenRec(Loopcnt).Head.Year
        wrkMonth = pblDenRec(Loopcnt).Head.Month
        wrkDay = pblDenRec(Loopcnt).Head.Day

        If ((ChkDateIndi(wrkYear, wrkMonth, wrkDay, wrkHosei) = False)) Then

            'エラーテーブルに値を確保
            pblErrCnt = pblErrCnt + 1
            ReDim Preserve pblErrTBL(pblErrCnt)
            With pblErrTBL(pblErrCnt)
                .ErrDenNo = Loopcnt
                .ErrLINE = 0
                .ErrField = "日付"
                .ErrData = pblComData.Reki & wrkYear & "年" & wrkMonth & "月" & wrkDay & "日"
                .ErrNotes = "存在しない日付です。"
                .ErrDpPos = DP_DENYEAR
            End With

            wrkRetValue = False

        End If

        ChkDate = wrkRetValue

    End Function

    '----------------------------------------------------------------
    '   日付チェック
    '----------------------------------------------------------------
    Public Function ChkDateIndi(ByVal Year As String, ByVal Month As String, ByVal Day As String, ByVal Hosei As String) As Boolean
        Dim Cnt As Integer
        Dim wrkRetValue As Boolean
        Dim wrkMaxDay As Integer
        Dim wrkADYear As Integer

        '空欄はNG
        If ((Year = "") Or (Month = "") Or (Day = "")) Then
            wrkRetValue = False

            '数字以外、2桁以上は×
        ElseIf (((IsNumeric(Year) = False) Or (Len(Year) > 2)) Or _
                ((IsNumeric(Month) = False) Or (Len(Month) > 2)) Or _
                ((IsNumeric(Day) = False) Or (Len(Day) > 2))) Then
            wrkRetValue = False
        Else
            wrkRetValue = True

            '判定用仮ループ
            Do While (1)
                '------
                ' 年
                '------
                '和暦のとき
                If (Hosei <> "0") Then
                    '0年はNG
                    If (Val(Year) = 0) Then
                        wrkRetValue = False
                        Exit Do
                    End If
                End If

                '------
                ' 月
                '------
                '0以下、もしくは13以上のときNG
                If ((Val(Month) < 1) Or (Val(Month) > 12)) Then
                    wrkRetValue = False
                    Exit Do
                End If

                '------
                ' 日
                '------
                '2月
                If (Val(Month) = 2) Then
                    '西暦を求める
                    If (Hosei <> "0") Then
                        wrkADYear = Val(Year) + Val(Hosei)
                    Else
                        If Val(Year) < SWITCH_SEIREKI Then
                            wrkADYear = Val(Year) + TO_SEIREKI
                        Else
                            wrkADYear = Val(Year) + TO_SEIREKI_OLD
                        End If
                    End If

                    '400で割り切れると29日
                    If ((wrkADYear Mod 400) = 0) Then
                        wrkMaxDay = 29

                        '100で割り切れると28日
                    ElseIf ((wrkADYear Mod 100) = 0) Then
                        wrkMaxDay = 28

                        '4で割り切れると29日
                    ElseIf ((wrkADYear Mod 4) = 0) Then
                        wrkMaxDay = 29

                        'それ以外は28日
                    Else
                        wrkMaxDay = 28
                    End If

                    '4,6,9,11月は30日
                ElseIf ((Val(Month) = 4) Or (Val(Month) = 6) Or _
                        (Val(Month) = 9) Or (Val(Month) = 11)) Then
                    wrkMaxDay = 30
                    '1,3,5,7,8,10,12は31日
                Else
                    wrkMaxDay = 31
                End If

                '0以下、もしくはMAX日より大きいときNG
                If ((Val(Day) < 1) Or (Val(Day) > wrkMaxDay)) Then
                    wrkRetValue = False
                    Exit Do
                End If

                Exit Do
            Loop

        End If

        ChkDateIndi = wrkRetValue

    End Function

    '--------------------
    '   決算日付チェック　
    '--------------------
    Public Function ChkDateKessan(ByVal Loopcnt) As Boolean
        Dim wrkRetValue As Boolean
        Dim ret As String
        Dim Gyou As Integer
        Dim wrkYear As String
        Dim wrkMonth As String
        Dim wrkDay As String

        wrkRetValue = True

        wrkYear = pblDenRec(Loopcnt).Head.Year
        wrkMonth = pblDenRec(Loopcnt).Head.Month
        wrkDay = pblDenRec(Loopcnt).Head.Day

        '決算チェックがあり、中間期決算をしない場合
        If (pblDenRec(Loopcnt).Head.Kessan = "1") And (pblComData.Middle = FLGOFF) Then
            ' 決算期間のチェック
            If ChkKikanIndi(wrkYear, wrkMonth, wrkDay, pblBfKessan) = False Then
                wrkRetValue = False
            End If

            '決算チェックがあり、中間期決算を行う場合
        ElseIf (pblDenRec(Loopcnt).Head.Kessan = "1") And (pblComData.Middle = FLGON) Then
            ' 中間期決算、決算期間のチェック
            If (ChkKikanIndi(wrkYear, wrkMonth, wrkDay, pblBfMidKessan) = False) And _
               (ChkKikanIndi(wrkYear, wrkMonth, wrkDay, pblBfKessan) = False) Then
                wrkRetValue = False
            End If

            '決算チェックがあり、四半期決算を行う場合
        ElseIf (pblDenRec(Loopcnt).Head.Kessan = "1") And (pblComData.Middle = FLGON_2) Then
            ' 四半期決算、決算期間のチェック
            If (ChkKikanIndi(wrkYear, wrkMonth, wrkDay, pblBfQuaKessan_1) = False) And _
               (ChkKikanIndi(wrkYear, wrkMonth, wrkDay, pblBfQuaKessan_2) = False) And _
               (ChkKikanIndi(wrkYear, wrkMonth, wrkDay, pblBfQuaKessan_3) = False) And _
               (ChkKikanIndi(wrkYear, wrkMonth, wrkDay, pblBfKessan) = False) Then
                wrkRetValue = False
            End If
        End If

        If wrkRetValue = False Then

            'エラーテーブルに値を確保
            pblErrCnt = pblErrCnt + 1
            ReDim Preserve pblErrTBL(pblErrCnt)
            With pblErrTBL(pblErrCnt)
                .ErrDenNo = Loopcnt
                .ErrLINE = 0
                .ErrField = "日付"
                .ErrData = pblComData.Reki & wrkYear & "年" & wrkMonth & "月" & wrkDay & "日"
                .ErrNotes = "決算日ではありません。"
                .ErrDpPos = DP_DENYEAR
            End With

            wrkRetValue = False
        End If

        ChkDateKessan = wrkRetValue

    End Function


    '-------------------------
    '   会計期間チェック -NEW- 
    '-------------------------
    Public Function ChkDateKikan(ByVal Loopcnt) As Boolean
        Dim wrkRetValue As Boolean
        Dim ret As String
        Dim Gyou As Integer
        Dim wrkYear As String
        Dim wrkMonth As String
        Dim wrkDay As String
        Dim wrkLimitDate As typLimitDateData

        wrkRetValue = True

        '会計期間の設定
        wrkLimitDate.FromYear = Trim(pblComData.FromYear)
        wrkLimitDate.FromMonth = Trim(pblComData.FromMonth)
        wrkLimitDate.FromDay = Trim(pblComData.FromDay)
        wrkLimitDate.ToYear = Trim(pblComData.ToYear)
        wrkLimitDate.ToMonth = Trim(pblComData.ToMonth)
        wrkLimitDate.ToDay = Trim(pblComData.ToDay)

        wrkYear = pblDenRec(Loopcnt).Head.Year
        wrkMonth = pblDenRec(Loopcnt).Head.Month
        wrkDay = pblDenRec(Loopcnt).Head.Day

        If ((ChkKikanIndi(wrkYear, wrkMonth, wrkDay, wrkLimitDate) = False)) Then

            'エラーテーブルに値を確保
            pblErrCnt = pblErrCnt + 1
            ReDim Preserve pblErrTBL(pblErrCnt)
            With pblErrTBL(pblErrCnt)
                .ErrDenNo = Loopcnt
                .ErrLINE = 0
                .ErrField = "日付"
                .ErrData = pblComData.Reki & wrkYear & "年" & wrkMonth & "月" & wrkDay & "日"
                .ErrNotes = "会計期間外の日付です。"
                .ErrDpPos = DP_DENYEAR
            End With
            wrkRetValue = False
        End If

        ChkDateKikan = wrkRetValue

    End Function

    '----------------------------------------------------------------
    '   日付入力範囲チェック
    '----------------------------------------------------------------
    Public Function ChkDateLimit(ByVal Loopcnt) As Boolean
        Dim wrkRetValue As Boolean
        Dim ret As String
        Dim Gyou As Integer
        Dim wrkYear As String
        Dim wrkMonth As String
        Dim wrkDay As String

        wrkRetValue = True

        '---------------
        ' 全データループ
        '---------------
        wrkYear = pblDenRec(Loopcnt).Head.Year
        wrkMonth = pblDenRec(Loopcnt).Head.Month
        wrkDay = pblDenRec(Loopcnt).Head.Day

        '決算チェックがない場合
        If (pblDenRec(Loopcnt).Head.Kessan <> "1") Then
            ' 通常入力禁止の場合はNG
            If pblLimitKikan.Flag = False Then
                wrkRetValue = False

                Call ErrDateLimit(Loopcnt, wrkYear, wrkMonth, wrkDay)

                ' 日付のチェック
            ElseIf (ChkKikanIndi(wrkYear, wrkMonth, wrkDay, pblLimitKikan) = False) Then
                wrkRetValue = False

                Call ErrDateLimit(Loopcnt, wrkYear, wrkMonth, wrkDay)

            End If

            '決算チェックがあり、中間期決算をしない場合
        ElseIf (pblDenRec(Loopcnt).Head.Kessan = "1") And (pblComData.Middle = FLGOFF) Then
            ' 決算禁止の場合はNG
            If pblKessanDate.Flag = False Then
                wrkRetValue = False

                Call ErrDateLimit(Loopcnt, wrkYear, wrkMonth, wrkDay)

                ' 決算期間のチェック
            ElseIf ChkKikanIndi(wrkYear, wrkMonth, wrkDay, pblKessanDate) = False Then
                wrkRetValue = False

                Call ErrDateLimit(Loopcnt, wrkYear, wrkMonth, wrkDay)

            End If

            '決算チェックがあり、中間期決算を行う場合
        ElseIf (pblDenRec(Loopcnt).Head.Kessan = "1") And (pblComData.Middle = FLGON) Then
            ' 中間期決算、決算ともに禁止の場合はNG
            If (pblMidKessanDate.Flag = False) And (pblKessanDate.Flag = False) Then
                wrkRetValue = False

                Call ErrDateLimit(Loopcnt, wrkYear, wrkMonth, wrkDay)

                ' 中間期決算のみ禁止
            ElseIf (pblMidKessanDate.Flag = False) And (pblKessanDate.Flag = True) Then
                If ChkKikanIndi(wrkYear, wrkMonth, wrkDay, pblKessanDate) = False Then
                    wrkRetValue = False

                    Call ErrDateLimit(Loopcnt, wrkYear, wrkMonth, wrkDay)

                End If
                ' 決算のみ禁止
            ElseIf (pblMidKessanDate.Flag = True) And (pblKessanDate.Flag = False) Then
                If ChkKikanIndi(wrkYear, wrkMonth, wrkDay, pblMidKessanDate) = False Then
                    wrkRetValue = False

                    Call ErrDateLimit(Loopcnt, wrkYear, wrkMonth, wrkDay)

                End If
                ' 中間期決算、決算ともに許可
            ElseIf (pblMidKessanDate.Flag = True) And (pblKessanDate.Flag = True) Then
                If (ChkKikanIndi(wrkYear, wrkMonth, wrkDay, pblMidKessanDate) = False) And _
                   (ChkKikanIndi(wrkYear, wrkMonth, wrkDay, pblKessanDate) = False) Then
                    wrkRetValue = False

                    Call ErrDateLimit(Loopcnt, wrkYear, wrkMonth, wrkDay)

                End If
            End If
        End If

        ChkDateLimit = wrkRetValue

    End Function

    '----------------------------------------------------------------
    '   日付入力範囲チェック
    '----------------------------------------------------------------
    Public Function ChkKikanIndi(ByVal Year As String, ByVal Month As String, ByVal Day As String, _
                                 ByVal LimitDate As typLimitDateData) As Boolean
        Dim Cnt As Integer
        Dim wrkRetValue As Boolean
        Dim wrkYear As String
        Dim wrkMonth As String
        Dim wrkDay As String
        Dim wrkFromYear As String
        Dim wrkFromMonth As String
        Dim wrkFromDay As String
        Dim wrkToYear As String
        Dim wrkToMonth As String
        Dim wrkToDay As String

        wrkRetValue = True

        wrkYear = Year
        wrkMonth = Month
        wrkDay = Day

        wrkFromYear = Trim(LimitDate.FromYear)
        wrkFromMonth = Trim(LimitDate.FromMonth)
        wrkFromDay = Trim(LimitDate.FromDay)
        wrkToYear = Trim(LimitDate.ToYear)
        wrkToMonth = Trim(LimitDate.ToMonth)
        wrkToDay = Trim(LimitDate.ToDay)

        ' 和暦
        If pblComData.Hosei <> "0" Then
            wrkYear = CStr(Val(wrkYear) + Val(pblComData.Hosei))
            ' 西暦
        Else
            If Val(wrkYear) < SWITCH_SEIREKI Then
                wrkYear = CStr(Val(wrkYear) + TO_SEIREKI)
            Else
                wrkYear = CStr(Val(wrkYear) + TO_SEIREKI_OLD)
            End If
        End If

        '判定用仮ループ
        Do While (1)

            '---------------------
            ' Fromより前のときはNG
            '---------------------
            '年が前はNG
            If (Val(wrkYear) < Val(wrkFromYear)) Then
                wrkRetValue = False
                Exit Do

                '同じ年
            ElseIf (Val(wrkYear) = Val(wrkFromYear)) Then
                '月が前はNG
                If (Val(wrkMonth) < Val(wrkFromMonth)) Then
                    wrkRetValue = False
                    Exit Do

                    '同じ月
                ElseIf (Val(wrkMonth) = Val(wrkFromMonth)) Then
                    '日が前はNG
                    If (Val(wrkDay) < Val(wrkFromDay)) Then
                        wrkRetValue = False
                        Exit Do
                    End If
                End If
            End If

            '---------------------
            ' Toより後のとき
            '---------------------
            '年が後はNG
            If (Val(wrkYear) > Val(wrkToYear)) Then
                wrkRetValue = False
                Exit Do

                '同じ年
            ElseIf (Val(wrkYear) = Val(wrkToYear)) Then
                '月が後はNG
                If (Val(wrkMonth) > Val(wrkToMonth)) Then
                    wrkRetValue = False
                    Exit Do

                    '同じ月
                ElseIf (Val(wrkMonth) = Val(wrkToMonth)) Then
                    '日が後はNG
                    If (Val(wrkDay) > Val(wrkToDay)) Then
                        wrkRetValue = False
                        Exit Do
                    End If
                End If
            End If

            Exit Do
        Loop

        ChkKikanIndi = wrkRetValue

    End Function
    Sub ErrDateLimit(ByVal p_Cnt, ByVal wrkYear, ByVal wrkMonth, ByVal wrkDay)
        '-------------------------------
        '   エラーテーブルに値を確保
        '-------------------------------
        pblErrCnt = pblErrCnt + 1
        ReDim Preserve pblErrTBL(pblErrCnt)
        With pblErrTBL(pblErrCnt)
            .ErrDenNo = p_Cnt
            .ErrLINE = 0
            .ErrField = "日付"
            .ErrData = pblComData.Reki & wrkYear & "年" & wrkMonth & "月" & wrkDay & "日"
            .ErrNotes = "入力範囲外の日付です。"
            .ErrDpPos = DP_DENYEAR
        End With
    End Sub
    '----------------------
    '   補助コードチェック  
    '----------------------
    Public Function ChkHojo(ByVal Loopcnt) As Boolean
        Dim wrkRetValue As Boolean
        Dim ret As String
        Dim Gyou As Integer

        wrkRetValue = True

        '---------------
        ' 全データループ
        '---------------

        '行ループ
        For Gyou = 1 To MAXGYOU

            '取消行は対象外とする
            If pblDenRec(Loopcnt).Gyou(Gyou).Torikeshi = "0" Then

                '借方
                If ((ChkHojoIndi(pblDenRec(Loopcnt).Gyou(Gyou).Kari.Hojo, _
                                 pblDenRec(Loopcnt).Gyou(Gyou).Kari.Kamoku) = False)) Then

                    'エラーテーブルに値を確保
                    pblErrCnt = pblErrCnt + 1
                    ReDim Preserve pblErrTBL(pblErrCnt)
                    With pblErrTBL(pblErrCnt)
                        .ErrDenNo = Loopcnt
                        .ErrLINE = Gyou
                        .ErrField = "借"
                        .ErrData = pblDenRec(Loopcnt).Gyou(Gyou).Kari.Hojo
                        .ErrNotes = "不正な補助科目コードです。"
                        .ErrDpPos = DP_KARI_CODEH
                    End With

                    wrkRetValue = False
                End If

                '貸方
                If ((ChkHojoIndi(pblDenRec(Loopcnt).Gyou(Gyou).Kashi.Hojo, _
                                 pblDenRec(Loopcnt).Gyou(Gyou).Kashi.Kamoku) = False)) Then

                    'エラーテーブルに値を確保
                    pblErrCnt = pblErrCnt + 1
                    ReDim Preserve pblErrTBL(pblErrCnt)
                    With pblErrTBL(pblErrCnt)
                        .ErrDenNo = Loopcnt
                        .ErrLINE = Gyou
                        .ErrField = "貸"
                        .ErrData = pblDenRec(Loopcnt).Gyou(Gyou).Kashi.Hojo
                        .ErrNotes = "不正な補助科目コードです。"
                        .ErrDpPos = DP_KASHI_CODEH
                    End With

                    wrkRetValue = False
                End If
            End If
        Next

        ChkHojo = wrkRetValue

    End Function


    '----------------------------------------------------------------
    '   補助コードチェック
    '----------------------------------------------------------------
    Private Function ChkHojoIndi(ByVal Hojo As String, ByVal Kamoku As String) As Boolean
        Dim Cnt As Integer
        Dim Loopcnt As Integer
        Dim wrkRetValue As Boolean
        Dim wrkExistKamoku As Boolean
        Dim wrkExistHojo As Boolean

        wrkRetValue = True

        '科目と補助がなしのときはOK
        If (Kamoku = "") And (Hojo = "") Then
            wrkRetValue = True

            '勘定科目なし、補助ありはNG
        ElseIf (Kamoku = "") And (Hojo <> "") Then
            wrkRetValue = False

            '空欄以外かつ数字以外、もしくは4桁以上は×
        ElseIf (((Hojo <> "") And (IsNumeric(Hojo) = False)) Or (Len(Hojo) > 4)) Then
            wrkRetValue = False

            '科目存在チェック
        Else
            '勘定科目検索
            For Loopcnt = 0 To UBound(pblKamokuData)

                If pblKamokuData(Loopcnt).Code = Kamoku Then
                    wrkExistKamoku = True
                    Exit For
                End If

            Next

            '勘定科目が見つからなかった時(多分必要ない処理)
            If wrkExistKamoku = False Then
                wrkRetValue = False

                '補助記入が無し、勘定科目の補助設定が無い場合はＯＫ
            ElseIf ((Hojo = "") And (pblKamokuData(Loopcnt).HojoExist = False)) Then
                wrkRetValue = True

                '補助記入が有り、勘定科目の補助設定が無い場合はNG
            ElseIf ((Hojo <> "") And (pblKamokuData(Loopcnt).HojoExist = False)) Then
                wrkRetValue = False

                '補助記入が無く、勘定科目の補助設定がある場合はNG
            ElseIf ((Hojo = "") And (pblKamokuData(Loopcnt).HojoExist = True)) Then
                wrkRetValue = False

                '補助の記入があり、勘定科目の補助設定が有る場合
            Else
                wrkExistHojo = False

                '補助科目リストループ
                For Cnt = 0 To UBound(pblKamokuData(Loopcnt).HojoData)

                    '補助が該当すればＯＫ
                    If (pblKamokuData(Loopcnt).HojoData(Cnt).Code = Hojo) Then

                        wrkExistHojo = True

                        Exit For
                    End If

                Next

                '補助科目がなかったらNG
                If (wrkExistHojo = False) Then
                    wrkRetValue = False
                Else
                    wrkRetValue = True
                End If
            End If

        End If

        ChkHojoIndi = wrkRetValue

    End Function

    '---------------------
    '   科目コードチェック 
    '---------------------
    Public Function ChkKamoku(ByVal Loopcnt) As Boolean
        Dim wrkRetValue As Boolean
        Dim ret As String
        Dim Gyou As Integer

        wrkRetValue = True

        '---------------
        ' 全データループ
        '---------------

        '行ループ
        For Gyou = 1 To MAXGYOU
            '取消行は対象外とする
            If pblDenRec(Loopcnt).Gyou(Gyou).Torikeshi = "0" Then
                '借方
                If ((ChkKamokuIndi(pblDenRec(Loopcnt).Gyou(Gyou).Kari.Kamoku) = False)) Then
                    'エラーテーブルに値を確保
                    pblErrCnt = pblErrCnt + 1
                    ReDim Preserve pblErrTBL(pblErrCnt)
                    With pblErrTBL(pblErrCnt)
                        .ErrDenNo = Loopcnt
                        .ErrLINE = Gyou
                        .ErrField = "借"
                        .ErrData = pblDenRec(Loopcnt).Gyou(Gyou).Kari.Kamoku
                        .ErrNotes = "不正な勘定科目コードです。"
                        .ErrDpPos = DP_KARI_CODE
                    End With
                    frmInfo.tabData.Tab = TAB_KAMOKU
                    wrkRetValue = False
                End If

                '貸方
                If ((ChkKamokuIndi(pblDenRec(Loopcnt).Gyou(Gyou).Kashi.Kamoku) = False)) Then

                    'エラーテーブルに値を確保
                    pblErrCnt = pblErrCnt + 1
                    ReDim Preserve pblErrTBL(pblErrCnt)
                    With pblErrTBL(pblErrCnt)
                        .ErrDenNo = Loopcnt
                        .ErrLINE = Gyou
                        .ErrField = "貸"
                        .ErrData = pblDenRec(Loopcnt).Gyou(Gyou).Kashi.Kamoku
                        .ErrNotes = "不正な勘定科目コードです。"
                        .ErrDpPos = DP_KASHI_CODE
                    End With

                    frmInfo.tabData.Tab = TAB_KAMOKU
                    wrkRetValue = False
                End If

            End If
        Next

        ChkKamoku = wrkRetValue

    End Function

    '----------------------------------------------------------------
    '   科目コードチェック
    '----------------------------------------------------------------
    Private Function ChkKamokuIndi(ByVal Kamoku As String) As Boolean
        Dim Cnt As Integer
        Dim wrkRetValue As Boolean

        '科目なしのときはOK
        If (Kamoku = "") Then
            wrkRetValue = True

            '数字以外、4桁以上は×
        ElseIf ((IsNumeric(Kamoku) = False) Or (Len(Kamoku) > 4)) Then
            wrkRetValue = False

            '科目存在チェック
        Else
            wrkRetValue = False
            '科目コード数分ループ
            For Cnt = 0 To UBound(pblKamokuData)
                '科目コードがあったらOK
                '--データ変換処理--Yamamoto
                If (pblKamokuData(Cnt).Code = Kamoku) Then
                    wrkRetValue = True
                    Exit For
                End If
            Next
        End If

        ChkKamokuIndi = wrkRetValue

    End Function

    '------------------------------------
    '   貸借不一致 及び　金額の不正チェック
    '------------------------------------
    Public Function ChkSum(ByVal Loopcnt) As Boolean
        Dim wrkKariKin As String               '借方金額
        Dim wrkKashiKin As String               '貸方金額
        Dim Gyou As Integer              '行毎ループカウンタ
        Dim wrkRetValue As Boolean              '戻り値
        Dim wrkNowDen As Integer
        Dim wrkNowDenNo As String
        Dim wrkNowYear As String
        Dim wrkNowMonth As String
        Dim wrkNowDay As String
        Dim ix As Integer

        wrkRetValue = True
        ChkSum = wrkRetValue

        '現在の伝票番号取得
        wrkNowDenNo = pblDenRec(Loopcnt).Head.DenNo
        wrkNowYear = pblDenRec(Loopcnt).Head.Year
        wrkNowMonth = pblDenRec(Loopcnt).Head.Month
        wrkNowDay = pblDenRec(Loopcnt).Head.Day
        wrkNowDen = Loopcnt

        '同伝票の金額を加算

        '複数枚チェック
        If Loopcnt = 1 Then
            Call Chkkin_IniTotal()
        Else
            If pblDenRec(Loopcnt).Head.FukusuChk = "0" Then

                For ix = 1 To (pblFukumai + 1)
                    pblDenRec(Loopcnt - ix).Head.Kari_T = pblKari_T
                    pblDenRec(Loopcnt - ix).Head.Kashi_T = pblKashi_T
                Next ix

                Call SumSagaku(Loopcnt - 1)
                Call Chkkin_IniTotal()
            Else
                pblFukumai = pblFukumai + 1
            End If
        End If

        '頁計初期化
        pblDenRec(Loopcnt).KariTotal = 0
        pblDenRec(Loopcnt).KashiTotal = 0

        '行毎ループ
        For Gyou = 1 To MAXGYOU

            '取消行は対象外とする
            If pblDenRec(Loopcnt).Gyou(Gyou).Torikeshi = "0" Then

                wrkKariKin = pblDenRec(Loopcnt).Gyou(Gyou).Kari.Kin
                wrkKashiKin = pblDenRec(Loopcnt).Gyou(Gyou).Kashi.Kin

                '借方金額の不正チェック
                If ChkKinIndi(wrkKariKin) = False Then
                    '借方金額のエラー

                    'エラーテーブルに値を確保
                    pblErrCnt = pblErrCnt + 1
                    ReDim Preserve pblErrTBL(pblErrCnt)
                    With pblErrTBL(pblErrCnt)
                        .ErrDenNo = Loopcnt
                        .ErrLINE = Gyou
                        .ErrField = "借"
                        .ErrData = wrkKariKin
                        .ErrNotes = "不正な金額です。"
                        .ErrDpPos = DP_KARI_KIN
                    End With

                    frmInfo.tabData.Tab = TAB_ERR
                    wrkRetValue = False
                End If

                '貸方金額の不正チェック
                If ChkKinIndi(wrkKashiKin) = False Then
                    '貸方金額のエラー

                    'エラーテーブルに値を確保
                    pblErrCnt = pblErrCnt + 1
                    ReDim Preserve pblErrTBL(pblErrCnt)
                    With pblErrTBL(pblErrCnt)
                        .ErrDenNo = Loopcnt
                        .ErrLINE = Gyou
                        .ErrField = "貸"
                        .ErrData = wrkKashiKin
                        .ErrNotes = "不正な金額です。"
                        .ErrDpPos = DP_KASHI_KIN
                    End With

                    frmInfo.tabData.Tab = TAB_ERR
                    wrkRetValue = False

                End If

                '借方合計
                If IsNumeric(wrkKariKin) Then
                    '頁合計
                    pblDenRec(Loopcnt).KariTotal = pblDenRec(Loopcnt).KariTotal + wrkKariKin
                    '伝票合計
                    pblKari_T = pblKari_T + wrkKariKin
                End If

                '貸方合計
                If IsNumeric(wrkKashiKin) Then
                    '頁合計
                    pblDenRec(Loopcnt).KashiTotal = pblDenRec(Loopcnt).KashiTotal + wrkKashiKin
                    '伝票合計
                    pblKashi_T = pblKashi_T + wrkKashiKin
                End If

            End If

        Next

        '複数枚チェック
        '       最後の伝票？

        If Loopcnt = pblDenNum Then
            For ix = 0 To pblFukumai
                pblDenRec(Loopcnt - ix).Head.Kari_T = pblKari_T
                pblDenRec(Loopcnt - ix).Head.Kashi_T = pblKashi_T
            Next ix

            Call SumSagaku(Loopcnt)

        End If

    End Function

    Sub SumSagaku(ByVal Cnt)
        Dim Sagaku As Currency             '貸借差額

        '--------------
        ' 貸借差額計算 
        '--------------
        If (pblKari_T <> pblKashi_T) Then
            '貸借差額計算
            Sagaku = Abs(pblKari_T - pblKashi_T)

            'エラーテーブルに値を確保
            pblErrCnt = pblErrCnt + 1
            ReDim Preserve pblErrTBL(pblErrCnt)
            With pblErrTBL(pblErrCnt)
                .ErrDenNo = Cnt
                .ErrLINE = 0
                .ErrField = "差額"
                .ErrData = Format(Sagaku, "#,##0")
                .ErrNotes = "貸借の金額に差額があります。"
                .ErrDpPos = DP_SAGAKU_T
            End With

            frmInfo.tabData.Tab = TAB_ERR
            pblSagakuFLG = True

        End If

    End Sub

    '----------------------------------------------------------------
    '   金額チェック
    '----------------------------------------------------------------
    Public Function ChkKinIndi(ByVal Kin As String) As Boolean
        Dim wrkRetValue As Boolean

        ' 金額未記入のときはOK
        If Kin = "" Then
            wrkRetValue = True

            ' 数字以外、11桁以上は×
        ElseIf ((IsNumeric(Kin) = False) Or (Len(Kin) > 10)) Then
            wrkRetValue = False

            ' 金額が０のときはNG
        ElseIf Kin = "0" Then
            wrkRetValue = False

        ElseIf Right(Kin, 1) = "-" Then
            '最後に-はエラーとする
            wrkRetValue = False

        Else
            wrkRetValue = True

        End If

        ChkKinIndi = wrkRetValue

    End Function

    '----------------------------------------------------------------
    '   伝票合計金額初期化
    '----------------------------------------------------------------
    Private Sub Chkkin_IniTotal()
        pblKari_T = 0
        pblKashi_T = 0
        pblFukumai = 0

    End Sub

    '-----------------------
    '   税区分コードチェック 
    '-----------------------
    Public Function ChkTaxKbn(ByVal Loopcnt) As Boolean
        Dim wrkRetValue As Boolean
        Dim ret As String
        Dim Gyou As Integer

        wrkRetValue = True

        '行ループ
        For Gyou = 1 To MAXGYOU

            '取消行は対象外とする
            If pblDenRec(Loopcnt).Gyou(Gyou).Torikeshi = "0" Then

                '借方
                If ((ChkTaxKbnIndi(pblDenRec(Loopcnt).Gyou(Gyou).Kari.TaxKbn) = False)) Then

                    'エラーテーブルに値を確保
                    pblErrCnt = pblErrCnt + 1
                    ReDim Preserve pblErrTBL(pblErrCnt)
                    With pblErrTBL(pblErrCnt)
                        .ErrDenNo = Loopcnt
                        .ErrLINE = Gyou
                        .ErrField = "借"
                        .ErrData = pblDenRec(Loopcnt).Gyou(Gyou).Kari.TaxKbn
                        .ErrNotes = "不正な税区分です"
                        .ErrDpPos = DP_KARI_ZEI
                    End With

                    wrkRetValue = False
                End If

                '貸方
                If ((ChkTaxKbnIndi(pblDenRec(Loopcnt).Gyou(Gyou).Kashi.TaxKbn) = False)) Then

                    'エラーテーブルに値を確保
                    pblErrCnt = pblErrCnt + 1
                    ReDim Preserve pblErrTBL(pblErrCnt)
                    With pblErrTBL(pblErrCnt)
                        .ErrDenNo = Loopcnt
                        .ErrLINE = Gyou
                        .ErrField = "貸"
                        .ErrData = pblDenRec(Loopcnt).Gyou(Gyou).Kashi.TaxKbn
                        .ErrNotes = "不正な税区分です。"
                        .ErrDpPos = DP_KASHI_ZEI
                    End With

                    wrkRetValue = False
                End If
            End If
        Next

        ChkTaxKbn = wrkRetValue

    End Function

    '----------------------------------------------------------------
    '   税区分コードチェック
    '----------------------------------------------------------------
    Private Function ChkTaxKbnIndi(ByVal TaxKbn As String) As Boolean
        Dim Cnt As Integer
        Dim wrkRetValue As Boolean

        '税区分なしのときはOK
        If TaxKbn = "" Then
            wrkRetValue = True

            '数字以外、4桁以上は×
        ElseIf ((IsNumeric(TaxKbn) = False) Or (Len(TaxKbn) > 4)) Then
            wrkRetValue = False

            '税区分存在チェック
        Else
            wrkRetValue = False
            '税区分コード数分ループ
            For Cnt = 0 To UBound(pblTaxKbnData)
                '税区分コードがあったらOK
                '--データ変換処理--Yamamoto
                If (pblTaxKbnData(Cnt).Code = TaxKbn) Then
                    wrkRetValue = True
                    Exit For
                End If
            Next
        End If

        ChkTaxKbnIndi = wrkRetValue

    End Function


    '---------------------------------
    '   消費税計算区分のコードチェック  
    '---------------------------------
    Public Function ChkOther(ByVal Loopcnt) As Boolean
        Dim wrkRetValue As Boolean
        Dim ret As String
        Dim Gyou As Integer

        wrkRetValue = True

        '行ループ
        For Gyou = 1 To MAXGYOU

            '取消行は対象外とする
            If pblDenRec(Loopcnt).Gyou(Gyou).Torikeshi = "0" Then

                '--------------------------------
                '消費税計算区分のチェック
                '--------------------------------
                '　借方
                If ((ChkTaxMasIndi(pblDenRec(Loopcnt).Gyou(Gyou).Kari.TaxMas) = False)) Then

                    'エラーテーブルに値を確保
                    pblErrCnt = pblErrCnt + 1
                    ReDim Preserve pblErrTBL(pblErrCnt)
                    With pblErrTBL(pblErrCnt)
                        .ErrDenNo = Loopcnt
                        .ErrLINE = Gyou
                        .ErrField = "借"
                        .ErrData = pblDenRec(Loopcnt).Gyou(Gyou).Kari.TaxMas
                        .ErrNotes = "不正な税処理です。"
                        .ErrDpPos = DP_KARI_ZEI_S
                    End With

                    wrkRetValue = False
                End If

                '　貸方
                If ((ChkTaxMasIndi(pblDenRec(Loopcnt).Gyou(Gyou).Kashi.TaxMas) = False)) Then

                    'エラーテーブルに値を確保
                    pblErrCnt = pblErrCnt + 1
                    ReDim Preserve pblErrTBL(pblErrCnt)
                    With pblErrTBL(pblErrCnt)
                        .ErrDenNo = Loopcnt
                        .ErrLINE = Gyou
                        .ErrField = "貸"
                        .ErrData = pblDenRec(Loopcnt).Gyou(Gyou).Kashi.TaxMas
                        .ErrNotes = "不正な税処理です。"
                        .ErrDpPos = DP_KASHI_ZEI_S
                    End With

                    wrkRetValue = False
                End If

            End If
        Next

        ChkOther = wrkRetValue

    End Function

    '----------------------------------------------------------------
    '   消費税計算区分のコードチェック
    '----------------------------------------------------------------
    Private Function ChkTaxMasIndi(ByVal TaxMas As String) As Boolean
        Dim wrkRetValue As String

        '未記入、1か0ならＯＫ
        If ((TaxMas = "") Or (TaxMas = "1") Or (TaxMas = "0")) Then
            wrkRetValue = True
        Else
            wrkRetValue = False
        End If

        ChkTaxMasIndi = wrkRetValue

    End Function


    '--------------------
    '   摘要チェック 
    '--------------------
    Public Function ChkTekiyou(ByVal Loopcnt) As Boolean
        Dim wrkRetValue As Boolean
        Dim Cnt As Integer

        wrkRetValue = True

        For Cnt = 1 To MAXGYOU
            '摘要は１明細全角２０、半角４０がMAX
            If LenB(StrConv(pblDenRec(Loopcnt).Gyou(Cnt).Tekiyou, vbFromUnicode)) > 40 Then
                pblErrCnt = pblErrCnt + 1
                ReDim Preserve pblErrTBL(pblErrCnt)
                With pblErrTBL(pblErrCnt)
                    .ErrDenNo = Loopcnt
                    .ErrLINE = Cnt
                    .ErrField = "摘要"
                    .ErrData = pblDenRec(Loopcnt).Gyou(Cnt).Tekiyou
                    .ErrNotes = "入力文字数が超えています。"
                    .ErrDpPos = DP_TEKIYOU
                End With

                wrkRetValue = False
            End If
        Next

        ChkTekiyou = wrkRetValue
    End Function


    '----------------------------------------------------------------
    '   摘要のみチェック
    '----------------------------------------------------------------
    Public Function ChkTekiyouOnly(ByVal Loopcnt) As Boolean
        Dim wrkRetValue As Boolean
        Dim CheckFlg As Integer
        Dim CntW As Integer
        Dim Cnt As Integer

        ChkTekiyouOnly = True

        If pblDenRec(Loopcnt).Head.FukusuChk = "0" Then
            '複写チェックがチェックされていない場合のみチェックを行う
            CheckFlg = False
            For Cnt = 1 To MAXGYOU
                With pblDenRec(Loopcnt).Gyou(Cnt)
                    If .Torikeshi = "0" Then
                        '取消行でない場合で、摘要の入力が１明細でもあれば摘要のみチェックを行う
                        If (Trim(.Tekiyou) <> "") Then
                            CheckFlg = True
                            CntW = Cnt
                            Exit For
                        End If
                    End If
                End With
            Next

            If CheckFlg = True Then
                '摘要のみチェックを行う
                wrkRetValue = False

                For Cnt = 1 To MAXGYOU
                    With pblDenRec(Loopcnt).Gyou(Cnt)
                        If .Torikeshi = "0" Then
                            '取消行でない場合
                            If (.Kari.Kamoku <> "") Or (.Kari.Kin <> "") Or (.Kashi.Kamoku <> "") Or (.Kashi.Kin <> "") Then
                                '借方・貸方の科目または金額が入力されていればOK
                                wrkRetValue = True
                                Exit For
                            End If
                        End If
                    End With
                Next

                If wrkRetValue = False Then
                    '摘要のみの場合
                    pblErrCnt = pblErrCnt + 1
                    ReDim Preserve pblErrTBL(pblErrCnt)
                    With pblErrTBL(pblErrCnt)
                        .ErrDenNo = Loopcnt
                        .ErrLINE = CntW
                        .ErrField = "摘要"
                        .ErrData = pblDenRec(Loopcnt).Gyou(CntW).Tekiyou
                        .ErrNotes = "勘定科目または金額が入力されていません。"
                        .ErrDpPos = DP_TEKIYOU
                    End With
                End If
                ChkTekiyouOnly = wrkRetValue
            End If
        End If

    End Function



End Class
