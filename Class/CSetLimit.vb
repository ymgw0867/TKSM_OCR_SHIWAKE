Public Class CSetLimit

    '--------------------------------------
    '　　日付入力範囲の設定
    '--------------------------------------
    Public Sub SetLimit()
        Dim wrkLock As Integer
        Dim wrkSt As Integer
        Dim wrkEd As Integer
        Dim wrkKaisi As Integer
        Dim wrkNextDay As typLimitDateData

        wrkLock = CInt(pblLimitData.Lock)
        wrkSt = CInt(pblLimitData.StSoeji)
        wrkEd = CInt(pblLimitData.EdSoeji)
        wrkKaisi = CInt(pblComData.Kaisi)

        ' 通常仕訳の入力期間　とりあえずマスターの指定期間を入れておく
        pblLimitKikan = pblLimitData

        '---> 四半期決算期間の取得  2004/6/20 softbit k.yamagiwa -
        '最初の四半期決算期間
        pblQuaKessanDate_1 = GetQuaKessanDate_1()
        pblBfQuaKessan_1 = pblQuaKessanDate_1

        '2度目の四半期決算期間
        pblQuaKessanDate_2 = GetQuaKessanDate_2()
        pblBfQuaKessan_2 = pblQuaKessanDate_2

        '3度目の四半期決算期間
        pblQuaKessanDate_3 = GetQuaKessanDate_3()
        pblBfQuaKessan_3 = pblQuaKessanDate_3

        ' 中間期決算期間の取得
        pblMidKessanDate = GetMidKessanDate()
        pblBfMidKessan = pblMidKessanDate

        ' 決算期間の取得
        pblKessanDate = GetKessanDate()
        pblBfKessan = pblKessanDate

        ' 使用可のフラグON
        pblLimitData.Flag = True
        pblLimitKikan.Flag = True
        pblMidKessanDate.Flag = True
        pblKessanDate.Flag = True

        Select Case wrkLock
            ' 入力制限なしの場合
            Case 0
                ' 入力開始月が中間期決算月以降の場合
                If wrkKaisi > 5 Then
                    ' 中間期決算期間の入力を禁止
                    pblMidKessanDate.Flag = False
                End If

                ' 入力期間表示
                Call ShowLimit(pblLimitKikan, 0, pblLimitKikan, 0)

                ' 指定期間を入力禁止
            Case 1
                If 0 <= wrkEd And wrkEd <= 5 Then
                    ' 通常仕訳　指定期間の翌日から期末まで
                    wrkNextDay = GetNextDay(pblLimitData)
                    pblLimitKikan.FromYear = wrkNextDay.FromYear
                    pblLimitKikan.FromMonth = wrkNextDay.FromMonth
                    pblLimitKikan.FromDay = wrkNextDay.FromDay
                    pblLimitKikan.ToYear = pblComData.ToYear
                    pblLimitKikan.ToMonth = pblComData.ToMonth
                    pblLimitKikan.ToDay = pblComData.ToDay

                    ' 入力期間表示
                    Call ShowLimit(pblLimitKikan, 0, pblKessanDate, 2)

                ElseIf wrkEd = 6 Then
                    ' 指定期間終了日が、中間期決算期間の終了日より後でなければ通常処理
                    If JudgeDate(pblLimitData.ToYear, pblLimitData.ToMonth, pblLimitData.ToDay, _
                                 pblMidKessanDate.ToYear, pblMidKessanDate.ToMonth, pblMidKessanDate.ToDay) = True Then
                        ' 通常仕訳　中間期決算期間の翌日から期末まで
                        wrkNextDay = GetNextDay(pblMidKessanDate)
                        pblLimitKikan.FromYear = wrkNextDay.FromYear
                        pblLimitKikan.FromMonth = wrkNextDay.FromMonth
                        pblLimitKikan.FromDay = wrkNextDay.FromDay
                        pblLimitKikan.ToYear = pblComData.ToYear
                        pblLimitKikan.ToMonth = pblComData.ToMonth
                        pblLimitKikan.ToDay = pblComData.ToDay

                        ' 中間期決算期間　指定期間の翌日から
                        wrkNextDay = GetNextDay(pblLimitData)
                        pblMidKessanDate.FromYear = wrkNextDay.FromYear
                        pblMidKessanDate.FromMonth = wrkNextDay.FromMonth
                        pblMidKessanDate.FromDay = wrkNextDay.FromDay

                        ' 入力期間表示
                        Call ShowLimit(pblMidKessanDate, 1, pblKessanDate, 2)

                    Else
                        ' 通常仕訳　指定期間の翌日から期末まで
                        wrkNextDay = GetNextDay(pblLimitData)
                        pblLimitKikan.FromYear = wrkNextDay.FromYear
                        pblLimitKikan.FromMonth = wrkNextDay.FromMonth
                        pblLimitKikan.FromDay = wrkNextDay.FromDay
                        pblLimitKikan.ToYear = pblComData.ToYear
                        pblLimitKikan.ToMonth = pblComData.ToMonth
                        pblLimitKikan.ToDay = pblComData.ToDay

                        ' 中間期決算を使用禁止
                        pblMidKessanDate.Flag = False

                        ' 入力期間表示
                        Call ShowLimit(wrkNextDay, 0, pblKessanDate, 2)

                    End If


                ElseIf 7 <= wrkEd And wrkEd <= 12 Then
                    ' 通常仕訳　指定期間の翌日から期末まで
                    wrkNextDay = GetNextDay(pblLimitData)
                    pblLimitKikan.FromYear = wrkNextDay.FromYear
                    pblLimitKikan.FromMonth = wrkNextDay.FromMonth
                    pblLimitKikan.FromDay = wrkNextDay.FromDay
                    pblLimitKikan.ToYear = pblComData.ToYear
                    pblLimitKikan.ToMonth = pblComData.ToMonth
                    pblLimitKikan.ToDay = pblComData.ToDay

                    ' 中間期決算を使用禁止
                    pblMidKessanDate.Flag = False

                    ' 入力期間表示
                    Call ShowLimit(pblLimitKikan, 0, pblKessanDate, 2)

                ElseIf wrkEd = 13 Then
                    ' 通常仕訳の使用禁止
                    pblLimitKikan.Flag = False

                    ' 中間期決算の使用禁止
                    pblMidKessanDate.Flag = False

                    ' 指定範囲末と期末が同じ場合
                    If pblLimitData.ToDay = pblComData.ToDay Then
                        ' 決算の使用禁止
                        pblKessanDate.Flag = False
                    Else
                        ' 決算期間　指定期間の翌日から
                        wrkNextDay = GetNextDay(pblLimitData)
                        pblKessanDate.FromYear = wrkNextDay.FromYear
                        pblKessanDate.FromMonth = wrkNextDay.FromMonth
                        pblKessanDate.FromDay = wrkNextDay.FromDay
                    End If

                    ' 入力期間表示
                    Call ShowLimit(pblKessanDate, 2, pblKessanDate, 2)

                End If

                ' 指定期間のみ入力許可
            Case 2
                If 0 <= wrkSt And wrkSt <= 5 Then

                    If 0 <= wrkEd And wrkEd <= 5 Then
                        ' 中間期決算の使用禁止
                        pblMidKessanDate.Flag = False

                        ' 決算の使用禁止
                        pblKessanDate.Flag = False

                        ' 入力期間表示
                        Call ShowLimit(pblLimitKikan, 0, pblLimitKikan, 0)

                    ElseIf wrkEd = 6 Then
                        ' 通常仕訳　現時点の中間期決算末まで
                        pblLimitKikan.ToYear = pblMidKessanDate.ToYear
                        pblLimitKikan.ToMonth = pblMidKessanDate.ToMonth
                        pblLimitKikan.ToDay = pblMidKessanDate.ToDay

                        ' 指定期間終了日が、中間期決算期間の終了日より後でなければ通常処理
                        If JudgeDate(pblLimitData.ToYear, pblLimitData.ToMonth, pblLimitData.ToDay, _
                                     pblMidKessanDate.ToYear, pblMidKessanDate.ToMonth, pblMidKessanDate.ToDay) = True Then
                            ' 中間期決算期間　指定期間まで
                            pblMidKessanDate.ToYear = pblLimitData.ToYear
                            pblMidKessanDate.ToMonth = pblLimitData.ToMonth
                            pblMidKessanDate.ToDay = pblLimitData.ToDay

                        End If

                        ' 決算の使用禁止
                        pblKessanDate.Flag = False

                        ' 入力期間表示
                        Call ShowLimit(pblLimitKikan, 0, pblMidKessanDate, 1)

                    ElseIf 7 <= wrkEd And wrkEd <= 12 Then
                        ' 決算の使用禁止
                        pblKessanDate.Flag = False

                        ' 入力期間表示
                        Call ShowLimit(pblLimitKikan, 0, pblLimitKikan, 0)

                    ElseIf wrkEd = 13 Then
                        ' 通常仕訳　期末まで
                        pblLimitKikan.ToYear = pblComData.ToYear
                        pblLimitKikan.ToMonth = pblComData.ToMonth
                        pblLimitKikan.ToDay = pblComData.ToDay

                        ' 決算期間　指定期間まで
                        pblKessanDate.ToYear = pblLimitData.ToYear
                        pblKessanDate.ToMonth = pblLimitData.ToMonth
                        pblKessanDate.ToDay = pblLimitData.ToDay

                        ' 入力期間表示
                        Call ShowLimit(pblLimitKikan, 0, pblKessanDate, 2)

                    End If

                ElseIf wrkSt = 6 Then

                    If wrkEd = 6 Then
                        ' 通常仕訳の使用禁止
                        pblLimitKikan.Flag = False

                        ' 指定期間開始日が、中間期決算期間の開始日より前でなければ通常処理
                        If JudgeDate(pblMidKessanDate.FromYear, pblMidKessanDate.FromMonth, pblMidKessanDate.FromDay, _
                                     pblLimitData.FromYear, pblLimitData.FromMonth, pblLimitData.FromDay) = True Then
                            ' 中間期決算期間の開始日 = 指定期間開始日
                            pblMidKessanDate.FromYear = pblLimitData.FromYear
                            pblMidKessanDate.FromMonth = pblLimitData.FromMonth
                            pblMidKessanDate.FromDay = pblLimitData.FromDay

                        End If

                        ' 指定期間終了日が、中間期決算期間の終了日より後でなければ通常処理
                        If JudgeDate(pblLimitData.ToYear, pblLimitData.ToMonth, pblLimitData.ToDay, _
                                     pblMidKessanDate.ToYear, pblMidKessanDate.ToMonth, pblMidKessanDate.ToDay) = True Then
                            ' 中間期決算期間の終了日 = 指定期間終了日
                            pblMidKessanDate.ToYear = pblLimitData.ToYear
                            pblMidKessanDate.ToMonth = pblLimitData.ToMonth
                            pblMidKessanDate.ToDay = pblLimitData.ToDay

                        End If

                        ' 決算の使用禁止
                        pblKessanDate.Flag = False

                        ' 入力期間表示
                        Call ShowLimit(pblMidKessanDate, 1, pblMidKessanDate, 1)

                    ElseIf 7 <= wrkEd And wrkEd <= 12 Then
                        ' 通常仕訳　中間期決算期間の翌日から指定期間末まで
                        wrkNextDay = GetNextDay(pblMidKessanDate)
                        pblLimitKikan.FromYear = wrkNextDay.FromYear
                        pblLimitKikan.FromMonth = wrkNextDay.FromMonth
                        pblLimitKikan.FromDay = wrkNextDay.FromDay

                        ' 指定期間開始日が、中間期決算期間の開始日より前でなければ通常処理
                        If JudgeDate(pblMidKessanDate.FromYear, pblMidKessanDate.FromMonth, pblMidKessanDate.FromDay, _
                                     pblLimitData.FromYear, pblLimitData.FromMonth, pblLimitData.FromDay) = True Then
                            ' 中間期決算期間の開始日 = 指定期間開始日
                            pblMidKessanDate.FromYear = pblLimitData.FromYear
                            pblMidKessanDate.FromMonth = pblLimitData.FromMonth
                            pblMidKessanDate.FromDay = pblLimitData.FromDay

                        End If

                        ' 決算の使用禁止
                        pblKessanDate.Flag = False

                        ' 入力期間表示
                        Call ShowLimit(pblMidKessanDate, 1, pblLimitKikan, 0)

                    ElseIf wrkEd = 13 Then
                        ' 通常仕訳　中間期決算期間の翌日から期末まで
                        wrkNextDay = GetNextDay(pblMidKessanDate)
                        pblLimitKikan.FromYear = wrkNextDay.FromYear
                        pblLimitKikan.FromMonth = wrkNextDay.FromMonth
                        pblLimitKikan.FromDay = wrkNextDay.FromDay
                        pblLimitKikan.ToYear = pblComData.ToYear
                        pblLimitKikan.ToMonth = pblComData.ToMonth
                        pblLimitKikan.ToDay = pblComData.ToDay

                        ' 指定期間開始日が、中間期決算期間の開始日より前でなければ通常処理
                        If JudgeDate(pblMidKessanDate.FromYear, pblMidKessanDate.FromMonth, pblMidKessanDate.FromDay, _
                                     pblLimitData.FromYear, pblLimitData.FromMonth, pblLimitData.FromDay) = True Then
                            ' 中間期決算期間　指定期間から
                            pblMidKessanDate.FromYear = pblLimitData.FromYear
                            pblMidKessanDate.FromMonth = pblLimitData.FromMonth
                            pblMidKessanDate.FromDay = pblLimitData.FromDay

                        End If

                        ' 決算期間　指定期間まで
                        pblKessanDate.ToYear = pblLimitData.ToYear
                        pblKessanDate.ToMonth = pblLimitData.ToMonth
                        pblKessanDate.ToDay = pblLimitData.ToDay

                        ' 入力期間表示
                        Call ShowLimit(pblMidKessanDate, 1, pblKessanDate, 2)

                    End If

                ElseIf 7 <= wrkSt And wrkSt <= 12 Then

                    If 7 <= wrkEd And wrkEd <= 12 Then
                        ' 中間期決算の使用禁止
                        pblMidKessanDate.Flag = False

                        ' 決算の使用禁止
                        pblKessanDate.Flag = False

                        ' 入力期間表示
                        Call ShowLimit(pblLimitKikan, 0, pblLimitKikan, 0)

                    ElseIf wrkEd = 13 Then
                        ' 通常仕訳　指定期間開始から期末まで
                        pblLimitKikan.ToYear = pblComData.ToYear
                        pblLimitKikan.ToMonth = pblComData.ToMonth
                        pblLimitKikan.ToDay = pblComData.ToDay

                        ' 中間期決算の使用禁止
                        pblMidKessanDate.Flag = False

                        ' 決算期間　指定期間まで
                        pblKessanDate.ToYear = pblLimitData.ToYear
                        pblKessanDate.ToMonth = pblLimitData.ToMonth
                        pblKessanDate.ToDay = pblLimitData.ToDay

                        ' 入力期間表示
                        Call ShowLimit(pblLimitKikan, 0, pblKessanDate, 2)

                    End If

                ElseIf wrkSt = 13 And wrkEd = 13 Then
                    ' 通常仕訳の使用禁止
                    pblLimitKikan.Flag = False

                    ' 中間期決算の使用禁止
                    pblMidKessanDate.Flag = False

                    ' 指定期間開始日が、中間期決算期間の開始日より前でなければ通常処理
                    If JudgeDate(pblKessanDate.FromYear, pblKessanDate.FromMonth, pblKessanDate.FromDay, _
                                 pblLimitData.FromYear, pblLimitData.FromMonth, pblLimitData.FromDay) = True Then
                        ' 決算期間 = 指定期間
                        pblKessanDate = pblLimitData

                    End If

                    ' 入力期間表示
                    Call ShowLimit(pblKessanDate, 2, pblKessanDate, 2)

                End If

        End Select

    End Sub

    '--------------------------------------
    '　　決算期間の取得
    '--------------------------------------
    Public Function GetKessanDate() As typLimitDateData
        Dim LimitDate As String
        Dim wrkYear As String
        Dim wrkMonth As String
        Dim wrkDay As String

        wrkYear = pblComData.FromYear
        wrkMonth = pblComData.FromMonth
        wrkDay = pblComData.FromDay

        LimitDate = wrkYear & "/" & wrkMonth & "/" & wrkDay

        '11ヶ月足す
        LimitDate = Format(DateAdd("m", 11, LimitDate), "yyyymmdd")

        '決算期間の代入
        GetKessanDate.FromYear = CStr(Val(Left(LimitDate, 4)))
        GetKessanDate.FromMonth = CStr(Val(Mid(LimitDate, 5, 2)))
        GetKessanDate.FromDay = CStr(Val(Right(LimitDate, 2)))
        GetKessanDate.ToYear = pblComData.ToYear
        GetKessanDate.ToMonth = pblComData.ToMonth
        GetKessanDate.ToDay = pblComData.ToDay

    End Function

    '--------------------------------------
    '　　中間期決算期間の取得
    '--------------------------------------
    Public Function GetMidKessanDate() As typLimitDateData
        Dim LimitDate As String
        Dim FromLimitDate As String
        Dim ToLimitDate As String
        Dim wrkFromYear As String
        Dim wrkFromMonth As String
        Dim wrkFromDay As String
        Dim wrkToYear As String
        Dim wrkToMonth As String
        Dim wrkToDay As String

        '--------------
        '　　開始日
        '--------------

        wrkFromYear = pblComData.FromYear
        wrkFromMonth = pblComData.FromMonth
        wrkFromDay = pblComData.FromDay

        LimitDate = wrkFromYear & "/" & wrkFromMonth & "/" & wrkFromDay

        '期首日に5ヶ月足す
        FromLimitDate = Format(DateAdd("m", 5, LimitDate), "yyyymmdd")

        '中間期決算開始日を代入
        GetMidKessanDate.FromYear = CStr(Val(Left(FromLimitDate, 4)))
        GetMidKessanDate.FromMonth = CStr(Val(Mid(FromLimitDate, 5, 2)))
        GetMidKessanDate.FromDay = CStr(Val(Right(FromLimitDate, 2)))

        '--------------
        '　　終了日
        '--------------
        wrkToYear = pblComData.ToYear
        wrkToMonth = pblComData.ToMonth
        wrkToDay = pblComData.ToDay

        LimitDate = wrkToYear & "/" & wrkToMonth & "/" & wrkToDay

        '期末日から6ヶ月引く
        ToLimitDate = Format(DateAdd("m", -6, LimitDate), "yyyymmdd")

        '中間期決算終了日を代入
        GetMidKessanDate.ToYear = CStr(Val(Left(ToLimitDate, 4)))
        GetMidKessanDate.ToMonth = CStr(Val(Mid(ToLimitDate, 5, 2)))
        GetMidKessanDate.ToDay = CStr(Val(Right(ToLimitDate, 2)))

    End Function

    '------------------------------------------------------------
    '　　四半期決算期間の取得１ -softbit k.yamagiwa- 2004/6/20
    '------------------------------------------------------------
    Public Function GetQuaKessanDate_1() As typLimitDateData
        Dim LimitDate As String
        Dim FromLimitDate As String
        Dim ToLimitDate As String
        Dim wrkFromYear As String
        Dim wrkFromMonth As String
        Dim wrkFromDay As String
        Dim wrkToYear As String
        Dim wrkToMonth As String
        Dim wrkToDay As String

        '--------------
        '　　開始日
        '--------------

        wrkFromYear = pblComData.FromYear
        wrkFromMonth = pblComData.FromMonth
        wrkFromDay = pblComData.FromDay

        LimitDate = wrkFromYear & "/" & wrkFromMonth & "/" & wrkFromDay

        '期首日に2ヶ月足す
        FromLimitDate = Format(DateAdd("m", 2, LimitDate), "yyyymmdd")

        '最初の四半期決算開始日を代入
        GetQuaKessanDate_1.FromYear = CStr(Val(Left(FromLimitDate, 4)))
        GetQuaKessanDate_1.FromMonth = CStr(Val(Mid(FromLimitDate, 5, 2)))
        GetQuaKessanDate_1.FromDay = CStr(Val(Right(FromLimitDate, 2)))

        '--------------
        '　　終了日
        '--------------
        wrkToYear = pblComData.ToYear
        wrkToMonth = pblComData.ToMonth
        wrkToDay = pblComData.ToDay

        LimitDate = wrkToYear & "/" & wrkToMonth & "/" & wrkToDay

        '期末日から9ヶ月引く
        ToLimitDate = Format(DateAdd("m", -9, LimitDate), "yyyymmdd")

        '最初の四半期決算終了日を代入
        GetQuaKessanDate_1.ToYear = CStr(Val(Left(ToLimitDate, 4)))
        GetQuaKessanDate_1.ToMonth = CStr(Val(Mid(ToLimitDate, 5, 2)))
        GetQuaKessanDate_1.ToDay = CStr(Val(Right(ToLimitDate, 2)))

    End Function

    '------------------------------------------------------------
    '　　四半期決算期間の取得２ -softbit k.yamagiwa- 2004/6/20
    '------------------------------------------------------------
    Public Function GetQuaKessanDate_2() As typLimitDateData
        Dim LimitDate As String
        Dim FromLimitDate As String
        Dim ToLimitDate As String
        Dim wrkFromYear As String
        Dim wrkFromMonth As String
        Dim wrkFromDay As String
        Dim wrkToYear As String
        Dim wrkToMonth As String
        Dim wrkToDay As String

        '--------------
        '　　開始日
        '--------------

        wrkFromYear = pblComData.FromYear
        wrkFromMonth = pblComData.FromMonth
        wrkFromDay = pblComData.FromDay

        LimitDate = wrkFromYear & "/" & wrkFromMonth & "/" & wrkFromDay

        '期首日に5ヶ月足す
        FromLimitDate = Format(DateAdd("m", 5, LimitDate), "yyyymmdd")

        '2度目の四半期決算開始日を代入
        GetQuaKessanDate_2.FromYear = CStr(Val(Left(FromLimitDate, 4)))
        GetQuaKessanDate_2.FromMonth = CStr(Val(Mid(FromLimitDate, 5, 2)))
        GetQuaKessanDate_2.FromDay = CStr(Val(Right(FromLimitDate, 2)))

        '--------------
        '　　終了日
        '--------------
        wrkToYear = pblComData.ToYear
        wrkToMonth = pblComData.ToMonth
        wrkToDay = pblComData.ToDay

        LimitDate = wrkToYear & "/" & wrkToMonth & "/" & wrkToDay

        '期末日から6ヶ月引く
        ToLimitDate = Format(DateAdd("m", -6, LimitDate), "yyyymmdd")

        '2度目の四半期決算終了日を代入
        GetQuaKessanDate_2.ToYear = CStr(Val(Left(ToLimitDate, 4)))
        GetQuaKessanDate_2.ToMonth = CStr(Val(Mid(ToLimitDate, 5, 2)))
        GetQuaKessanDate_2.ToDay = CStr(Val(Right(ToLimitDate, 2)))

    End Function

    '------------------------------------------------------------
    '　　四半期決算期間の取得3 -softbit k.yamagiwa- 2004/6/20
    '------------------------------------------------------------
    Public Function GetQuaKessanDate_3() As typLimitDateData
        Dim LimitDate As String
        Dim FromLimitDate As String
        Dim ToLimitDate As String
        Dim wrkFromYear As String
        Dim wrkFromMonth As String
        Dim wrkFromDay As String
        Dim wrkToYear As String
        Dim wrkToMonth As String
        Dim wrkToDay As String

        '--------------
        '　　開始日
        '--------------

        wrkFromYear = pblComData.FromYear
        wrkFromMonth = pblComData.FromMonth
        wrkFromDay = pblComData.FromDay

        LimitDate = wrkFromYear & "/" & wrkFromMonth & "/" & wrkFromDay

        '期首日に8ヶ月足す
        FromLimitDate = Format(DateAdd("m", 8, LimitDate), "yyyymmdd")

        '3度目の四半期決算開始日を代入
        GetQuaKessanDate_3.FromYear = CStr(Val(Left(FromLimitDate, 4)))
        GetQuaKessanDate_3.FromMonth = CStr(Val(Mid(FromLimitDate, 5, 2)))
        GetQuaKessanDate_3.FromDay = CStr(Val(Right(FromLimitDate, 2)))

        '--------------
        '　　終了日
        '--------------
        wrkToYear = pblComData.ToYear
        wrkToMonth = pblComData.ToMonth
        wrkToDay = pblComData.ToDay

        LimitDate = wrkToYear & "/" & wrkToMonth & "/" & wrkToDay

        '期末日から3ヶ月引く
        ToLimitDate = Format(DateAdd("m", -3, LimitDate), "yyyymmdd")

        '3度目の四半期決算終了日を代入
        GetQuaKessanDate_3.ToYear = CStr(Val(Left(ToLimitDate, 4)))
        GetQuaKessanDate_3.ToMonth = CStr(Val(Mid(ToLimitDate, 5, 2)))
        GetQuaKessanDate_3.ToDay = CStr(Val(Right(ToLimitDate, 2)))

    End Function

    Public Function GetNextDay(ByVal wrkLimit As typLimitDateData) As typLimitDateData
        Dim LimitDate As String
        Dim wrkYear As String
        Dim wrkMonth As String
        Dim wrkDay As String

        wrkYear = wrkLimit.ToYear
        wrkMonth = wrkLimit.ToMonth
        wrkDay = wrkLimit.ToDay

        LimitDate = wrkYear & "/" & wrkMonth & "/" & wrkDay

        '11ヶ月足す
        LimitDate = Format(DateAdd("d", 1, LimitDate), "yyyymmdd")

        '決算期間の代入
        GetNextDay.FromYear = CStr(Val(Left(LimitDate, 4)))
        GetNextDay.FromMonth = CStr(Val(Mid(LimitDate, 5, 2)))
        GetNextDay.FromDay = CStr(Val(Right(LimitDate, 2)))
        GetNextDay.Flag = True

    End Function

    Public Function JudgeDate(ByVal FrYear As String, ByVal FrMonth As String, ByVal FrDay As String, _
                              ByVal BkYear As String, ByVal BkMonth As String, ByVal BkDay As String) As Boolean
        Dim RetValue As Boolean
        Dim wrkFrDate As String
        Dim wrkBkDate As String

        RetValue = True

        wrkFrDate = FrYear & "/" & FrMonth & "/" & FrDay
        wrkBkDate = BkYear & "/" & BkMonth & "/" & BkDay

        ' 日付が前後逆だった場合は、NG
        If DateDiff("d", wrkFrDate, wrkBkDate) <= 0 Then
            RetValue = False
        End If

        JudgeDate = RetValue

    End Function

    Private Sub ShowLimit(ByVal FromLimit As typLimitDateData, ByVal FromKind As Integer, _
                              ByVal ToLimit As typLimitDateData, ByVal ToKind As Integer)
        Dim LimitDate As String
        Dim wrkFromYear As String
        Dim wrkFromMonth As String
        Dim wrkFromDay As String
        Dim wrkToYear As String
        Dim wrkToMonth As String
        Dim wrkToDay As String
        Dim wrkLimit As String
        Dim ans As MsgBoxResult

        ' すべての日付が入力禁止の場合終了
        If pblLimitKikan.Flag = False And pblMidKessanDate.Flag = False And pblKessanDate.Flag = False Then

            ans = MessageBox.Show("勘定奉行の伝票入力が禁止になっています。" & vbNewLine & _
                    "勘定奉行の伝票入力期間制限の設定を変更して「OK」をクリックして下さい。", Title, MessageBoxButtons.OKCancel)

            If ans <> MsgBoxResult.Ok Then
                ans = MessageBox.Show("処理を中断してもよろしいですか？", Title, MessageBoxButtons.OKCancel)
                'さらにキャンセルで中断
                If ret = MsgBoxResult.Ok Then
                    Unload(frmProg)
                    Call ErrEnd()
                Else
                    Call ShowLimit(FromLimit, FromKind, ToLimit, ToKind)
                End If
            Else
                Call LoadMaster()
            End If
        End If

        ' 会社情報のタブの書換え（制限の種類が１，２の場合）
        If (pblLimitData.Lock = 1) Or (pblLimitData.Lock = 2) Then

            ' 会計期間のフォーマット
            ' 入力期間開始が通常日付から
            If FromKind = 0 Then
                ' 和暦
                If pblComData.Hosei <> "0" Then
                    wrkFromYear = CStr(Val(FromLimit.FromYear) - Val(pblComData.Hosei))
                    wrkFromYear = AddSpace(wrkFromYear, 2) & "年"
                Else
                    If Len(FromLimit.FromYear) > 2 Then
                        wrkFromYear = Right(FromLimit.FromYear, 2) & "年"
                    End If
                End If
                wrkFromMonth = AddSpace(Val(FromLimit.FromMonth), 2) & "月"
                wrkFromDay = AddSpace(Val(FromLimit.FromDay), 2) & "日"

                ' 入力期間開始が中間決算期間から
            ElseIf FromKind = 1 Then
                wrkFromYear = "中間決算月"
                wrkFromMonth = ""
                wrkFromDay = AddSpace(Val(FromLimit.FromDay), 2) & "日"

                ' 入力期間開始が決算期間から
            ElseIf FromKind = 2 Then
                wrkFromYear = "決算整理月"
                wrkFromMonth = ""
                wrkFromDay = AddSpace(Val(FromLimit.FromDay), 2) & "日"

            End If

            ' 入力期間終了が通常日付
            If ToKind = 0 Then
                ' 和暦
                If pblComData.Hosei <> "0" Then
                    wrkToYear = CStr(Val(ToLimit.ToYear) - Val(pblComData.Hosei))
                    wrkToYear = AddSpace(wrkToYear, 2) & "年"
                Else
                    If Len(ToLimit.ToYear) > 2 Then
                        wrkToYear = Right(ToLimit.ToYear, 2) & "年"
                    End If
                End If
                wrkToMonth = AddSpace(Val(ToLimit.ToMonth), 2) & "月"
                wrkToDay = AddSpace(Val(ToLimit.ToDay), 2) & "日"

                ' 入力期間終了が中間決算期間
            ElseIf ToKind = 1 Then
                wrkToYear = "中間決算月"
                wrkToMonth = ""
                wrkToDay = AddSpace(Val(ToLimit.ToDay), 2) & "日"

                ' 入力期間終了が決算期間
            ElseIf ToKind = 2 Then
                wrkToYear = "決算整理月"
                wrkToMonth = ""
                wrkToDay = AddSpace(Val(ToLimit.ToDay), 2) & "日"

            End If

            ' 表示文字
            wrkLimit = "●伝票入力期間" & vbNewLine & _
                        wrkFromYear & wrkFromMonth & wrkFromDay & "～" & _
                        wrkToYear & wrkToMonth & wrkToDay

        End If

    End Sub

End Class
