Imports 仕訳伝票.PCData
Public Class CShowData
    Const TEKIYO_SPACE_ZEN As String = "　"
    Const TEKIYO_SPACE_HAN As String = " "

    Public frmI As New 仕訳伝票.frmInfo
    '----------------------------------------------------------------
    '   伝票データ表示
    '----------------------------------------------------------------
    Public Sub ShowData()
        Dim Gyou As Integer
        Dim ret As String

        On Error GoTo ErrPrc

        frmI.Show()
        If frmI.Enabled = True Then
            frmI.Focus()
        End If

        '---------------
        ' 伝票ページ表示
        '---------------
        frmI.lblNowDen = "（ " & CStr(pblNowden) & "／" & CStr(pblDenNum) & " ）"

        '---------------
        ' 伝票ヘッダ表示
        '---------------
        With frmI
            .DpDenpyo2.DataMatrix(DP_DENYEAR, 0) = pblDenRec(pblNowden).Head.Year
            .DpDenpyo2.DataMatrix(DP_DENMONTH, 0) = pblDenRec(pblNowden).Head.Month
            .DpDenpyo2.DataMatrix(DP_DENDAY, 0) = pblDenRec(pblNowden).Head.Day
            .DpDenpyo2.DataMatrix(DP_DENNO, 0) = pblDenRec(pblNowden).Head.DenNo

        End With

        '西暦のときは二桁表示
        frmI.DpDenpyo2.DataMatrix("lblGengo", 0) = pblComData.Reki

        If pblComData.Hosei = "0" Then
            Dim pf As New PCfunc
            frmI.DpDenpyo2.DataMatrix(DP_DENYEAR, 0) = pf.AddZero(pblDenRec(pblNowden).Head.Year, 2)

            '和暦のときは一桁表示
        Else
            frmI.DpDenpyo2.DataMatrix(DP_DENYEAR, 0) = pblDenRec(pblNowden).Head.Year
        End If

        '決算処理フラグ
        If pblDenRec(pblNowden).Head.Kessan = "1" Then
            frmI.ChkKessan.Value = True
        Else
            frmI.ChkKessan.Value = False
        End If

        '複数枚チェック
        If pblDenRec(pblNowden).Head.FukusuChk = "1" Then
            frmI.chkFukusuChk.Value = True
        Else
            frmI.chkFukusuChk.Value = False
        End If

        frmI.Show()
        '---------------
        ' 伝票行表示
        '---------------
        With frmInfo.DpDenpyo1
            'コントロールを初期化します。
            .DeleteAllRows()

            '行数を固定するかどうかを設定します。
            .RowsFixed = True

            '行ヘッダに行番号を表示する。
            .TLeftHeader(0).Data = "%n%"
            .TLeftHeader(0).AlignHorizontal = dpAlignHCenter

            '行数の最大値を設定します。
            .MaxRows = MAXGYOU

            '行を自動的に追加できるようにします。
            .AutoAddRow = True

            '行を追加します。
            For Gyou = 0 To MAXGYOU - 1
                .InsertRow(Gyou, True)
            Next Gyou

        End With

        For Gyou = 1 To MAXGYOU
            Call ShowGyou(pblDenRec(pblNowden).Gyou(Gyou), Gyou)
        Next

        Call ShowTekiyou()

        With frmInfo.DpDenpyo1
            '頁合計
            .TFooter.Item(1) = Format(pblDenRec(pblNowden).KariTotal, "#,##0")      '借方合計
            .TFooter.Item(3) = Format(pblDenRec(pblNowden).KashiTotal, "#,##0")     '貸方合計
            .TFooter.Item(5) = Format(System.Math.Abs(pblDenRec(pblNowden).KariTotal - _
                            pblDenRec(pblNowden).KashiTotal), "#,##0")              '差額合計

            '差額があれば赤表示
            If .TFooter(5) <> 0 Then
                .TFooter(5).ForeColor = &HFF&
            Else
                .TFooter(5).ForeColor = &H0&
            End If

            '伝票合計
            .TFooter.Item(9) = Format(pblDenRec(pblNowden).Head.Kari_T, "#,##0")       '借方合計
            .TFooter.Item(11) = Format(pblDenRec(pblNowden).Head.Kashi_T, "#,##0")     '貸方合計
            .TFooter.Item(13) = Format(System.Math.Abs(pblDenRec(pblNowden).Head.Kari_T - _
                            pblDenRec(pblNowden).Head.Kashi_T), "#,##0")               '差額合計

            '差額があれば赤表示
            If .TFooter(13) <> 0 Then
                .TFooter(13).ForeColor = &HFF&
            Else
                .TFooter(13).ForeColor = &H0&
            End If

        End With

        '---------------------
        ' スクロールバー設定
        '---------------------
        frmInfo.fltScr.Min = 1
        frmInfo.fltScr.Max = pblDenNum
        frmInfo.fltScr.Value = pblNowden
        frmInfo.fltScr.LargeChange = Int(pblDenNum / 10) + 1

        '------------------------------
        ' 伝票めくりボタン可能・不可設定
        '------------------------------
        frmInfo.btnFirst.Enabled = True
        frmInfo.btnBefore.Enabled = True
        frmInfo.btnNext.Enabled = True
        frmInfo.btnEnd.Enabled = True

        '先頭の伝票のとき
        If (pblNowden = 1) Then
            frmInfo.btnFirst.Enabled = False
            frmInfo.btnBefore.Enabled = False
        End If

        '最終の伝票のとき
        If (pblNowden = pblDenNum) Then
            frmInfo.btnNext.Enabled = False
            frmInfo.btnEnd.Enabled = False
        End If

        '-------------------
        ' ダイアログ色初期化
        '-------------------
        '伝票ヘッダの色初期化
        Call frmInfo.HeaderColor()

        '-------------------
        ' エラー情報表示
        '-------------------
        Call ShowNG_Grid()        '2004/6/18

        '-------------------------
        '   エラー箇所バックカラー  
        '-------------------------
        With frmInfo
            pblKessanColor = vbButtonFace
            .ChkKessan.BackColor = vbButtonFace

            pblFukuColor = vbButtonFace
            .chkFukusuChk.BackColor = vbButtonFace

            pblSagakuColor = BACK_COLOR
            .DpDenpyo1.TFooter(13).BackColor = BACK_COLOR
        End With

        Call Show_NGColor(pblNowden)

        '-------------------
        ' 画像表示
        '-------------------
        Call ShowImage()

        'フォーカス
        If frmI.Enabled = True Then
            frmI.Focus()
        End If

        On Error GoTo 0

        Exit Sub

        ' エラー処理
ErrPrc:
        Dim pf As PCfunc
        pf.ErrMessage("画面表示中")

    End Sub

    '----------------------------------------------------------------
    '   伝票行データ表示
    '----------------------------------------------------------------
    Public Sub ShowGyou(ByVal Gyou As strInGyou, ByVal GyouNum As Integer)
        Dim i As Integer
        Dim Ih As Integer

        '----------------------------------------------------------------------------
        '   借方
        '----------------------------------------------------------------------------
        With frmInfo.DpDenpyo1
            .DataMatrix(DP_KARI_CODEB, GyouNum - 1) = Gyou.Kari.Bumon
            .DataMatrix(DP_KARI_CODE, GyouNum - 1) = Gyou.Kari.Kamoku
            .DataMatrix(DP_KARI_CODEH, GyouNum - 1) = Gyou.Kari.Hojo
            If ChkKinIndi(Gyou.Kari.Kin) = True Then
                .DataMatrix(DP_KARI_KIN, GyouNum - 1) = Format(Gyou.Kari.Kin, "#,###,###,###")
            Else
                .DataMatrix(DP_KARI_KIN, GyouNum - 1) = Gyou.Kari.Kin
            End If
            .DataMatrix(DP_KARI_ZEI_S, GyouNum - 1) = Gyou.Kari.TaxMas
            .DataMatrix(DP_KARI_ZEI, GyouNum - 1) = Gyou.Kari.TaxKbn

            .RowIndex = GyouNum - 1

            'テキストカラーの設定
            '部門コード
            .CellKey = DP_KARI_CODEB
            .CellForeColor = FORE_COLOR
            '借方コード
            .CellKey = DP_KARI_CODE
            .CellForeColor = FORE_COLOR
            '補助コード
            .CellKey = DP_KARI_CODEH
            .CellForeColor = FORE_COLOR
            '借方金額
            .CellKey = DP_KARI_KIN
            .CellForeColor = FORE_COLOR
            '税処理
            .CellKey = DP_KARI_ZEI_S
            .CellForeColor = FORE_COLOR
            '税区分
            .CellKey = DP_KARI_ZEI
            .CellForeColor = FORE_COLOR

        End With

        '借方部門名表示　DenpyoMan
        If IsNumeric(Gyou.Kari.Bumon) Then
            '初期設定
            With frmInfo.DpDenpyo1
                .DataMatrix(DP_KARI_NAMEB, GyouNum - 1) = "存在しないコードです"
                .CellKey = DP_KARI_NAMEB
                .RowIndex = GyouNum - 1
                .CellForeColor = ERROR_COLOR
            End With

            '部門名表示
            For i = 0 To UBound(pblBumonData)
                If Gyou.Kari.Bumon = pblBumonData(i).Code Then
                    frmInfo.DpDenpyo1.DataMatrix(DP_KARI_NAMEB, GyouNum - 1) = pblBumonData(i).Name
                    frmInfo.DpDenpyo1.CellKey = DP_KARI_NAMEB
                    frmInfo.DpDenpyo1.RowIndex = GyouNum - 1
                    frmInfo.DpDenpyo1.CellForeColor = FORE_COLOR
                    Exit For
                End If
            Next i
        End If

        '------> 借方科目表示　DenpyoMan
        If IsNumeric(Gyou.Kari.Kamoku) Then
            '初期設定
            With frmInfo.DpDenpyo1
                .DataMatrix(DP_KARI_NAME, GyouNum - 1) = "存在しないコードです"
                .CellKey = DP_KARI_NAME
                .RowIndex = GyouNum - 1
                .CellForeColor = ERROR_COLOR
            End With

            '科目名表示
            For i = 0 To UBound(pblKamokuData)

                If Gyou.Kari.Kamoku = pblKamokuData(i).Code Then
                    frmInfo.DpDenpyo1.DataMatrix(DP_KARI_NAME, GyouNum - 1) = pblKamokuData(i).Name
                    frmInfo.DpDenpyo1.CellKey = DP_KARI_NAME
                    frmInfo.DpDenpyo1.RowIndex = GyouNum - 1
                    frmInfo.DpDenpyo1.CellForeColor = FORE_COLOR

                    '補助科目表示
                    '①補助コード設定がなければ終わる
                    If pblKamokuData(i).HojoExist = False Then
                        Exit For
                    End If

                    '初期設定
                    With frmInfo.DpDenpyo1
                        If IsNumeric(Gyou.Kari.Hojo) Then
                            .DataMatrix(DP_KARI_NAMEH, GyouNum - 1) = "存在しないコードです"
                        Else
                            .DataMatrix(DP_KARI_NAMEH, GyouNum - 1) = "補助コード未登録"
                        End If

                        .CellKey = DP_KARI_NAMEH
                        .RowIndex = GyouNum - 1
                        .CellForeColor = ERROR_COLOR
                    End With

                    '②補助科目名を検索して表示する
                    For Ih = 0 To UBound(pblKamokuData(i).HojoData)
                        If Gyou.Kari.Hojo = pblKamokuData(i).HojoData(Ih).Code Then
                            frmInfo.DpDenpyo1.DataMatrix(DP_KARI_NAMEH, GyouNum - 1) = pblKamokuData(i).HojoData(Ih).Name
                            frmInfo.DpDenpyo1.CellKey = DP_KARI_NAMEH
                            frmInfo.DpDenpyo1.RowIndex = GyouNum - 1
                            frmInfo.DpDenpyo1.CellForeColor = FORE_COLOR
                            Exit For
                        End If
                    Next Ih

                End If

            Next i

        End If

        '--------------------------------
        '   貸方        ------> DenpyoMan
        '--------------------------------
        With frmInfo.DpDenpyo1
            .DataMatrix(DP_KASHI_CODEB, GyouNum - 1) = Gyou.Kashi.Bumon
            .DataMatrix(DP_KASHI_CODE, GyouNum - 1) = Gyou.Kashi.Kamoku
            .DataMatrix(DP_KASHI_CODEH, GyouNum - 1) = Gyou.Kashi.Hojo
            If ChkKinIndi(Gyou.Kashi.Kin) = True Then
                .DataMatrix(DP_KASHI_KIN, GyouNum - 1) = Format(Gyou.Kashi.Kin, "#,###,###,###")
            Else
                .DataMatrix(DP_KASHI_KIN, GyouNum - 1) = Gyou.Kashi.Kin
            End If
            .DataMatrix(DP_KASHI_ZEI_S, GyouNum - 1) = Gyou.Kashi.TaxMas
            .DataMatrix(DP_KASHI_ZEI, GyouNum - 1) = Gyou.Kashi.TaxKbn

            .RowIndex = GyouNum - 1

            'テキストカラーの設定
            '部門コード
            .CellKey = DP_KASHI_CODEB
            .CellForeColor = FORE_COLOR
            '借方コード
            .CellKey = DP_KASHI_CODE
            .CellForeColor = FORE_COLOR
            '補助コード
            .CellKey = DP_KASHI_CODEH
            .CellForeColor = FORE_COLOR
            '借方金額
            .CellKey = DP_KASHI_KIN
            .CellForeColor = FORE_COLOR
            '税処理
            .CellKey = DP_KASHI_ZEI_S
            .CellForeColor = FORE_COLOR
            '税区分
            .CellKey = DP_KASHI_ZEI
            .CellForeColor = FORE_COLOR

        End With

        '------> 貸方部門名表示　DenpyoMan
        If IsNumeric(Gyou.Kashi.Bumon) Then
            '初期設定
            With frmInfo.DpDenpyo1
                .DataMatrix(DP_KASHI_NAMEB, GyouNum - 1) = "存在しないコードです"
                .CellKey = DP_KASHI_NAMEB
                .RowIndex = GyouNum - 1
                .CellForeColor = ERROR_COLOR
            End With

            '部門名表示
            For i = 0 To UBound(pblBumonData)
                If Gyou.Kashi.Bumon = pblBumonData(i).Code Then
                    frmInfo.DpDenpyo1.DataMatrix(DP_KASHI_NAMEB, GyouNum - 1) = pblBumonData(i).Name
                    frmInfo.DpDenpyo1.CellKey = DP_KASHI_NAMEB
                    frmInfo.DpDenpyo1.RowIndex = GyouNum - 1
                    frmInfo.DpDenpyo1.CellForeColor = FORE_COLOR
                    Exit For
                End If
            Next i

        End If

        '------> 貸方科目表示　DenpyoMan 2004/6/15    -softbit k.yamagiwa-
        If IsNumeric(Gyou.Kashi.Kamoku) Then
            '初期設定
            With frmInfo.DpDenpyo1
                .DataMatrix(DP_KASHI_NAME, GyouNum - 1) = "存在しないコードです"
                .CellKey = DP_KASHI_NAME
                .RowIndex = GyouNum - 1
                .CellForeColor = ERROR_COLOR
            End With

            '勘定科目名表示
            For i = 0 To UBound(pblKamokuData)

                If Gyou.Kashi.Kamoku = pblKamokuData(i).Code Then
                    frmInfo.DpDenpyo1.DataMatrix(DP_KASHI_NAME, GyouNum - 1) = pblKamokuData(i).Name
                    frmInfo.DpDenpyo1.CellKey = DP_KASHI_NAME
                    frmInfo.DpDenpyo1.RowIndex = GyouNum - 1
                    frmInfo.DpDenpyo1.CellForeColor = FORE_COLOR

                    '補助科目表示
                    '①補助コード設定がなければ終わる
                    If pblKamokuData(i).HojoExist = False Then
                        Exit For
                    End If

                    '初期設定
                    With frmInfo.DpDenpyo1
                        If IsNumeric(Gyou.Kashi.Hojo) Then
                            .DataMatrix(DP_KASHI_NAMEH, GyouNum - 1) = "存在しないコードです"
                        Else
                            .DataMatrix(DP_KASHI_NAMEH, GyouNum - 1) = "補助コード未登録"
                        End If

                        .CellKey = DP_KASHI_NAMEH
                        .RowIndex = GyouNum - 1
                        .CellForeColor = ERROR_COLOR
                    End With

                    '②補助科目名を検索して表示する
                    For Ih = 0 To UBound(pblKamokuData(i).HojoData)
                        If Gyou.Kashi.Hojo = pblKamokuData(i).HojoData(Ih).Code Then
                            frmInfo.DpDenpyo1.DataMatrix(DP_KASHI_NAMEH, GyouNum - 1) = pblKamokuData(i).HojoData(Ih).Name
                            frmInfo.DpDenpyo1.CellKey = DP_KASHI_NAMEH
                            frmInfo.DpDenpyo1.RowIndex = GyouNum - 1
                            frmInfo.DpDenpyo1.CellForeColor = FORE_COLOR
                            Exit For
                        End If
                    Next Ih

                End If

            Next i

        End If

        If Trim(Gyou.CopyChk) <> "0" Then
            frmInfo.chkFu(GyouNum - 1).Value = 1
        Else
            frmInfo.chkFu(GyouNum - 1).Value = 0
        End If

        '摘要
        frmInfo.DpDenpyo1.DataMatrix(DP_TEKIYOU, GyouNum - 1) = Gyou.Tekiyou
        frmInfo.DpDenpyo1.CellKey = DP_TEKIYOU
        frmInfo.DpDenpyo1.RowIndex = GyouNum - 1
        frmInfo.DpDenpyo1.CellForeColor = TEKIYOU_COLOR

        If (Gyou.Torikeshi <> "0") Then
            frmInfo.DpDenpyo1.DataMatrix(DP_DELCHK, GyouNum - 1) = vbChecked
        Else
            frmInfo.DpDenpyo1.DataMatrix(DP_DELCHK, GyouNum - 1) = vbUnchecked
        End If

    End Sub


    Public Sub ShowTekiyou()
        Dim Cnt As Integer

        '2007.05.24 コメントにする
        '    With frmInfo.DpDenpyo1
        '        For Cnt = 0 To (MAXGYOU - 1)
        '            If frmInfo.chkFu(Cnt).Value = 0 Then
        '                .DataMatrix(DP_TEKIYOU, Cnt) = .DataMatrix(DP_TEKIYOU, Cnt) & "　"
        '            End If
        '        Next
        '    End With

        For Cnt = 0 To (MAXGYOU - 1)
            Call ShowTekiyou1gyo(Cnt, 0)
        Next

    End Sub


    '----------------------------------------------------------
    ' Nowcnt 行番号
    ' Mode 編集モード　0:初期表示時、1:複数チェックON・OFFの時
    '----------------------------------------------------------
    Public Sub ShowTekiyou1gyo(ByVal Nowcnt As Integer, ByVal Mode As Integer)
        Dim Cnt As Integer
        Dim iSpacePos As Integer
        Dim sTekiyo As String
        Dim sTekiyoW As String

        With frmInfo.DpDenpyo1
            For Cnt = 0 To (MAXGYOU - 1)
                If frmInfo.chkFu(Cnt).Value = 0 Then
                    .DataMatrix(DP_TEKIYOU, Cnt) = .DataMatrix(DP_TEKIYOU, Cnt) & "　"
                End If
            Next
        End With

        With frmInfo.DpDenpyo1
            If Nowcnt <> 0 Then
                '先頭行は摘要複写機能はない
                If (.DataMatrix(DP_DELCHK, Nowcnt - 1) = "0") And (Trim(.DataMatrix(DP_TEKIYOU, Nowcnt - 1)) <> "") Then
                    '前行が取り消し行でなく、かつ摘要記述がある場合、かつ現行の摘要入力がある場合のみ摘要複写が有効
                    If frmInfo.chkFu(Nowcnt).Value = "1" Then
                        '前行が取消行でない場合のみ適用複写が有効
                        '右のスペースは削除する
                        '摘要複写の対象は、１文字目から次のスペースまでとする(全角チェック)
                        iSpacePos = InStr(1, .DataMatrix(DP_TEKIYOU, Nowcnt - 1), TEKIYO_SPACE_ZEN, vbTextCompare)
                        If iSpacePos <= 0 Then
                            '半角スペースチェック
                            iSpacePos = InStr(1, .DataMatrix(DP_TEKIYOU, Nowcnt - 1), TEKIYO_SPACE_HAN, vbTextCompare)
                        End If
                        If iSpacePos <= 0 Then
                            'スペースが見つからない場合は、摘要すべてが複写対象
                            iSpacePos = Len(.DataMatrix(DP_TEKIYOU, Nowcnt - 1)) + 1
                        End If
                        If iSpacePos > 1 Then
                            sTekiyo = Left(.DataMatrix(DP_TEKIYOU, Nowcnt - 1), iSpacePos - 1)
                        Else
                            sTekiyo = ""
                        End If
                        If Trim(sTekiyo) <> "" Then
                            sTekiyoW = .DataMatrix(DP_TEKIYOU, Nowcnt)
                            If Len(sTekiyoW) < Len(sTekiyo) Then
                                sTekiyoW = sTekiyo
                            Else
                                Mid(sTekiyoW, 1) = sTekiyo
                            End If
                            .DataMatrix(DP_TEKIYOU, Nowcnt) = sTekiyoW
                        End If
                    Else
                        If Mode = 1 Then        '複写をON/OFFしたときのみ有効
                            '摘要複写のチェックを解除した場合
                            iSpacePos = InStr(1, .DataMatrix(DP_TEKIYOU, Nowcnt - 1), TEKIYO_SPACE_ZEN, vbTextCompare)
                            If iSpacePos <= 0 Then
                                '半角スペースチェック
                                iSpacePos = InStr(1, .DataMatrix(DP_TEKIYOU, Nowcnt - 1), TEKIYO_SPACE_HAN, vbTextCompare)
                            End If
                            If iSpacePos <= 0 Then
                                'スペースが見つからない場合は、摘要すべてが複写対象
                                iSpacePos = Len(.DataMatrix(DP_TEKIYOU, Nowcnt - 1)) + 1
                            End If
                            If iSpacePos > 1 Then
                                sTekiyo = Left(.DataMatrix(DP_TEKIYOU, Nowcnt - 1), iSpacePos - 1)
                            Else
                                sTekiyo = ""
                            End If
                            'スペース埋めをする
                            If Trim(sTekiyo) <> "" Then
                                sTekiyo = StrConv(Space$(Len(sTekiyo)), vbWide)
                                sTekiyoW = .DataMatrix(DP_TEKIYOU, Nowcnt)
                                If Len(sTekiyoW) < Len(sTekiyo) Then
                                    sTekiyoW = sTekiyo
                                Else
                                    Mid(sTekiyoW, 1) = sTekiyo
                                End If
                                .DataMatrix(DP_TEKIYOU, Nowcnt) = sTekiyoW
                            End If
                        End If
                    End If
                End If
            End If
        End With

        '2008.01.22 摘要にスペースが７個設定されてしまう不具合の対策
        With frmInfo.DpDenpyo1
            For Cnt = 0 To (MAXGYOU - 1)
                '右のスペースを削除する
                .DataMatrix(DP_TEKIYOU, Cnt) = RTrim(.DataMatrix(DP_TEKIYOU, Cnt))
            Next
        End With

    End Sub

    '----------------------------------------------------------------
    '   伝票画像表示
    '----------------------------------------------------------------
    Public Sub ShowImage()
        Dim wrkFileName As String
        Dim ret As Long

        '-------------------------------------------------------------------
        '   修正画面へ組み入れた画像フォームの表示 -softbit k.yamagiwa -
        '-------------------------------------------------------------------

        '画像の出力が無い場合は、画像表示をしない。
        If Trim(pblDenRec(pblNowden).Head.Image) = "" Then
            frmInfo.leadImg.Visible = False
            pblImageFile = ""
            Exit Sub
        End If

        '画像ファイル名取得　---> ＣＳＶ分割後のフォルダから読み込む 2004/6/24
        wrkFileName = WorkDir & DIR_INCSV & pblDenRec(pblNowden).Head.Image

        With frmInfo.leadImg
            '画像ファイルがあるときのみ表示
            If (Dir(wrkFileName) <> "") Then

                .Visible = True

                '画像表示倍率設定
                If miMdlZoomRate = 0 Then
                    .PaintZoomFactor = ZOOM_RATE
                Else
                    .PaintZoomFactor = miMdlZoomRate
                End If

                '   表示位置をリセット
                .DstLeft = 0
                .DstTop = 0

                .SetDstClipRect.DstLeft, .DstTop, .DstWidth, .DstHeight

                '画像ロード
                ret = .Load(wrkFileName, 0, 0, 1)

                pblImageFile = wrkFileName

                '画像ファイルがないとき
            Else
                .Visible = False
                pblImageFile = ""
            End If

        End With

    End Sub

    '----------------------------------------------------------------
    '   エラーリストグリッド表示    -softbit k.yamagiwa-  2004/6/18
    '----------------------------------------------------------------
    Public Sub ShowNG_Grid()
        Dim i As Integer
        Dim lCnt As Integer
        Dim pCnt As Integer
        Dim LineCnt As Integer

        With frmInfo.fgNg
            .Rows = 1
            .FixedCols = 0
            .Cols = 6

            For i = 0 To 5
                .FixedAlignment(i) = flexAlignCenterCenter
            Next i

            .ColAlignment(3) = flexAlignLeftCenter

            .TextMatrix(0, 0) = "頁"
            .TextMatrix(0, 1) = "行"
            .TextMatrix(0, 2) = "貸借"
            .TextMatrix(0, 3) = "データ"
            .TextMatrix(0, 4) = "エラー内容"

            .ColWidth(0) = 350
            .ColWidth(1) = 350
            .ColWidth(2) = 500
            .ColWidth(3) = 1400
            .ColWidth(4) = 3100

            .ColHidden(5) = True

            lCnt = 0

            For pCnt = 1 To pblDenNum

                For LineCnt = 0 To MAXGYOU

                    For i = 1 To (pblErrCnt)

                        If (pblErrTBL(i).ErrDenNo = pCnt) And _
                            (pblErrTBL(i).ErrLINE = LineCnt) Then

                            .Rows = .Rows + 1
                            lCnt = lCnt + 1
                            .TextMatrix(lCnt, 0) = pblErrTBL(i).ErrDenNo
                            .TextMatrix(lCnt, 1) = pblErrTBL(i).ErrLINE
                            .TextMatrix(lCnt, 2) = pblErrTBL(i).ErrField
                            .TextMatrix(lCnt, 3) = pblErrTBL(i).ErrData
                            .TextMatrix(lCnt, 4) = pblErrTBL(i).ErrNotes
                            .TextMatrix(lCnt, 5) = pblErrTBL(i).ErrDpPos

                        End If

                    Next i

                Next LineCnt

            Next pCnt

            frmInfo.lblErrCnt = "エラー件数（" & pblErrCnt & "件）"
            frmInfo.tabData.Tab = TAB_ERR

        End With

    End Sub

    '----------------------------------------------------------------
    '   エラー項目バックカラー切替  -softbit k.yamagiwa-  2004/6/18
    '----------------------------------------------------------------
    Public Sub Show_NGColor(ByVal pCnt)

        Dim i As Integer

        For i = 1 To pblErrCnt

            If (pblErrTBL(i).ErrDenNo = pCnt) Then

                If pblErrTBL(i).ErrLINE = 0 Then

                    Select Case pblErrTBL(i).ErrDpPos
                        '日付関連エラー
                        Case DP_DENYEAR
                            If frmInfo.ChkErrColor.Value = 1 Then

                                With frmInfo.DpDenpyo2
                                    .RowIndex = 0

                                    .CellKey = pblErrTBL(i).ErrDpPos
                                    .CellBackColor = ERROR_BACK_COLOR
                                    .CellKey = DP_DENMONTH
                                    .CellBackColor = ERROR_BACK_COLOR
                                    .CellKey = DP_DENDAY
                                    .CellBackColor = ERROR_BACK_COLOR
                                End With
                            End If

                            '複数枚チェック
                        Case DP_FUKU
                            pblFukuColor = ERROR_BACK_COLOR

                            If frmInfo.ChkErrColor.Value = 1 Then
                                frmInfo.chkFukusuChk.BackColor = ERROR_BACK_COLOR
                            Else
                                frmInfo.chkFukusuChk.BackColor = vbButtonFace
                            End If

                            '決算チェック
                        Case DP_KESSAN
                            pblKessanColor = ERROR_BACK_COLOR

                            If frmInfo.ChkErrColor.Value = 1 Then
                                frmInfo.ChkKessan.BackColor = ERROR_BACK_COLOR
                            Else
                                frmInfo.ChkKessan.BackColor = vbButtonFace
                            End If


                            '差額エラー
                        Case DP_SAGAKU_T
                            pblSagakuColor = ERROR_BACK_COLOR

                            If frmInfo.ChkErrColor.Value = 1 Then
                                frmInfo.DpDenpyo1.TFooter(13).BackColor = ERROR_BACK_COLOR
                            Else
                                frmInfo.DpDenpyo1.TFooter(13).BackColor = BACK_COLOR
                            End If

                    End Select

                Else
                    With frmInfo.DpDenpyo1
                        .CellKey = pblErrTBL(i).ErrDpPos
                        .RowIndex = pblErrTBL(i).ErrLINE - 1
                        .CellBackColor = ERROR_BACK_COLOR
                    End With
                End If

            End If

        Next i

    End Sub


End Class
