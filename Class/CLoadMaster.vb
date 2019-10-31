'Imports 固定項目保守.Cfunction
'Imports 仕訳伝票.frmInfo
Imports 仕訳伝票.PCData
Imports 仕訳伝票.frmProg
Imports Microsoft.VisualBasic.Strings
Imports 仕訳伝票.CMain

Public Class CLoadMaster

    '----------------------------------------------------------------
    '   各種情報ロード
    '----------------------------------------------------------------
    Public Sub LoadMaster()
        Dim AdoData As New ADODB.Connection     'ADOのConnectionオブジェクトを生成
        Dim Rs As ADODB.Recordset               'Recordsetオブジェクトの生成
        Dim wrkConnectInfo As String
        Dim Cnt As Integer
        Dim Loopcnt As Integer
        Dim ErrMsg As String
        Dim cntKamoku As Integer                '補助科目用
        Dim wrkBfNcd As String
        Dim wrkNcd As String
        Dim wrkKamoku As String
        Dim kamokuFlg As Boolean
        Dim strOtherName As String
        Dim lCnt As Integer          'フレックスグリッドの表示行添え字　2004/6/14
        Dim pf As New PCfunc

        On Error GoTo ErrPrc


        '処理中表示開始
        Dim fp As New frmProg
        fp.Show()

        '接続文字列
        Dim pf As New PCfunc
        wrkConnectInfo = pf.GetDbConnect(pblDbName)

        'DBを開く
        AdoData.ConnectionString = wrkConnectInfo
        AdoData.Open()

        '-----------------------------------------------------
        ' 会社データ取得
        '-----------------------------------------------------
        fp.lblsyori.Text = "データロード中・・・会社データ"
        'frmProg.prgBar.Value = 0
        fp.lblsyori.Update()
        ErrMsg = "会社データ取得中"

        'レコードを開く
        Rs = New ADODB.Recordset
        Rs.Open("SELECT sDateKisyu,sDateKimatu,tiKaisi,sCorpNm,sGngo,sHosei,tiIsMiddle FROM wdhead" _
             , AdoData, , ADODB.LockTypeEnum.adLockReadOnly)

        With pblComData
            .Name = Trim(Rs.Fields("sCorpNm").Value)
            .FromYear = Mid(Trim(Rs.Fields("sDateKisyu").Value), 1, 4)
            .FromMonth = Mid(Trim(Rs.Fields("sDateKisyu").Value), 5, 2)
            .FromDay = Strings.Right(Trim(Rs.Fields("sDateKisyu").Value), 2)
            .ToYear = Mid(Trim(Rs.Fields("sDateKimatu").Value), 1, 4)
            .ToMonth = Mid(Trim(Rs.Fields("sDateKimatu").Value), 5, 2)
            .ToDay = Strings.Right(Trim(Rs.Fields("sDateKimatu").Value), 2)
            .Kaisi = Trim(CStr(Rs.Fields("tiKaisi").Value))
            .Gengou = Trim(Rs.Fields("sGngo").Value)
            .Hosei = Trim(Rs.Fields("sHosei").Value)
            .Middle = Trim(CStr(Rs.Fields("tiIsMiddle").Value))
        End With
        pblComData.TaxMas = frmComSelect.TaxMas

        '接続を切断
        Rs.Close()

        '西暦のとき
        If pblComData.Hosei = "0" Then
            pblComData.Reki = "20"

            '和暦のとき
        Else
            pblComData.Reki = pblComData.Gengou
        End If

        '会社データ表示
        Call ShowComData()

        '伝票入力指定期間を取得
        Rs = New ADODB.Recordset
        Rs.Open("SELECT tiStSoeji,sDnStDate,tiEdSoeji,sDnEdDate,tiIsLock FROM wjdnpyo2" _
             , AdoData, , ADODB.LockTypeEnum.adLockReadOnly)

        With pblLimitData
            .StSoeji = Trim(CStr(Rs.Fields("tiStSoeji").Value))
            .FromYear = Mid(Trim(Rs.Fields("sDnStDate").Value), 1, 4)
            .FromMonth = Mid(Trim(Rs.Fields("sDnStDate").Value), 5, 2)
            .FromDay = Strings.Right(Trim(Rs.Fields("sDnStDate").Value), 2)
            .EdSoeji = Trim(CStr(Rs.Fields("tiEdSoeji").Value))
            .ToYear = Mid(Trim(Rs.Fields("sDnEdDate").Value), 1, 4)
            .ToMonth = Mid(Trim(Rs.Fields("sDnEdDate").Value), 5, 2)
            .ToDay = Strings.Right(Trim(Rs.Fields("sDnEdDate").Value), 2)
            .Lock = Trim(CStr(Rs.Fields("tiIsLock").Value))
        End With

        '接続を切断
        Rs.Close()

        '入力制限期間を設定
        Dim sl As New CSetLimit
        sl.SetLimit()

        '-----------------------------------------------------
        ' 科目データ
        '-----------------------------------------------------
        fp.lblsyori.Text = "データロード中・・・科目コード"
        fp.lblsyori.Update()
        'frmProg.prgBar.Value = 2
        ErrMsg = "勘定科目データ取得中"

        '-----------------------------------------------------
        'フレックスグリッド見出し表示
        '↓---------------------------------------------------
        With frmInfo.fgKamoku
            .Rows = 1
            .FixedCols = 0
            .Cols = 2
            .ExtendLastCol = True
            .SelectionMode = flexSelectionByRow

            .FixedAlignment(0) = flexAlignCenterCenter
            .FixedAlignment(1) = flexAlignCenterCenter

            .ColWidth(0) = 600

            .TextMatrix(0, 0) = "コード"
            .TextMatrix(0, 1) = "勘定科目名"

            lCnt = 0

        End With
        '↑----------------------------------------------------------------------

        Rs = New ADODB.Recordset
        Rs.Open("SELECT sUcd,sNcd,sNm,tiIsTrk,tiIsZei FROM wkskm01 WHERE tiIsTrk = 1 ORDER BY sUcd", AdoData, , ADODB.LockTypeEnum.adLockReadOnly)
        Cnt = 0
        Do While Rs.EOF = False
            ReDim Preserve pblKamokuData(Cnt)
            pblKamokuData(Cnt).Code = Trim(Rs.Fields("sUcd").Value)
            pblKamokuData(Cnt).Ncd = Trim(Rs.Fields("sNcd").Value)
            pblKamokuData(Cnt).Name = Trim(Rs.Fields("sNm").Value)

            pblKamokuData(Cnt).IsZei = CStr(Rs.Fields("tiIsZei").Value)

            '----------------------------------------------------------------
            'フレックスグリッドに表示　
            '↓--------------------------------------------------------------
            With frmInfo.fgKamoku
                .Rows = .Rows + 1
                lCnt = lCnt + 1
                .TextMatrix(lCnt, 0) = pf.AddSpace(pblKamokuData(Cnt).Code, 4)
                .TextMatrix(lCnt, 1) = pblKamokuData(Cnt).Name
            End With
            '↑-----------------------------------------------------------------

            Cnt = Cnt + 1
            Rs.MoveNext()
        Loop
        Rs.Close()


        '-----------------------------------------------------
        ' 補助データ
        '-----------------------------------------------------
        fp.lblsyori.Text = "データロード中・・・補助科目コード"
        fp.lblsyori.Update()
        'frmProg.prgBar.Value = 4
        ErrMsg = "補助科目取得中"

        '------------------------------------------------------------------
        '   フレックスグリッド見出し表示　
        '↓----------------------------------------------------------------
        With frmInfo.fgHojo
            .Rows = 1
            .FixedCols = 0
            .Cols = 2
            .ExtendLastCol = True

            .FixedAlignment(0) = flexAlignCenterCenter
            .FixedAlignment(1) = flexAlignCenterCenter

            .ColWidth(0) = 600

            .TextMatrix(0, 0) = "コード"
            .TextMatrix(0, 1) = "補助科目名"

            lCnt = 0

        End With
        '↑----------------------------------------------------------------------

        Rs = New ADODB.Recordset
        Rs.Open("SELECT sHjoUcd,sNm,sSknNcd FROM wkhjm01 ORDER BY sSknNcd,sHjoUcd", AdoData, , ADODB.LockTypeEnum.adLockReadOnly)

        Cnt = 0
        cntKamoku = -1
        wrkBfNcd = ""

        Do While Rs.EOF = False
            '勘定科目内部コード取得
            wrkNcd = Trim(Rs.Fields("sSknNcd").Value)

            '勘定科目の内部コードが前回と異なっていた場合
            If wrkNcd <> wrkBfNcd Then
                'コード「0」が先頭の場合は、書き込みをスキップ
                If Val(Rs("sHjoUcd")) = 0 Then
                    GoTo Skip1
                End If

                'カウント初期化
                Cnt = 0
                cntKamoku = -1
                '勘定科目の検索
                For Loopcnt = 0 To UBound(pblKamokuData)
                    '科目コードがあったらOK
                    If (pblKamokuData(Loopcnt).Ncd = wrkNcd) Then
                        pblKamokuData(Loopcnt).HojoExist = True
                        cntKamoku = Loopcnt
                        Exit For
                    End If
                Next
                '万が一、勘定科目が見つからなかった場合は書込スキップ
                If cntKamoku = -1 Then
                    '-- 処理追加「前回番号に戻す」 (障害対応)-- 2004.03.26
                    wrkNcd = wrkBfNcd
                    GoTo Skip1
                End If
            End If

            '補助データの追加
            ReDim Preserve pblKamokuData(cntKamoku).HojoData(Cnt)
            pblKamokuData(cntKamoku).HojoData(Cnt).Code = Format(Val(Trim(Rs.Fields("sHjoUcd").Value)))
            pblKamokuData(cntKamoku).HojoData(Cnt).Name = Trim(Rs.Fields("sNm").Value)
            Cnt = Cnt + 1

Skip1:
            wrkBfNcd = wrkNcd
            Rs.MoveNext()
        Loop
        Rs.Close()


        '-----------------------------------------------------
        ' 部門データ
        '-----------------------------------------------------
        fp.lblsyori.Text = "データロード中・・・部門コード"
        fp.lblsyori.Update()
        'frmProg.prgBar.Value = 6
        ErrMsg = "部門データ取得中"

        '------------------------------------------------------------------
        '   フレックスグリッド見出し表示
        '↓----------------------------------------------------------------
        With frmInfo.fgBumon
            .Rows = 1
            .FixedCols = 0
            .Cols = 2
            .ExtendLastCol = True
            .SelectionMode = flexSelectionByRow

            .FixedAlignment(0) = flexAlignCenterCenter
            .FixedAlignment(1) = flexAlignCenterCenter

            .ColWidth(0) = 600

            .TextMatrix(0, 0) = "コード"
            .TextMatrix(0, 1) = "部門名"

            lCnt = 0

        End With
        '↑----------------------------------------------------------------------

        strOtherName = ""
        pblBumonFlg = False

        Rs = New ADODB.Recordset
        Rs.Open("SELECT sUcd,sNm FROM wkbnm01 ORDER BY sUcd", AdoData, , ADODB.LockTypeEnum.adLockReadOnly)
        Cnt = 0
        'pblBumonNum = 0
        ReDim Preserve pblBumonData(Cnt)

        Do While Rs.EOF = False
            'コードが「その他」以外
            If (Val(Trim(Rs.Fields("sUcd").Value)) <> 0) Then
                '部門フラグをON
                If pblBumonFlg = False Then
                    pblBumonFlg = True
                End If
                ReDim Preserve pblBumonData(Cnt)
                pblBumonData(Cnt).Code = Trim(Rs.Fields("sUcd").Value)
                pblBumonData(Cnt).Name = Trim(Rs.Fields("sNm").Value)

                '----------------------------------------------------------------
                '   部門フレックスグリッドに表示
                '↓--------------------------------------------------------------
                With frmInfo.fgBumon
                    .Rows = .Rows + 1
                    lCnt = lCnt + 1
                    .TextMatrix(lCnt, 0) = pf.AddSpace(pblBumonData(Cnt).Code, 4)
                    .TextMatrix(lCnt, 1) = pblBumonData(Cnt).Name
                End With
                '↑-----------------------------------------------------------------

                Cnt = Cnt + 1

            Else
                'コードが「0」の場合は、名称のみ取得
                strOtherName = Trim(Rs.Fields("sNm").Value)
            End If
            Rs.MoveNext()
        Loop
        Rs.Close()

        '部門ありなら、「その他」追加
        If pblBumonFlg = True Then
            ReDim Preserve pblBumonData(Cnt)
            pblBumonData(Cnt).Code = "0"
            pblBumonData(Cnt).Name = strOtherName

            '----------------------------------------------------------------
            '   部門フレックスグリッドに表示　
            '↓--------------------------------------------------------------
            With frmInfo.fgBumon
                .Rows = .Rows + 1
                lCnt = lCnt + 1
                .TextMatrix(lCnt, 0) = pf.AddSpace(pblBumonData(Cnt).Code, 4)
                .TextMatrix(lCnt, 1) = pblBumonData(Cnt).Name
            End With
            '↑-----------------------------------------------------------------

        End If

        '-----------------------------------------------------
        ' 税区分データ
        '-----------------------------------------------------
        fp.lblsyori.Text = "データロード中・・・税区分コード"
        fp.lblsyori.Update()
        'frmProg.prgBar.Value = 8
        ErrMsg = "税区分データ取得中"

        '------------------------------------------------------------------
        '   フレックスグリッド見出し表示
        '↓----------------------------------------------------------------
        With frmInfo.fgTax
            .Rows = 1
            .FixedCols = 0
            .Cols = 2
            .ExtendLastCol = True
            .SelectionMode = flexSelectionByRow

            .FixedAlignment(0) = flexAlignCenterCenter
            .FixedAlignment(1) = flexAlignCenterCenter

            .ColWidth(0) = 600

            .TextMatrix(0, 0) = "コード"
            .TextMatrix(0, 1) = "税区分名"

            lCnt = 0

        End With
        '↑-------------------------------------------------------------------

        Rs = New ADODB.Recordset
        Rs.Open("SELECT tiZeiCd,sZeiNm FROM wktax01 ORDER BY tiZeiCd", AdoData, , ADODB.LockTypeEnum.adLockReadOnly)
        Cnt = 0
        'pblTaxKbnNum = 0
        ReDim Preserve pblTaxKbnData(Cnt)
        Do While Rs.EOF = False
            'コードが「その他」以外
            ReDim Preserve pblTaxKbnData(Cnt)
            pblTaxKbnData(Cnt).Code = Trim(Rs.Fields("tiZeiCd").Value)
            pblTaxKbnData(Cnt).Name = Trim(Rs.Fields("sZeiNm").Value)

            '----------------------------------------------------------------
            '   フレックスグリッドに表示　2004/6/14 -soft bit k.yamagiwa-
            '↓--------------------------------------------------------------
            With frmInfo.fgTax
                .Rows = .Rows + 1
                lCnt = lCnt + 1
                .TextMatrix(lCnt, 0) = pf.AddSpace(pblTaxKbnData(Cnt).Code, 3)
                .TextMatrix(lCnt, 1) = pblTaxKbnData(Cnt).Name

            End With
            '↑-----------------------------------------------------------------

            Cnt = Cnt + 1
            Rs.MoveNext()
        Loop

        Rs.Close()

        '-----------------------------------------------------
        ' 税処理データ
        '-----------------------------------------------------
        fp.lblsyori.Text = "データロード中・・・税処理データ"
        fp.lblsyori.Update()
        'frmProg.prgBar.Value = 9
        ErrMsg = "その他データ"

        '------------------------------------------------------------------
        '   フレックスグリッド見出し表示
        '↓----------------------------------------------------------------
        With frmInfo.fgTaxMas
            .Rows = 1
            .FixedCols = 0
            .Cols = 2
            .ExtendLastCol = True
            .SelectionMode = flexSelectionByRow

            .FixedAlignment(0) = flexAlignCenterCenter
            .FixedAlignment(1) = flexAlignCenterCenter

            .ColWidth(0) = 600

            .TextMatrix(0, 0) = "コード"
            .TextMatrix(0, 1) = "税処理名"

            lCnt = 0

        End With
        '↑-------------------------------------------------------------------

        '消費税計算区分をセット
        pblTaxMasData(0).Code = "0"
        pblTaxMasData(0).Name = "税抜別段"
        pblTaxMasData(1).Code = "1"
        pblTaxMasData(1).Name = "税込自動"

        'リスト格納
        For Loopcnt = 0 To UBound(pblTaxMasData)

            '----------------------------------------------------------------
            '   フレックスグリッドに表示
            '↓--------------------------------------------------------------
            With frmInfo.fgTaxMas
                .Rows = .Rows + 1
                lCnt = lCnt + 1
                .TextMatrix(lCnt, 0) = pf.AddSpace(pblTaxMasData(Loopcnt).Code, 4)
                .TextMatrix(lCnt, 1) = pblTaxMasData(Loopcnt).Name

            End With
            '↑---------------------------------------------------------------

        Next

        '---------------------------------------------------------------------
        '       摘要データ　
        '---------------------------------------------------------------------
        fp.lblsyori.Text = "データロード中・・・摘要データ"
        fp.lblsyori.Update()
        'frmProg.prgBar.Value = 10
        ErrMsg = "摘要データ取得中"

        With frmInfo.fgTekiyo
            .Rows = 1
            .FixedCols = 0
            .Cols = 1
            .ExtendLastCol = True
            .SelectionMode = flexSelectionByRow
            .FixedAlignment(0) = flexAlignCenterCenter
            .TextMatrix(0, 0) = "摘要名"

            lCnt = 0

        End With

        Rs = New ADODB.Recordset
        Rs.Open("SELECT sUcd,sNm FROM wktkm01 ORDER BY sUcd", AdoData, , ADODB.LockTypeEnum.adLockReadOnly)
        Cnt = 0

        ReDim Preserve pblTekiMst(Cnt)
        Do While Rs.EOF = False

            ReDim Preserve pblTekiMst(Cnt)
            pblTekiMst(Cnt).Name = Trim(Rs.Fields("sNm").Value)

            With frmInfo.fgTekiyo
                .Rows = .Rows + 1
                lCnt = lCnt + 1
                .TextMatrix(lCnt, 0) = pblTekiMst(Cnt).Name

            End With

            Cnt = Cnt + 1
            Rs.MoveNext()
        Loop

        Rs.Close()
        AdoData.Close()
        AdoData = Nothing
        '↑--------------------------------------------------------------------------------

        fp.Close()
        On Error GoTo 0
        Exit Sub

ErrPrc:
        Dim em As CErr
        em.ErrMessage(ErrMsg)

    End Sub

    '----------------------------------------------------------------
    '   会社データ表示
    '----------------------------------------------------------------
    Public Sub ShowComData()
        Dim wrkGengou As String
        Dim wrkKikan As String
        Dim wrkFromYear As String
        Dim wrkFromMonth As String
        Dim wrkFromDay As String
        Dim wrkToYear As String
        Dim wrkToMonth As String
        Dim wrkToDay As String
        Dim wrkKaisi As String

        Dim pf As New PCfunc

        '会計期間のフォーマット
        If pblComData.Hosei <> "0" Then
            wrkGengou = pblComData.Gengou
            wrkFromYear = CStr(Val(pblComData.FromYear) - Val(pblComData.Hosei))
            wrkFromYear = pf.AddSpace(wrkFromYear, 2)
            wrkToYear = CStr(Val(pblComData.ToYear) - Val(pblComData.Hosei))
            wrkToYear = pf.AddSpace(wrkToYear, 2)
        Else
            wrkGengou = "  "
            wrkFromYear = pblComData.FromYear
            wrkToYear = pblComData.ToYear
        End If

        wrkFromMonth = pf.AddSpace(Val(pblComData.FromMonth), 2)
        wrkFromDay = pf.AddSpace(Val(pblComData.FromDay), 2)
        wrkToMonth = pf.AddSpace(Val(pblComData.ToMonth), 2)
        wrkToDay = pf.AddSpace(Val(pblComData.ToDay), 2)

        '入力開始月フォーマット
        wrkKaisi = CStr(Val(pblComData.FromMonth) + Val(pblComData.Kaisi))
        If Val(wrkKaisi) > 12 Then
            wrkKaisi = CStr(Val(wrkKaisi) - 12)
        End If

        '-- 取得方法追加「税処理を取得」 (v6.0対応)--
        If Trim(gsTaxMas) = "2" Then
            pblComData.TaxMas = "1"
        Else
            pblComData.TaxMas = "0"
        End If

        '---------------------------------
        '   フレックスグリッド見出し表示
        '↓-------------------------------
        With frmInfo.fgCom
            .Rows = 7
            .FixedCols = 1
            .Cols = 2
            .ExtendLastCol = True
            .SelectionMode = flexSelectionByRow

            .FixedAlignment(0) = flexAlignCenterCenter
            .FixedAlignment(1) = flexAlignCenterCenter

            .ColWidth(0) = 1500

            .TextMatrix(0, 0) = "項目名"
            .TextMatrix(0, 1) = "摘要"

            'lCnt = 0

            '会社名
            .TextMatrix(1, 0) = "会社名" : .TextMatrix(1, 1) = pblComData.Name

            '会計期間期首
            .TextMatrix(2, 0) = "会計期間・期首" : .TextMatrix(2, 1) = wrkFromYear & "年" & wrkFromMonth & "月" & wrkFromDay & "日"

            '会計期間期末
            .TextMatrix(3, 0) = "会計期間・期末" : .TextMatrix(3, 1) = wrkToYear & "年" & wrkToMonth & "月" & wrkToDay & "日"

            '入力開始月
            .TextMatrix(4, 0) = "入力開始月" : .TextMatrix(4, 1) = pf.AddSpace(Val(wrkKaisi), 2) & "月"

            '中間期決算
            .TextMatrix(5, 0) = "決算回数"
            If pblComData.Middle = FLGON Then
                .TextMatrix(5, 1) = "する"
            Else
                .TextMatrix(5, 1) = "しない"
            End If

            '決算回数 2004/6/28 -softbit k.yamagiwa-
            Select Case pblComData.Middle
                Case 0
                    .TextMatrix(5, 1) = "年1回"
                Case 1
                    .TextMatrix(5, 1) = "年2回（中間決算）"
                Case 2
                    .TextMatrix(5, 1) = "年4回（四半期決算）"
                Case Else
                    .TextMatrix(5, 1) = "不明"
            End Select

            '税処理
            .TextMatrix(6, 0) = "税処理"
            If pblComData.TaxMas = "0" Then
                .TextMatrix(6, 1) = "税抜別段"
            Else
                .TextMatrix(6, 1) = "税込自動"
            End If

        End With
        '↑-------------------------------------------------------------------

    End Sub
End Class
