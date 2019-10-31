Imports 仕訳伝票.PCData
Imports 仕訳伝票.frmProg

Public Class CChkMain
    '--------------------
    '   チェックメイン 
    '--------------------
    Public Function ChkMain() As Boolean
        Dim ret As String
        Dim wrkErrMsg As String
        Dim chk As Integer
        Dim Cnt As Integer

        '処理中表示
        Dim fp As New frmProg
        fp.Show()

        ret = True

        pblNowden = 1
        pblNowGyou = 1
        pblNowErrNo = 0
        pblNowErrNum = 0

        'エラーテーブルクリア
        pblErrCnt = 0

        '差額エラー発生フラグ
        pblSagakuFLG = False

        On Error GoTo ErrPrc

        For Cnt = 1 To pblDenNum
            System.Windows.Forms.Application.DoEvents()

            '-------------------
            ' 結合チェック
            ' -------------------
            fp.lblsyori.Text = "ＮＧチェック中・・・結合"
            'frmProg.prgBar.Value = 20
            wrkErrMsg = "結合チェック中"

            '結合枚数チェック
            ret = ChkCombine(Cnt)
            DoEvents()

            '結合行数チェック
            ret = ChkCombineItem(Cnt)
            DoEvents()

            '結合日付チェック
            ret = ChkCombineDate(Cnt)
            DoEvents()

            '結合伝票No.チェック
            ret = ChkCombineDenNo(Cnt)
            DoEvents()

            '結合決算チェック
            ret = ChkCombineKessan(Cnt)
            DoEvents()

            '----------------------
            ' 日付チェック
            '----------------------
            fp.lblsyori.Text = "ＮＧチェック中・・・日付"
            'frmProg.prgBar.Value = 30
            wrkErrMsg = "日付チェック中"

            ret = ChkDate(Cnt)
            DoEvents()
            If ret = True Then
                '----------------------
                ' 決算日付チェック
                '----------------------
                fp.lblsyori.Text = "ＮＧチェック中・・・決算日付"
                'frmProg.prgBar.Value = 34
                wrkErrMsg = "決算日付チェック中"
                ret = ChkDateKessan(Cnt)
                System.Windows.Forms.Application.DoEvents()
            End If
            '----------------------
            ' 会計期間チェック
            '----------------------
            If ret = True Then
                fp.lblsyori.Text = "ＮＧチェック中・・・会計期間"
                'frmProg.prgBar.Value = 38
                wrkErrMsg = "会計期間チェック中"
                ret = ChkDateKikan(Cnt)
                System.Windows.Forms.Application.DoEvents()
            End If
            '----------------------
            ' 日付入力範囲チェック
            '----------------------
            If ret = True Then
                fp.lblsyori.Text = "ＮＧチェック中・・・日付入力範囲"
                'frmProg.prgBar.Value = 42
                wrkErrMsg = "日付入力範囲チェック中"
                ret = ChkDateLimit(Cnt)
                System.Windows.Forms.Application.DoEvents()
            End If

            '----------------------
            ' 入力不備チェック
            '----------------------
            fp.lblsyori.Text = "ＮＧチェック中・・・入力不備"
            'frmProg.prgBar.Value = 46
            wrkErrMsg = "入力不備チェック中"
            ret = ChkDataPoor(Cnt)
            System.Windows.Forms.Application.DoEvents()

            '----------------------
            ' 勘定科目コードチェック
            '----------------------
            fp.lblsyori.Text = "ＮＧチェック中・・・勘定科目コード"
            'frmProg.prgBar.Value = 50
            wrkErrMsg = "勘定科目コードチェック中"
            ret = ChkKamoku(Cnt)
            System.Windows.Forms.Application.DoEvents()

            '-------------------
            ' 補助コードチェック
            '-------------------
            fp.lblsyori.Text = "ＮＧチェック中・・・補助科目コード"
            'frmProg.prgBar.Value = 60
            wrkErrMsg = "補助科目コードチェック中"
            ret = ChkHojo(Cnt)
            System.Windows.Forms.Application.DoEvents()

            '-------------------
            ' 部門コードチェック
            '-------------------
            fp.lblsyori.Text = "ＮＧチェック中・・・部門コード"
            'frmProg.prgBar.Value = 70
            wrkErrMsg = "部門コードチェック中"
            ret = ChkBumon(Cnt)
            System.Windows.Forms.Application.DoEvents()

            '-----------------------------------------------------
            ' 消費税計算区分（略名：税処理）コードチェック
            '-----------------------------------------------------
            fp.lblsyori.Text = "ＮＧチェック中・・・税処理コード"
            'frmProg.prgBar.Value = 80
            wrkErrMsg = "税処理コードチェック中"
            ret = ChkOther(Cnt)
            System.Windows.Forms.Application.DoEvents()

            '-------------------
            ' 税区分コードチェック
            '-------------------
            fp.lblsyori.Text = "ＮＧチェック中・・・税区分コード"
            'frmProg.prgBar.Value = 85
            wrkErrMsg = "税区分コードチェック中"
            ret = ChkTaxKbn(Cnt)
            System.Windows.Forms.Application.DoEvents()

            '-------------------
            ' 貸借差額チェック
            '-------------------
            fp.lblsyori.Text = "ＮＧチェック中・・・貸借差額"
            'frmProg.prgBar.Value = 90
            wrkErrMsg = "貸借差額チェック中"
            ret = ChkSum(Cnt)

            '-------------------------
            '相手科目未記入チェック
            '-------------------------
            fp.lblsyori.Text = "ＮＧチェック中・・・相手科目"
            'frmProg.prgBar.Value = 95
            wrkErrMsg = "相手科目未記入チェック中"
            ret = ChkAite(Cnt)
            System.Windows.Forms.Application.DoEvents()

            '----------------------
            ' 摘要複写
            '----------------------
            fp.lblsyori.Text = "ＮＧチェック中・・・摘要文字数"
            'frmProg.prgBar.Value = 96
            wrkErrMsg = "摘要文字数チェック中"
            ret = ChkTekiyou(Cnt)
            System.Windows.Forms.Application.DoEvents()

            '----------------------
            ' 有効明細
            '----------------------
            fp.lblsyori.Text = "ＮＧチェック中・・・有効明細"
            'frmProg.prgBar.Value = 98
            wrkErrMsg = "有効明細チェック中"
            ret = ChkYukoMeisai(Cnt)
            System.Windows.Forms.Application.DoEvents()

            '日付摘要のみの場合は、エラーとする
            '----------------------
            ' 有効明細
            '----------------------
            fp.lblsyori.Text = "ＮＧチェック中・・・摘要のみ"
            'frmProg.prgBar.Value = 99
            wrkErrMsg = "摘要のみチェック中"
            ret = ChkTekiyouOnly(Cnt)
            System.Windows.Forms.Application.DoEvents()

            'NEXTDEN:
        Next Cnt

        'frmProg.prgBar.Value = 100
        On Error GoTo 0

        chk = pblDenNum
        ChkMainNEW = ret

        Exit Function

        ' エラー処理
ErrPrc:
        Dim em As New CErr
        em.ErrMessage(wrkErrMsg)

    End Function


End Class
