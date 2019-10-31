Imports 仕訳伝票.PCData

Public Class CMain
    Public Shared finfo As New frmInfo

    Shared Sub Main()
        '初期設定
        Dim Cini As New CInitial
        Cini.Initial()

        'マスターデータロード
        Dim lm As New CLoadMaster
        lm.LoadMaster()

        '処理するファイルの読込
        Dim lc As New CLoadCSV
        Select Case pblSelFILE

            Case 0
                '伝票データ分割(TMPREAD→CSVFILE)
                lc.LoadCSV()

            Case 1
                '中断データリカバリー
                '処理するファイルを選択
                Dim fs As frmFilSelect
                fs.ShowDialog()
                fs.Close()
                lc.LoadChudan()

        End Select

        lc = Nothing

        '伝票データロード(分割CSV→pblDenRec)
        Call LoadData()

        'エラーチェックループ
        Call ChkLoop()

    End Sub

    '----------------------------------------------------------------
    '   メインループ処理
    '----------------------------------------------------------------
    Public Sub ChkLoop()
        Dim ret As Boolean
        Dim retYesNo As String
        Dim Cnt As Integer
        Dim ans As MsgBoxResult

        frmInfo.Enabled = False

        '処理中表示
        'Dim fp As New frmProg
        'fp.Show()
        'fp.lblsyori.Text = "NGチェック中・・・"
        'frmProg.prgBar.Value = 20

        '-------------------
        ' データチェック
        '-------------------
        ret = ChkMain

        'チェックNGのとき
        If pblErrCnt > 0 Then
            '処理中終了
            'fp.Dispose()

            frmInfo.Enabled = True
            '----------------
            ' ダイアログセット
            '----------------
            pblNowden = pblErrTBL(1).ErrDenNo
            Call ShowData()
            Beep()
        Else
            'チェックOKのとき
            'エラー表示あり後、エラー修正してNGがない場合、ワーニングカラーが残っているのを回避
            Call ShowData()
            ans = MessageBox.Show("ＮＧは見つかりませんでした。処理を終了しますか？", Title, MessageBoxIcon.Question + MessageBoxButtons.YesNo, MessageBoxDefaultButton.Button2)
            'No
            If (ans = MsgBoxResult.No) Then
                '編集画面表示
                '処理中終了
                'fp.Dispose()
                frmInfo.Enabled = True
                'Yes
            Else
                ' 終了処理コール
                Call MainEnd()
            End If
        End If
    End Sub

End Class
