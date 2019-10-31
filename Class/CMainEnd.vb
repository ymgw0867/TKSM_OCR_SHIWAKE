Imports 仕訳伝票.PCData
Imports 仕訳伝票.frmInfo


Public Class CMainEnd

    '----------------------------------------------------------------
    '   終了処理
    '----------------------------------------------------------------
    Public Sub MainEnd()
        'Dim fp As New frmProg
        'fp.Show()

        'fp.lblsyori.Text = "データ変換中・・・"
        'frmProg.prgBar.Value = 60

        'データ変換
        Dim csd As New CSaveData
        csd.SaveData()

        'frmProg.prgBar.Value = 100

        'fp.Dispose()

        '分割ファイル削除
        Dim pf As New PCfunc
        pf.FileDelete(pblInstPath & DIR_INCSV & "*.*")

        '終了メッセージ表示
        If (pblDenNum > 0) Then
            Call EndMsg()
        End If

        'Unload frmInfo

        End
    End Sub


    '----------------------------------------------------------------
    '   終了メッセージ
    '----------------------------------------------------------------
    Private Sub EndMsg()
        MessageBox.Show("処理が終了しました。" & vbCrLf & "勘定奉行でデータの受入れを行ってください。", Title)
    End Sub


End Class
