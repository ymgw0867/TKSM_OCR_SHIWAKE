Imports 仕訳伝票.PCData


Public Class CErr
    '----------------------------------------------------------------
    '   エラー時の処理
    '----------------------------------------------------------------
    Public Sub ErrMessage(ByVal Msg As String)
        Dim pf As New PCfunc

        Call pf.FileDelete(pblInstPath & DIR_HENKAN & TMPREAD)
        Call pf.FileDelete(pblInstPath & DIR_HENKAN & tmpFile)
        MessageBox.Show(Msg & "にエラーが発生したため、処理を終了します。", Title)
        'frmProg.Close()
        End

    End Sub

End Class
