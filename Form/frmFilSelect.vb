
Public Class frmFilSelect

    Private Sub frmFilSelect_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed

    End Sub

    Private Sub frmFilSelect_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim ans As MsgBoxResult

        '×ボタン押下時
        If (UnloadMode = vbFormControlMenu) Or Me.Tag = "btnNo" Then
            ans = MessageBox.Show("変換プログラムを終了します。よろしいですか？", "終了", MessageBoxButtons.YesNo + MessageBoxDefaultButton.Button2)
            Me.Tag = ""
            If ans = MsgBoxResult.Yes Then
                Dim cr As New CErr
                cr.ErrEnd()
            Else
                e.Cancel = True
                Exit Sub
            End If
        End If
    End Sub

    Private Sub frmFilSelect_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim sFileName As String
        Dim gFileName As String
        Dim inTmp As Integer
        Dim fuCHK As String
        Dim Cnt As Integer
        Dim hcnt As Integer
        Dim readbuf As String

        '中断伝票をリスト表示
        CList1.Items.Clear()
        X = 0
        Y = 0

        readbuf = ""
        If Dir(pblInstPath & DIR_BREAK & Format(pblComNo, "000") & "\*.csv") <> "" Then
            sFileName = Dir(pblInstPath & DIR_BREAK & Format(pblComNo, "000") & "\*.csv")
            Do While sFileName <> ""
                '日付時間を取得
                gFileName = Left(Right(sFileName, 25), 10)

                '異なる日付時間なら値を保持
                If (readbuf <> "") And (readbuf <> gFileName) Then
                    Call RecovSet(readbuf)
                End If
                Y = Y + 1
                readbuf = gFileName
                '次のファイル
                sFileName = Dir()
            Loop
            Call RecovSet(readbuf)
        End If

        '--------------------------------------
        '   作業中断情報リスト表示
        '--------------------------------------
        For i = 1 To UBound(pblRecov)
            List1.AddItem(Left(pblRecov(i).recName, 2) & "月" & Mid(pblRecov(i).recName, 3, 2) & "日" & _
                          Mid(pblRecov(i).recName, 5, 2) & "時" & Mid(pblRecov(i).recName, 7, 2) & "分" & _
                          Mid(pblRecov(i).recName, 9, 2) & "秒 : " & pblRecov(i).recFuku & " 件")
        Next i

        Exit Sub
        On Error GoTo 0

        'エラー処理
ErrPrc:
        Dim em As New CErr
        em.ErrMessage("ファイル選択中")

    End Sub

    Private Sub CList_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CList1.SelectedIndexChanged
        Dim Cnt As Integer
        Dim FLUG As Integer

        On Error GoTo ErrPrc

        '中断伝票チェック
        X = Item + 1

        If List1.Selected(Item) = True Then
            pblRecov(X).recFlg = 1
        Else
            pblRecov(X).recFlg = 0
        End If

        FLUG = pblRecov(X).recFlg


        Exit Sub
        On Error GoTo 0

        'エラー処理
ErrPrc:
        Dim em As New CErr
        em.ErrMessage("ファイル選択中")

    End Sub

    Private Sub CmdEnd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdEnd.Click

        '終了ボタン押下時
        frmFilSelect.Tag = "btnNo"
        frmFilSelect.Close()

    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        '-------------------------------------------
        '   ファイル選択確認
        '-------------------------------------------
        Dim ans As MsgBoxResult
        Dim strMSG As String

        strMSG = "中断処理データを読み込みます。よろしいですか？"
        ans = MessageBox.Show(strMSG,, "確認" MessageBoxIcon.Question + MessageBoxButtons.YesNo)
        If ans = MsgBoxResult.No Then
            Exit Sub
        End If

        '中断ファイルのリカバリーファイル名の取得
        If SelRecovery(pblSelFILE) = False Then
            MessageBox.Show("伝票が選択されていません。", "中断伝票未選択", MessageBoxIcon.Exclamation + MessageBoxButtons.OK)
            Exit Sub
        End If

        frmFilSelect.Close()

    End Sub

    Private Sub cmdTrue_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTrue.Click
        '---------------------------------
        '   全ての伝票を選択する
        '---------------------------------
        For i = 0 To CList1.ListCount - 1
            CList1.Selected(i) = True
        Next i

        CList1.ListIndex = 0

        cmdOK.Enabled = True

    End Sub

    Private Sub cmdFalse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFalse.Click
        '---------------------------------
        '   全ての伝票を選択する
        '---------------------------------
        For i = 0 To CList1.ListCount - 1
            CList1.Selected(i) = False
        Next i

        CList1.ListIndex = 0

    End Sub
End Class
