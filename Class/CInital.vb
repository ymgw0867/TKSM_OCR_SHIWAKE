Imports 固定項目保守.PCData
Imports 固定項目保守.PCfunc
Imports 固定項目保守.frmComSelect

Public Class CInital

    '----------------------------------------------------------------
    '   初期化
    '----------------------------------------------------------------
    Public Sub Initial()
        On Error GoTo ErrPrc
        '---------------------------------------------
        ' インスタンス実行中チェック
        '---------------------------------------------
        If UBound(Diagnostics.Process.GetProcessesByName(Diagnostics.Process.GetCurrentProcess.ProcessName)) > 0 Then
            MsgBox("このアプリケーションはすでに実行されています。", vbOKOnly Or vbExclamation, "ＮＧ")
            End
        End If

        '---------------------------------------------
        ' 作業ディレクトリ情報取得
        '---------------------------------------------
        Dim Getp As New PCfunc
        pblInstPath = Getp.GetPath()

        '----------------------------------------------
        ' 設定データ取得
        '----------------------------------------------
        Dim IniLoad As New CInitialLoad
        IniLoad.InitialLoad()

        '----------------------------------------------
        ' データベース名取得（会社選択）
        '----------------------------------------------
        pblDbName = GetDBName()

        On Error GoTo 0

        Exit Sub

        ' エラー処理
ErrPrc:
        Dim errm As New PCfunc
        errm.ErrMessage("初期設定中")

    End Sub

    '----------------------------------------------------------------
    '   データベース名取得
    '----------------------------------------------------------------
    Public Function GetDBName()
        Dim wrkRetValue As String

        '会社選択フォームロード
        Dim fcSel As New frmComSelect
        fcSel.ShowDialog()

        '会社選択画面表示
        'fcSel.Visible = True

        'データベース名取得
        wrkRetValue = frmComSelect.fcRetValue

        'フォームアンロード
        fcSel.Dispose()

        'データベース名
        GetDBName = wrkRetValue

    End Function

End Class
