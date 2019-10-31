Imports 仕訳伝票.PCData
Imports 仕訳伝票.frmProg
Imports 仕訳伝票.frmComSelect

Public Class CInitial
    '----------------------------------------------------------------
    '   初期化
    '----------------------------------------------------------------
    Public Sub Initial()
        On Error GoTo ErrPrc

        Dim fp As New frmProg

        '処理中表示
        fp.Show()

        '---------------------------------------------
        ' 二重実行チェック
        '---------------------------------------------
        If UBound(Diagnostics.Process.GetProcessesByName(Diagnostics.Process.GetCurrentProcess.ProcessName)) > 0 Then
            fp.Close()
            MessageBox.Show("このアプリケーションはすでに実行されています。", "ＮＧ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End
        End If

        '---------------------------------------------
        ' 作業ディレクトリ情報取得
        '---------------------------------------------
        Dim Getp As New PCfunc
        pblInstPath = Getp.GetPath()

        '---------------------------------------------
        ' フォルダ作成
        '---------------------------------------------
        On Error Resume Next
        MkDir(pblInstPath & DIR_OK)
        MkDir(pblInstPath & DIR_INCSV)     '分割したＣＳＶの格納フォルダ 2004/6/24
        MkDir(pblInstPath & DIR_BREAK)     '中断伝票フォルダ 2004/6/24
        On Error GoTo 0

        On Error GoTo ErrPrc

        '----------------------------------------------
        ' ファイル有無チェック
        '----------------------------------------------
        Call FileExistChk()

        If (pblSelFILE <> 1) And (pblFlgINFILE = False) And (pblFlgDIVFILE = False) Then
            fp.Close()
            MessageBox.Show("処理を行うデータがありません。", "ＮＧ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End
        End If

        '----------------------------------------------
        ' 設定データ取得
        '----------------------------------------------
        Dim il As New CInitialLoad
        il.InitialLoad()

        '----------------------------------------------
        ' データベース名取得（会社選択）
        '----------------------------------------------
        pblDbName = GetDBName()

        '----------------------------------------------
        ' 結果ファイル削除
        '----------------------------------------------
        Dim cf As New PCfunc
        cf.FileDelete(pblInstPath & DIR_OK & OUTFILE)

        '-----------------------------------------
        ' 入力ファイルコピー(INFILE→TEMPREAD)
        '-----------------------------------------
        If pblFlgINFILE = True Then
            FileCopy(pblInstPath & DIR_HENKAN & INFILE, pblInstPath & DIR_HENKAN & TMPREAD)

            '-------------------------------------
            ' 入力ファイルチェック、設定
            '-------------------------------------
            Call DenKindJudge()

        End If

        '処理中表示終了
        fp.Close()

        On Error GoTo 0
        Exit Sub

        ' エラー処理
ErrPrc:
        Dim em As New CErr
        em.ErrMessage("初期設定中")

    End Sub

    '----------------------------------------------------------------
    '   ファイル有無チェック
    '----------------------------------------------------------------
    Private Sub FileExistChk()
        Dim db As New ADODB.Connection   'ADOのConnectionオブジェクトを生成
        Dim rs As New ADODB.Recordset
        Dim ssql As String

        '-------------------------------------------------------
        '   設定データベース
        '-------------------------------------------------------
        If (Dir(pblInstPath & DIR_CONFIG & CONFIGFILE) = "") Then
            MessageBox.Show("設定データベースがありません。" & vbNewLine & "ソフトを再インストールしてください。", "ＮＧ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End
        End If

        Dim cf As New PCfunc
        pblInstPath = cf.GetPath()

        ' データベースを開く
        db.Open(MDBCONNECT & pblInstPath & DIR_HENKAN & CONFIGFILE & ";")

        'SQL文のみを入力
        ssql = "SELECT * FROM Config"

        rs = db.Execute(ssql)

        'メニューで指定した伝票の区分
        pblSelFILE = Trim(rs.Fields("sub2").Value)

        Rs.Close()
        db.Close()

        Rs = Nothing
        db = Nothing

        '入力ファイル
        pblFlgINFILE = True

        If (Dir(pblInstPath & DIR_HENKAN & INFILE) = "") Then
            pblFlgINFILE = False
        Else
            '入力ファイルがある場合、分割ファイルを削除する
            Dim pf As New PCfunc
            pf.FileDelete(pblInstPath & DIR_INCSV & "*.*")
            pblFlgDIVFILE = False
            Exit Sub
        End If

        '分割ファイル
        pblFlgDIVFILE = True

        If (Dir(pblInstPath & DIR_INCSV & "*.csv") = "") Then
            pblFlgDIVFILE = False

        End If

    End Sub

    '----------------------------------------------------------------
    '   入力ファイルチェック、設定
    '----------------------------------------------------------------
    Private Sub DenKindJudge()
        Dim sReadBuf As String               '1行読込みバッファ

        '-----------------
        ' ファイルオープン
        '-----------------
        Dim fileReader = My.Computer.FileSystem.OpenTextFileReader(pblInstPath & DIR_HENKAN & TMPREAD)
        If fileReader.EndOfStream = True Then
            fileReader.Close()
            Call InFileErr()
            Exit Sub
        End If

        '１行読み込み
        sReadBuf = fileReader.ReadLine()

        '先頭に「*」以外ならエラー伝票
        If (Strings.Left(sReadBuf, 1) <> "*") Then
            fileReader.Close()
            Call InFileErr()
        End If

        fileReader.Close()

    End Sub

    Private Sub InFileErr()
        Dim cf As New PCfunc
        'Dim fp As New frmProg

        cf.FileDelete(pblInstPath & DIR_HENKAN & TMPREAD)
        FileCopy(pblInstPath & DIR_HENKAN & INFILE, pblInstPath & DIR_HENKAN & LOGFILE)
        MessageBox.Show("不正な伝票ファイルです。伝票データが存在しません。", "ＮＧ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

        'fp.Dispose()

    End Sub

    '----------------------------------------------------------------
    '   データベース名取得
    '----------------------------------------------------------------
    Public Function GetDBName()
        Dim fc As New frmComSelect
        Dim wrkRetValue As String

        '会社選択フォームロード
        fc.ShowDialog()

        'データベース名（今回は会社名）取得
        wrkRetValue = frmComSelect.fcRetValue

        'フォームアンロード
        fc.Close()

        'データベース名
        GetDBName = wrkRetValue

    End Function

End Class
