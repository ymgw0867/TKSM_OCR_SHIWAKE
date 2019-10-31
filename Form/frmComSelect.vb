Imports Microsoft.VisualBasic.Strings
Imports 仕訳伝票.PCData
Imports 仕訳伝票.PCfunc


Public Class frmComSelect
    Dim wrkComData() As strComDBData
    Public Shared fcRetValue As String
    Public Shared TaxMas As String

    Private Sub frmComSelect_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim AdoData As New ADODB.Connection   'ADOのConnectionオブジェクトを生成
        Dim Rs As ADODB.Recordset             'Recordsetオブジェクトの生成
        Dim pf As New PCfunc

        Dim Cnt As Integer
        Dim wrkBfComIndex As Integer
        Dim wrkConnectInfo As String
        Dim wrkComSelect As String
        Dim wrkComNo As String
        Dim wrkComName As String
        Dim wrkFromYear As String
        Dim wrkFromMonth As String
        Dim wrkFromDay As String
        Dim wrkKessanKi As String

        With Me.gvComp
            .ColumnCount = 4
            .Columns(0).Name = "№"
            .Columns(1).Name = "期首"
            .Columns(2).Name = "決算期"
            .Columns(3).Name = "会社名"
            .Columns(0).Width = 50
            .Columns(1).Width = 100
            .Columns(2).Width = 100
            .Columns(3).Width = 285

            .MultiSelect = False
            .RowHeadersVisible = False
        End With
        wrkConnectInfo = pf.fncGetConnect(pblDsnPath)
        Dim bstr() As Byte

        bstr = System.Text.Encoding.GetEncoding(932).GetBytes(wrkConnectInfo)
        wrkConnectInfo = System.Text.Encoding.UTF8.GetString(bstr)

        'DBを開く
        AdoData.ConnectionString = wrkConnectInfo
        AdoData.Open()

        ' 会社情報を開く
        Rs = New ADODB.Recordset
        Rs.Open("SELECT sDbNm,siCorpNo,sDateKisyu,sHosei,siKsnKi,sCorpNm FROM wcompany", AdoData, , ADODB.LockTypeEnum.adLockReadOnly)

        Cnt = 0
        Do While Rs.EOF = False

            '中断データ処理のときは該当する会社しか表示しない 2004/7/21
            If pblSelFILE = 1 Then

                Me.Text = "中断伝票処理　会社選択"
                pblBfDbName = ""

                If Dir(pblInstPath & DIR_BREAK & Format(Rs.Fields("siCorpNo").Value, "000") & "\", vbDirectory) <> "" Then
                    ReDim Preserve wrkComData(Cnt)
                    With wrkComData(Cnt)
                        .DbName = Trim(Rs.Fields("sDbNm").Value)
                        .Name = Trim(Rs.Fields("sCorpNm").Value)
                        .ComNo = Trim(CStr(Rs.Fields("siCorpNo").Value))
                        .FromYear = Mid(Trim(Rs.Fields("sDateKisyu").Value), 1, 4)
                        .FromMonth = Mid(Trim(Rs.Fields("sDateKisyu").Value), 5, 2)
                        .FromDay = Strings.Right(Trim(Rs.Fields("sDateKisyu").Value), 2)
                        .Hosei = Trim(Rs.Fields("sHosei").Value)
                        .KessanKi = Trim(CStr(Rs.Fields("siKsnKi").Value))
                    End With

                    wrkComData(Cnt).TaxMas = GetTaxMas(wrkComData(Cnt).DbName)

                    Cnt = Cnt + 1
                End If

            Else
                ReDim Preserve wrkComData(Cnt)
                With wrkComData(Cnt)
                    .DbName = Trim(Rs.Fields("sDbNm").Value)
                    .Name = Trim(Rs.Fields("sCorpNm").Value)
                    .ComNo = Trim(CStr(Rs.Fields("siCorpNo").Value))
                    .FromYear = Mid(Trim(Rs.Fields("sDateKisyu").Value), 1, 4)
                    .FromMonth = Mid(Trim(Rs.Fields("sDateKisyu").Value), 5, 2)
                    .FromDay = Strings.Right(Trim(Rs.Fields("sDateKisyu").Value), 2)
                    .Hosei = Trim(Rs.Fields("sHosei").Value)
                    .KessanKi = Trim(CStr(Rs.Fields("siKsnKi").Value))
                End With

                wrkComData(Cnt).TaxMas = GetTaxMas(wrkComData(Cnt).DbName)

                Cnt = Cnt + 1
            End If

            Rs.MoveNext()
        Loop

        '接続を切断
        Rs.Close()
        AdoData.Close()
        AdoData = Nothing

        '------------------------------
        '   グリッド明細表示
        '------------------------------
        For Cnt = 0 To UBound(wrkComData)

            If wrkComData(Cnt).TaxMas <> "" Then

                '会社No
                wrkComNo = pf.AddSpace(CStr(Val(wrkComData(Cnt).ComNo)), 4)

                '会計期間のフォーマット
                If wrkComData(Cnt).Hosei <> "0" Then
                    wrkFromYear = CStr(Val(wrkComData(Cnt).FromYear) - Val(wrkComData(Cnt).Hosei))
                    wrkFromYear = pf.AddSpace(wrkFromYear, 2)
                Else
                    wrkFromYear = Strings.Right(wrkComData(Cnt).FromYear, 2)
                End If

                wrkFromMonth = pf.AddSpace(CStr(Val(wrkComData(Cnt).FromMonth)), 2)
                wrkFromDay = pf.AddSpace(CStr(Val(wrkComData(Cnt).FromDay)), 2)

                '決算期
                wrkKessanKi = pf.AddSpace(CStr(Val(wrkComData(Cnt).KessanKi)), 3)

                '会社名
                wrkComName = pf.AddBackSpace(wrkComData(Cnt).Name, 60)

                '明細表示
                'If Cnt = 0 Then
                '    gvComp.Rows. = {wrkComNo, wrkFromYear & "/" & wrkFromMonth & "/" & wrkFromDay, "第" & Trim(wrkKessanKi) & "期", wrkComName}
                'Else
                gvComp.Rows.Add({wrkComNo, wrkFromYear & "/" & wrkFromMonth & "/" & wrkFromDay, "第" & Trim(wrkKessanKi) & "期", wrkComName})

                'End If
            End If

        Next Cnt


        '   -------------------------------------------------------------------------
        '
        If gvComp.RowCount < 2 Then
            'メッセージ
            MessageBox.Show("会社情報が存在しません。", Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End
        End If

        ' 及び前回選択した会社を検索
        If pblBfDbName = "" Then
            wrkBfComIndex = 0
        Else
            wrkBfComIndex = -1

            For Cnt = 0 To UBound(wrkComData)
                If wrkComData(Cnt).DbName = pblBfDbName Then
                    wrkBfComIndex = Cnt
                    Exit For
                End If
            Next

            ' もし会社が見つからなければIndex=0
            If wrkBfComIndex = -1 Then
                wrkBfComIndex = 0
            End If
        End If

        ' 画面表示
        Call ShowTaxMas(wrkBfComIndex)

        '最終行に追加の空白行を削除するために追加を不可とする
        gvComp.AllowUserToAddRows = False

        'フレックスグリッドに変更
        gvComp.Rows(wrkBfComIndex).Selected = True
        gvComp.Select()

        On Error GoTo 0
        Exit Sub

ErrPrc:
        ' エラー処理
        MessageBox.Show("勘定奉行のデータベースに接続出来ません",  "接続エラー")
        Call pf.ErrMessage("会社情報取得中")
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles gvComp.CellContentClick
        Me.btnOk.Select()
    End Sub

    '----------------------------------------------------------------
    '   各社の消費税計算区分を取得
    '----------------------------------------------------------------
    Private Function GetTaxMas(ByVal DbName As String) As String
        Dim AdoData As New ADODB.Connection   'ADOのConnectionオブジェクトを生成
        Dim Rs As ADODB.Recordset             'Recordsetオブジェクトの生成
        Dim RetValue As String
        Dim wrkConnectInfo As String

        On Error GoTo ErrProc

        ' データベースを開く
        Dim gc As New PCfunc
        wrkConnectInfo = gc.GetDbConnect(DbName)

        AdoData.ConnectionString = wrkConnectInfo
        AdoData.Open()

        ' 消費税情報テーブルを開く
        Rs = New ADODB.Recordset
        Rs.Open("SELECT tiIsZei FROM wktax03", AdoData, , ADODB.LockTypeEnum.adLockReadOnly)

        ' 消費税計算区分の取得
        RetValue = CStr(Trim(Rs.Fields("tiIsZei").Value))

        ' 接続を切断
        Rs.Close()
        AdoData.Close()
        AdoData = Nothing

        GetTaxMas = RetValue

        On Error GoTo 0

        Exit Function

ErrProc:
        'エラー時は空文字列を返す
        GetTaxMas = ""

    End Function

    Private Sub ShowTaxMas(ByVal Index As Integer)
        ' 税抜別段
        If wrkComData(Index).TaxMas = "0" Then
            '-- 表示名変更「会社を選択してください。」 (v6.0対応)--
            lblmsg.Text = "会社を選択してください。"

            ' 税込自動
        ElseIf wrkComData(Index).TaxMas = "2" Then
            lblmsg.Text = "会社を選択してください。"

            ' 税抜自動
        Else
            lblmsg.Text = "会社及び税処理を選択してください。"
        End If

    End Sub

    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
        Dim AdoData As New ADODB.Connection   'ADOのConnectionオブジェクトを生成
        Dim sSql As String

        'フレックスグリッドに変更
        If (gvComp.RowCount < 1) Then
            MessageBox.Show("会社を選択してください。", "ＮＧ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        'フレックスグリッドに変更
        fcRetValue = wrkComData(gvComp.SelectedRows(0).Index).DbName

        pblComNo = wrkComData(gvComp.SelectedRows(0).Index).ComNo

        'フレックスグリッドに変更
        gsTaxMas = GetTaxMas(wrkComData(gvComp.SelectedRows(0).Index).DbName)

        ' 選択されたデータベースを記録
        ' データベースを開く
        Dim cf As New PCfunc
        pblInstPath = cf.GetPath()

        ' データベースを開く
        AdoData.Open(MDBCONNECT & pblInstPath & DIR_HENKAN & CONFIGFILE & ";")

        sSql = "UPDATE Config SET BfDb = '" & fcRetValue & "'"
        AdoData.Execute(sSql)

        AdoData.Close()
        AdoData = Nothing

        'フォーム非表示
        Me.Dispose()

    End Sub
End Class