Imports 仕訳伝票.PCData

Public Class CInitialLoad
    '----------------------------------------------------------------
    '   設定情報ロード
    '----------------------------------------------------------------
    Public Sub InitialLoad()
        Dim db As New ADODB.Connection   'ADOのConnectionオブジェクトを生成
        Dim rs As New ADODB.Recordset

        Dim ssql As String
        Dim wrkConnectInfo As String
        Dim wrkErrBackColor As Long
        Dim wrkErrForeColor As Long
        Dim wrkKinBackColor As Long

        On Error GoTo ErrPrc

        Dim cf As New PCfunc
        pblInstPath = cf.GetPath()

        ' データベースを開く
        db.Open(MDBCONNECT & pblInstPath & DIR_HENKAN & CONFIGFILE & ";")

        'SQL文のみを入力
        ssql = "SELECT * FROM Config"

        rs = db.Execute(ssql)

        ' カラーの取得
        wrkErrBackColor = RGB(rs.Fields("ErBkR").Value, rs.Fields("ErBKG").Value, rs.Fields("ErBkB").Value)
        wrkErrForeColor = RGB(rs.Fields("ErFrR").Value, rs.Fields("ErFrG").Value, rs.Fields("ErFrB").Value)
        wrkKinBackColor = RGB(rs.Fields("KinBkR").Value, rs.Fields("KinBKG").Value, rs.Fields("KinBkB").Value)

        ' 画像表示設定
        pblImageHeight = CInt(rs.Fields("ImgH").Value)
        pblImageWidth = CInt(rs.Fields("ImgW").Value)
        pblImageX = CInt(rs.Fields("ImgX").Value)

        ' 前回選択したデータベース名
        pblBfDbName = Trim(rs.Fields("BfDb").Value)

        ' 接続関連
        pblDsnPath = Trim(rs.Fields("DsnPath").Value)
        pblDsnFlg = Trim(rs.Fields("DsnFlg").Value)

        'メニューで指定した伝票の区分   2004/8/13
        pblSelFILE = Trim(rs.Fields("sub2").Value)

        '振替モード　固定部門、勘定科目,補助科目
        pblHeadBumon = ""
        pblHeadKamoku = ""
        pblHeadHojo = ""

        '------------------
        ' カラーの設定
        '------------------
        pblBackColor = BACK_COLOR
        pblForeColor = FORE_COLOR
        pblNonColor = NON_COLOR
        pblErrBackColor = wrkErrBackColor
        pblErrForeColor = wrkErrForeColor
        pblKinBackColor = wrkKinBackColor

        If Microsoft.VisualBasic.Information.IsDBNull(rs.Fields("DsnPassWord").Value) = True Then
            'NULLの場合
            pblDsnPassWord = ""
        Else
            pblDsnPassWord = Trim(rs.Fields("DsnPassWord").Value)
        End If

        rs.Close()
        db.Close()
        rs = Nothing
        db = Nothing

        '-----------------------
        '   最大結合枚数セット
        '-----------------------
        pblCombineMax = MAXDEN

        Exit Sub

        On Error GoTo 0

        ' エラー処理
ErrPrc:
        Dim em As New CErr
        em.ErrMessage("設定データ取得中")

    End Sub


End Class
