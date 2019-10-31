Public Class CData
    Public Const MDBCONNECT As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
    Public Const DIR_CONFIG As String = "henkan\"                   'フォルダ名
    Public Const CONFIGFILE As String = "Kanjo2kconfig.mdb"         '設定データベース
    Public Const DIR_HENKAN As String = "henkan\"
    Public Const TMPREAD As String = "kanjo2ktmpread.dat"        '入力ファイルのコピー
    Public Const tmpFile As String = "kanjo2ktmpfile.dat"        '出力ファイルのコピー
    Public Const Title As String = "勘定奉行　固定項目設定"

    Public Shared pblBumonFlg As Boolean               '   マスターデータチェックフラグ
    Public Shared pblInstPath As String         'インストールディレクトリ
    Public Shared pblDbName As String           '選択された会社のデータベース名
    Public Shared gsTaxMas As String            '取得方法追加「税処理を取得
    Public Shared gsVersion As String           ' 取得方法追加「バージョンを取得」
    Public Shared pblBumonData() As strCommonData    '部門データ
    Public pblKamokuData() As strKamokuData    '勘定科目データ
    Public pblHojoData() As strHojoData      '補助データ
    Public Shared pblDsnPath As String               'データソースのパス
    Public Shared pblDsnPassWord As String          'DSNパスワード
    Public Shared pblComData As strCompanyData   '会社データ

    '----------------------------------------------------------------
    '   共通データ
    '----------------------------------------------------------------
    Structure strCommonData
        Dim Code As String              'コード
        Dim Name As String              '名前
    End Structure
    '----------------------------------------------------------------
    '   補助データ
    '----------------------------------------------------------------
    Structure strHojoData
        Dim KamokuCode As String        '科目コード
        Dim Code As String              '補助コード
    dimName As String              '補助名
    End Structure

    '----------------------------------------------------------------
    '   勘定科目データ
    '----------------------------------------------------------------
    Structure strKamokuData
        Dim Code As String
        Dim Name As String
        Dim Ncd As String
        Dim IsZei As String             '税処理フラグ
        Dim HojoExist As Boolean
        Dim HojoData() As strHojoData
    End Structure

    '----------------------------------------------------------------
    '   会社データ
    '----------------------------------------------------------------
    Structure strCompanyData     '会社コードデータ
        Dim Name As String       '会社名
        Dim FromYear As String   '会計期間期首年
        Dim FromMonth As String  '会計期間期首月
        Dim FromDay As String    '会計期間期首日
        Dim ToYear As String     '会計期間期末年
        Dim ToMonth As String    '会計期間期末月
        Dim ToDay As String      '会計期間期末日
        Dim Kaisi As String      '入力開始月
        Dim TaxMas As String     '消費税計算区分
        Dim Gengou As String     '元号
        Dim Hosei As String      '年号補正値
        Dim Middle As String     '中間期決算フラグ
        Dim Reki As String       '西暦年または元号
    End Structure
End Class
