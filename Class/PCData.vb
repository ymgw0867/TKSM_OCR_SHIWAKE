Public Class PCData

    '****** CONST ********
    Public Const MDBCONNECT As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
    Public Const CONFIGFILE As String = "Kanjo2kconfig.mdb"         '設定データベース
    Public Const DIR_HENKAN As String = "henkan\"
    Public Const DIR_INCSV As String = "分割\"
    Public Const DIR_BREAK As String = "中断\"
    Public Const DIR_CONFIG As String = "henkan\"                   'フォルダ名
    Public Const DIR_OK As String = "Ok\"

    Public Const DIVFILE As String = "div.csv"                   '分割ファイル
    Public Const OUTFILE As String = "ＯＣＲ勘定奉行変換.csv"    '出力ファイル
    Public Const INFILE As String = "kanjo2kocr.csv"            '入力ファイル
    Public Const LOGFILE As String = "kanjo2kerrlog.csv"         'エラーログファイル
    Public Const TMPREAD As String = "kanjo2ktmpread.dat"        '入力ファイルのコピー
    Public Const tmpFile As String = "kanjo2ktmpfile.dat"        '出力ファイルのコピー
    Public Const Title As String = "勘定奉行　仕訳伝票"
    Public Const KANMA As String = ","

    Public Const FLGON As String = "1"
    Public Const FLGOFF As String = "0"

    '----------------------------------------------------------------
    '   システムデータ
    '----------------------------------------------------------------
    Public Const MAXGYOU As Integer = 7
    Public Const MAXTEKIYOU As Integer = 20
    Public Const MAXDEN As Integer = 36
    Public Const MAX21 As Integer = 250

    '----------------------------------------------------------------
    '    印刷
    '----------------------------------------------------------------
    Public Const PRINTFONTSIZE As Integer = 8
    Public Const PRINTXBASE As Integer = 4
    Public Const PRINTYBASE As Integer = 1
    Public Const PRINTMAXGYOU As Integer = 5
    Public Const PRINTMODEONE As Integer = 0
    Public Const PRINTMODEALL As Integer = 1
    '----------------------------------------------------------------
    '    テキストボックス表示色
    '----------------------------------------------------------------
    Public Const BACK_COLOR As Long = 16777215
    Public Const FORE_COLOR As Long = CLng(&HFF0000) 'vbBlue
    Public Const NON_COLOR As Long = -2147483633
    Public Const ERROR_COLOR As Long = CLng(&HFF)      'vbRed
    Public Const ERROR_BACK_COLOR As Long = CLng(&HFFFF)    'vbYellow
    Public Const TEKIYOU_COLOR As Long = CLng(&H8000)

    '----------------------------------------------------------------
    '   伝票マンのセル名
    '----------------------------------------------------------------
    Public Const DP_DENYEAR As String = "txtDenYear"
    Public Const DP_DENMONTH As String = "txtDenMonth"
    Public Const DP_DENDAY As String = "txtDenDay"
    Public Const DP_DENNO As String = "txtDenNo"
    Public Const DP_DENCOMCODE As String = "txtCom"

    Public Const DP_FUCHK As String = "chk"
    Public Const DP_DELCHK As String = "chkDel"

    Public Const DP_KARI_CODE As String = "txtKCode_K"
    Public Const DP_KARI_NAME As String = "txtKName_K"
    Public Const DP_KARI_CODEH As String = "txtHojo_K"
    Public Const DP_KARI_NAMEH As String = "txtHojoName_K"
    Public Const DP_KARI_CODEB As String = "txtBCode_K"
    Public Const DP_KARI_NAMEB As String = "txtBName_K"
    Public Const DP_KARI_KIN As String = "txtKin_K"
    Public Const DP_KARI_ZEI_S As String = "txtZeis_K"
    Public Const DP_KARI_ZEI As String = "txtZeik_K"
    Public Const DP_KARI_P As String = "txtKari_P"
    Public Const DP_KARI_T As String = "txtKari_T"

    Public Const DP_KASHI_CODE As String = "txtKCode_S"
    Public Const DP_KASHI_NAME As String = "txtKName_S"
    Public Const DP_KASHI_CODEH As String = "txtHojo_S"
    Public Const DP_KASHI_NAMEH As String = "txtHojoName_S"
    Public Const DP_KASHI_CODEB As String = "txtBCode_S"
    Public Const DP_KASHI_NAMEB As String = "txtBName_S"
    Public Const DP_KASHI_KIN As String = "txtKin_S"
    Public Const DP_KASHI_ZEI_S As String = "txtZeis_S"
    Public Const DP_KASHI_ZEI As String = "txtZeik_S"
    Public Const DP_KASHI_P As String = "txtKashi_P"
    Public Const DP_KASHI_T As String = "txtKashi_T"

    Public Const DP_SAGAKU_P As String = "txtSagaku_P"
    Public Const DP_SAGAKU_T As String = "txtSagaku_T"

    Public Const DP_TEKIYOU As String = "txtTekiyou"
    Public Const DP_INDEX As Integer = 22

    Public Const DP_FUKU As String = "fukusu"
    Public Const DP_KESSAN As String = "kessan"

    '----------------------------------------------------------------
    '   エラー種別
    '----------------------------------------------------------------
    Public Const ERRKIND_NOERR As Integer = 0
    Public Const ERRKIND_DATE As Integer = 1
    Public Const ERRKIND_DATE_KIKAN As Integer = 2
    Public Const ERRKIND_DATE_LIMIT As Integer = 3
    Public Const ERRKIND_DATE_KESSAN As Integer = 4
    Public Const ERRKIND_DENNO As Integer = 5
    Public Const ERRKIND_COMBINE As Integer = 6
    Public Const ERRKIND_COMBINECHK As Integer = 7
    Public Const ERRKIND_COMBINE_DATE As Integer = 8
    Public Const ERRKIND_COMBINE_DENNO As Integer = 9
    Public Const ERRKIND_COMBINE_KESSAN As Integer = 10
    Public Const ERRKIND_DATAPOOR As Integer = 11
    Public Const ERRKIND_KAMOKU As Integer = 12
    Public Const ERRKIND_HOJO As Integer = 13
    Public Const ERRKIND_BUMON As Integer = 14
    Public Const ERRKIND_TAXMAS As Integer = 15
    Public Const ERRKIND_TAXKBN As Integer = 16
    Public Const ERRKIND_KIN As Integer = 17
    Public Const ERRKIND_SUMCHK As Integer = 18
    Public Const ERRKIND_AITE As Integer = 19

    Public Const ERRPLACE_YEAR As Integer = 1
    Public Const ERRPLACE_MONTH As Integer = 2
    Public Const ERRPLACE_DAY As Integer = 3
    Public Const ERRPLACE_BUMON As Integer = 4
    Public Const ERRPLACE_KAMOKU As Integer = 5
    Public Const ERRPLACE_HOJO As Integer = 6
    Public Const ERRPLACE_KIN As Integer = 7
    Public Const ERRPLACE_TAXKBN As Integer = 8

    Public Const TAB_ERR As Integer = 0
    Public Const TAB_KAMOKU As Integer = 1
    Public Const TAB_BUMON As Integer = 2
    Public Const TAB_TAXKBN As Integer = 3
    Public Const TAB_COM As Integer = 4


    '****** 変数 ********
    Public Shared pblComNo As Integer                   '会社番号
    Public Shared pblBumonFlg As Boolean                'マスターデータチェックフラグ
    Public Shared pblInstPath As String                 'インストールディレクトリ
    Public Shared pblDbName As String                   '選択された会社のデータベース名
    Public Shared gsTaxMas As String                    '取得方法追加「税処理を取得
    Public Shared gsVersion As String                   ' 取得方法追加「バージョンを取得」

    Public Shared pblBumonData() As strCommonData       '部門データ
    Public Shared pblKamokuData() As strKamokuData      '勘定科目データ
    Public Shared pblHojoData() As strHojoData          '補助データ
    Public Shared pblTaxKbnData() As strCommonData    '税区分データ
    Public Shared pblTaxMasData(1) As strCommonData    '消費税計算区分
    Public Shared pblTekiMst() As strCommonData    '摘要データ

    Public Shared pblDsnPath As String                  'データソースのパス
    Public Shared pblDsnPassWord As String              'DSNパスワード
    Public Shared pblComData As strCompanyData          '会社データ
    Public Shared pblBfDbName As String                 '前回選択したデータベース名
    Public Shared pblDsnFlg As String                   'LAN使用のフラグ
    Public Shared pblSelFILE As Integer                 '選択ファイル　0:OCR,1:中断

    '----------------------------------------------------------------
    '   中断リカバリーファイル 
    '----------------------------------------------------------------
    Public Shared pblRecov() As strRecovery     '中断フォルダからリカバリーするファイル名

    '----------------------------------------------------------------
    '   入力伝票データ
    '----------------------------------------------------------------
    Public Shared pblDenRec() As strInputRecord    '読み込みデータ
    Public Shared pblDenNum As Integer           'データ数

    '----------------------------------------------------------------
    '   固定部門、勘定科目,補助科目
    '----------------------------------------------------------------
    Public Shared pblHeadBumon As String
    Public Shared pblHeadKamoku As String
    Public Shared pblHeadHojo As String

    '----------------------------------------------------------------
    '   画像表示
    '----------------------------------------------------------------
    Public Shared pblImageX As Integer      '画像の表示倍率（％）
    Public Shared pblImageHeight As Integer      '画像フォームの高さ
    Public Shared pblImageWidth As Integer      '画像フォームの幅

    '----------------------------------------------------------------
    '   画面デザイン
    '----------------------------------------------------------------
    Public Shared pblBackColor As Long
    Public Shared pblForeColor As Long
    Public Shared pblNonColor As Long
    Public Shared pblErrBackColor As Long
    Public Shared pblErrForeColor As Long
    Public Shared pblKinBackColor As Long

    '----------------------------------------------------------------
    '   入力データ有無フラグ
    '----------------------------------------------------------------
    Public Shared pblFlgINFILE As Boolean      '変換データ
    Public Shared pblFlgDIVFILE As Boolean      '分割後データ
    Public Shared pblFlgRecFILE As Boolean      '中断データ

    '----------------------------------------------------------------
    '   その他
    '----------------------------------------------------------------
    Public Shared pblCombineMax As Integer      '伝票結合最大枚数
    Public Shared pblBeforeTekiyou As String    '前行の摘要文
    Public Shared pblFirstGyouFlg As Boolean    '出力時、最初の行判定
    Public Shared pblMaisu As Integer           '伝票結合枚数(枚数チェック用）
    Public Shared pblItem As Integer      '伝票結合行数(チェック用）

    '----------------------------------------------------------------
    '   印刷
    '----------------------------------------------------------------
    Public Shared pblPrintMode As Integer       '印刷に渡すモード
    Public Shared pblPrintFile As String        '画像印刷ファイル名

    '----------------------------------------------------------------
    '   データ表示用
    '----------------------------------------------------------------
    Public Shared pblNowden As Integer
    Public Shared pblNowGyou As Integer

    '----------------------------------------------------------------
    '   エラー関係
    '----------------------------------------------------------------
    Public Shared pblNowErrNo As Integer
    Public Shared pblNowErrNum As Integer
    Public Shared pblErrCnt As Integer
    Public Shared pblSagakuFLG As Boolean          '差額エラー発生フラグ
    Public Shared pblErrTBL() As strErrtbl

    '----------------------------------------------------------------
    '   入力期間制限
    '----------------------------------------------------------------
    Public Shared pblLimitData As strLimitDateData     '日付入力期間　マスター内の指定期間

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
        Dim Name As String              '補助名
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

    '----------------------------------------------------------------
    '   会社選択
    '----------------------------------------------------------------
    Structure strComDBData
        Dim DbName As String            'DBネーム
        Dim Name As String              '名前
        Dim ComNo As String             '会社番号
        Dim FromYear As String          '会計期間期首年
        Dim FromMonth As String         '会計期間期首月
        Dim FromDay As String           '会計期間期首日
        Dim Hosei As String             '元号補正値
        Dim KessanKi As String          '決算期
        Dim TaxMas As String            '消費税計算区分
    End Structure

    '----------------------------------------------------------------
    '   日付の入力範囲
    '----------------------------------------------------------------
    Structure strLimitDateData      '日付の入力範囲データ
        Dim FromYear As String      '入力開始年
        Dim FromMonth As String     '入力開始月
        Dim FromDay As String       '入力開始日
        Dim StSoeji As String       '入力期間開始添え字
        Dim ToYear As String        '入力期限年
        Dim ToMonth As String       '入力期限月
        Dim ToDay As String         '入力期限日
        Dim EdSoeji As String       '入力期間終了添え字
        Dim Lock As String          '制限の種類
        Dim Flag As Boolean         '入力可能フラグ
    End Structure

    Structure strInKariKashi
        Dim Bumon As String         '部門コード
        Dim Kamoku As String        '科目コード
        Dim Hojo As String          '補助コード
        Dim Kin As String           '金額
        Dim TaxMas As String        '消費税計算区分
        Dim TaxKbn As String        '税区分
    End Structure

    Structure strInGyou
        Dim GyouNum As String           '行番号
        Dim Kari As strInKariKashi      '借方データ
        Dim Kashi As strInKariKashi     '貸方データ
        Dim CopyChk As String           '摘要複写チェック
        Dim Tekiyou As String           '摘要
        Dim Torikeshi As String          '取消チェック

    End Structure

    Structure strInHead
        Dim Image As String         '画像ファイル名
        Dim CsvFile As String       'CSVファイル名
        Dim Year As String          '年
        Dim Month As String         '月
        Dim Day As String           '日
        Dim Kessan As String        '決算処理フラグ
        Dim FukusuChk As String     '複数毎チェック
        Dim DenNo As String         '伝票No.
        Dim Kari_T As Decimal       '借方伝票計
        Dim Kashi_T As Decimal      '貸方伝票計
        Dim FukuMai As Integer      '複数毎数
    End Structure

    Structure strInputRecord
        Dim Head As strInHead    'ヘッダ部
        'Dim Gyou(0 To MAXGYOU) As strInGyou    '行データ
        Dim Gyou() As strInGyou
        Dim KariTotal As Decimal    '借方貸方合計
        Dim KashiTotal As Decimal
    End Structure

    '----------------------------------------------------------------
    '   中断リカバリーファイル
    '----------------------------------------------------------------
    Structure strRecovery
        Dim recFlg As Integer       '読込フラグ 0:対象外,1:読み込む
        Dim recName As String       'ファイル名
        Dim recFuku As Integer      '枚数
    End Structure

    '----------------------------------------------------------------
    '   出力データ
    '----------------------------------------------------------------
    Structure strOutKamokuData
        Dim Bumon As String         '部門
        Dim Kamoku As String        '科目
        Dim Hojo As String          '補助
        Dim Kin As String           '金額
        Dim Tax As String           '消費税額（空欄）
        Dim TaxMas As String        '消費税計算区分
        Dim TaxKbn As String        '税区分
        Dim JigyoKbn As String      '事業区分（空欄）
    End Structure

    Structure strOutRecord          '出力データ
        Dim Kugiri As String        '伝票区切
        Dim Kessan As String        '決算フラグ
        Dim DenDate As String       '伝票日付
        Dim DenNo As String         '伝票番号
        Dim Kari As strOutKamokuData '借方データ
        Dim Kashi As strOutKamokuData '貸方データ
        Dim Tekiyou As String         '摘要
    End Structure

    '----------------------------------------------------------------
    '   エラー情報
    '----------------------------------------------------------------
    Structure strErrInfo                 'エラー情報
        Dim FriendErrKind As Integer          'エラー種別
        Dim FriendErrDen As Integer           'エラー伝票配列番号
        Dim FriendErrGyou As Integer          'エラー行番号
        Dim FriendErrKariKashi As String      'エラー借方貸方
        Dim FriendListMsg As String           'エラーリスト上のメッセージ
        Dim FriendErrSagaku As String         'エラー賃貸差額
        Dim FriendErrDenNo As String          'エラー伝票番号
        Dim FriendErrYear As String           'エラー伝票年
        Dim FriendErrMonth As String          'エラー伝票月
        Dim FriendErrDay As String            'エラー伝票日
    End Structure

    Structure strErrtbl
        Dim FriendErrDenNo As Integer       'エラー伝票番号
        Dim FriendErrLINE As Integer        'エラー行番号
        Dim ErrField As String              'エラーフィールド
        Dim FriendErrData As String         'エラーデータ
        Dim FriendErrNotes As String        'エラー備考
        Dim FriendErrDpPos As String        '伝票マンのセル名
    End Structure
End Class