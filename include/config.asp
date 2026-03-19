<%
' ============================================
' 設定ファイル
' データベース接続情報などの設定を管理
' ============================================

' データベース接続文字列（依頼管理DB）
' ※環境に合わせて変更してください
Const DB_CONNECTION_STRING = "Provider=SQLOLEDB;Data Source=DESKTOP-3JSMI33\MSSQLSERVER2022;Initial Catalog=IraiKanri;User ID=testuser;Password=testuser;"

' 社員情報DB接続文字列
' ─────────────────────────────────────────────────────────
' 社員検索・一覧取得など、社員DBへ直接接続するクエリ用の接続文字列
' 依頼管理DBと同一サーバー上の別DB（Employee）を参照する
Const EMPLOYEE_DB_CONNECTION_STRING = "Provider=SQLOLEDB;Data Source=DESKTOP-3JSMI33\MSSQLSERVER2022;Initial Catalog=Employee;User ID=testuser;Password=testuser;"

' 取引先・案件DB接続文字列
' ─────────────────────────────────────────────────────────
' 取引先・案件検索など、Client DBへ直接接続するクエリ用の接続文字列
' 依頼管理DBと同一サーバー上の別DB（Client）を参照する
Const CLIENT_DB_CONNECTION_STRING = "Provider=SQLOLEDB;Data Source=DESKTOP-3JSMI33\MSSQLSERVER2022;Initial Catalog=Client;User ID=testuser;Password=testuser;"

' デバッグ設定
' ─────────────────────────────────────────────────────────
' DEBUG_USER に値を設定すると LOGON_USER の代わりに使用される
' 本番環境では必ず空文字("")にすること
Const DEBUG_USER = "823456"

' アプリケーション設定
Const APP_TITLE = "依頼事項 期限管理システム"
Const APP_VERSION = "1.0.0"

' メール設定
Const MAIL_SMTP_SERVER = "smtp.example.com"
Const MAIL_SMTP_PORT = 25
Const MAIL_FROM_ADDRESS = "noreply@example.com"
Const MAIL_FROM_NAME = "依頼管理システム"

' リマインド設定（期限何日前に通知するか）
Const REMIND_DAYS_BEFORE = 3

' ページネーション設定
Const ITEMS_PER_PAGE = 20

' ステータス定数
Const STATUS_NOT_STARTED = 1    ' 未着手
Const STATUS_IN_PROGRESS = 2    ' 着手済
Const STATUS_COMPLETED = 3      ' 対応終了
Const STATUS_NOT_APPLICABLE = 4 ' 対象外

' メール種別定数
Const MAIL_TYPE_NEW = 1         ' 新規依頼
Const MAIL_TYPE_REMIND = 2      ' リマインド
Const MAIL_TYPE_OVERDUE = 3     ' 期限切れ
%>
