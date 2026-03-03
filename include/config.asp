<%
' ============================================
' 設定ファイル
' データベース接続情報などの設定を管理
' ============================================

' データベース接続文字列
' ※環境に合わせて変更してください
Const DB_CONNECTION_STRING = "Provider=SQLOLEDB;Data Source=DESKTOP-3JSMI33\MSSQLSERVER2022;Initial Catalog=IraiKanri;User ID=testuser;Password=testuser;"

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
Const STATUS_COMPLETED = 3      ' 対応完了
Const STATUS_NOT_APPLICABLE = 4 ' 対象外

' メール種別定数
Const MAIL_TYPE_NEW = 1         ' 新規依頼
Const MAIL_TYPE_REMIND = 2      ' リマインド
Const MAIL_TYPE_OVERDUE = 3     ' 期限切れ
%>
