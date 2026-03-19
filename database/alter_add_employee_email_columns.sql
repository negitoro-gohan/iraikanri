-- ============================================
-- T_Request に社員メールアドレスカラムを追加
--
-- 背景:
--   社員情報を別DB（Employee）で管理しているため、
--   メール送信・表示のたびに別DBへアクセスするのを避けるため
--   依頼登録時に取得したメールアドレスをT_Requestに保存する。
-- ============================================

USE IraiKanri;
GO

ALTER TABLE IRAI.T_Request ADD requester_email NVARCHAR(256) NULL; -- 依頼元メールアドレス
GO

ALTER TABLE IRAI.T_Request ADD assignee_email  NVARCHAR(256) NULL; -- 依頼先メールアドレス
GO
