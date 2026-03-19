-- ============================================
-- T_Request に取引先・案件のコード・名称カラムを追加
--
-- 背景:
--   取引先・案件情報を別DB（Client）で管理するようになったため、
--   JOINで取得するのではなく、依頼登録時に
--   コード・名称をT_Requestに保存するように変更する。
-- ============================================

USE IraiKanri;
GO

ALTER TABLE IRAI.T_Request ADD client_code  NVARCHAR(50)  NULL; -- 取引先コード
GO

ALTER TABLE IRAI.T_Request ADD client_name  NVARCHAR(200) NULL; -- 取引先名
GO

ALTER TABLE IRAI.T_Request ADD project_code NVARCHAR(50)  NULL; -- 案件コード
GO

ALTER TABLE IRAI.T_Request ADD project_name NVARCHAR(200) NULL; -- 案件名
GO
