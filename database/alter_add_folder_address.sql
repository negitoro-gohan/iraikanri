-- ============================================
-- 依頼事項テーブルにフォルダアドレスカラムを追加
-- ============================================

USE IraiKanri;
GO

ALTER TABLE IRAI.T_Request
    ADD folder_address NVARCHAR(500) NULL;
GO

PRINT N'フォルダアドレスカラムの追加が完了しました。';
GO
