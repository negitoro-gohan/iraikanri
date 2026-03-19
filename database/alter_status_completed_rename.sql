-- ============================================
-- M_Status テーブルの「対応完了」を「対応終了」に変更
-- ============================================

USE IraiKanri;
GO

UPDATE IRAI.M_Status
SET status_name = N'対応終了'
WHERE status_id = 3;
GO

PRINT N'ステータス名の変更が完了しました。';
GO
