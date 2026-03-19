-- ============================================
-- 依頼事項テーブルに重要度カラムを追加
-- 重要度: A（高）, B（中）, C（低）
-- ============================================

USE IraiKanri;
GO

ALTER TABLE IRAI.T_Request
    ADD importance CHAR(1) DEFAULT 'B';
GO

-- 既存レコードのデフォルト値を確定
UPDATE IRAI.T_Request SET importance = 'B' WHERE importance IS NULL;
GO

-- NULL不可に変更
ALTER TABLE IRAI.T_Request
    ALTER COLUMN importance CHAR(1) NOT NULL;
GO

-- 値制約（A/B/Cのみ許可）
ALTER TABLE IRAI.T_Request
    ADD CONSTRAINT CK_Request_Importance CHECK (importance IN ('A', 'B', 'C'));
GO

PRINT N'重要度カラムの追加が完了しました。';
GO
