-- ============================================
-- 取引先コード・案件コード追加
-- ============================================

USE IraiKanri;
GO

-- T_Requestテーブルに取引先コード・案件コードを追加
ALTER TABLE T_Request ADD client_code NVARCHAR(50) NULL;      -- 取引先コード
ALTER TABLE T_Request ADD project_code NVARCHAR(50) NULL;     -- 案件コード
GO

-- インデックス追加
CREATE INDEX IX_Request_ClientCode ON T_Request(client_code);
CREATE INDEX IX_Request_ProjectCode ON T_Request(project_code);
GO

PRINT N'取引先コード・案件コードの追加が完了しました。';
GO
