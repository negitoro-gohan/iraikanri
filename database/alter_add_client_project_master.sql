-- ============================================
-- 取引先マスター・案件マスター追加
-- ============================================

USE IraiKanri;
GO

-- 取引先マスター
CREATE TABLE M_Client (
    client_id INT IDENTITY(1,1) PRIMARY KEY,    -- 取引先ID
    client_code NVARCHAR(50) NOT NULL,           -- 取引先コード
    client_name NVARCHAR(200) NOT NULL,          -- 取引先名
    is_active BIT DEFAULT 1,                     -- 有効フラグ
    created_at DATETIME DEFAULT GETDATE(),       -- 作成日時
    updated_at DATETIME DEFAULT GETDATE()        -- 更新日時
);
GO

-- 取引先コードにユニークインデックス
CREATE UNIQUE INDEX IX_Client_Code ON M_Client(client_code);
GO

-- 案件マスター（取引先に紐づく）
CREATE TABLE M_Project (
    project_id INT IDENTITY(1,1) PRIMARY KEY,   -- 案件ID
    client_id INT NOT NULL,                      -- 取引先ID（外部キー）
    project_code NVARCHAR(50) NOT NULL,          -- 案件コード
    project_name NVARCHAR(200) NOT NULL,         -- 案件名
    is_active BIT DEFAULT 1,                     -- 有効フラグ
    created_at DATETIME DEFAULT GETDATE(),       -- 作成日時
    updated_at DATETIME DEFAULT GETDATE(),       -- 更新日時

    CONSTRAINT FK_Project_Client FOREIGN KEY (client_id) REFERENCES M_Client(client_id)
);
GO

-- 案件コードにユニークインデックス
CREATE UNIQUE INDEX IX_Project_Code ON M_Project(project_code);
GO

-- 案件の取引先IDにインデックス
CREATE INDEX IX_Project_ClientId ON M_Project(client_id);
GO

-- T_Requestテーブルの修正（client_code, project_code を client_id, project_id に変更）
-- 既存のカラムを削除して新しいカラムを追加
ALTER TABLE T_Request DROP COLUMN client_code;
ALTER TABLE T_Request DROP COLUMN project_code;
GO

ALTER TABLE T_Request ADD client_id INT NULL;
ALTER TABLE T_Request ADD project_id INT NULL;
GO

-- 外部キー制約を追加
ALTER TABLE T_Request ADD CONSTRAINT FK_Request_Client FOREIGN KEY (client_id) REFERENCES M_Client(client_id);
ALTER TABLE T_Request ADD CONSTRAINT FK_Request_Project FOREIGN KEY (project_id) REFERENCES M_Project(project_id);
GO

-- インデックス追加
CREATE INDEX IX_Request_ClientId ON T_Request(client_id);
CREATE INDEX IX_Request_ProjectId ON T_Request(project_id);
GO

-- サンプルデータ
INSERT INTO M_Client (client_code, client_name) VALUES
(N'C001', N'株式会社ABC商事'),
(N'C002', N'有限会社XYZ工業'),
(N'C003', N'合同会社テスト');
GO

INSERT INTO M_Project (client_id, project_code, project_name) VALUES
(1, N'P2026-001', N'基幹システム更新'),
(1, N'P2026-002', N'Webサイトリニューアル'),
(2, N'P2026-003', N'業務効率化支援'),
(3, N'P2026-004', N'新規事業立ち上げ');
GO

PRINT N'取引先マスター・案件マスターの追加が完了しました。';
GO
