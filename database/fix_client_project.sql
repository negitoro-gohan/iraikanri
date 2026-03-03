-- ============================================
-- 取引先・案件マスター修正スクリプト
-- 現在の状態を確認して必要な修正を行う
-- ============================================

USE IraiKanri;
GO

-- 1. M_Clientテーブルが存在するか確認し、なければ作成
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'M_Client')
BEGIN
    CREATE TABLE M_Client (
        client_id INT IDENTITY(1,1) PRIMARY KEY,
        client_code NVARCHAR(50) NOT NULL,
        client_name NVARCHAR(200) NOT NULL,
        is_active BIT DEFAULT 1,
        created_at DATETIME DEFAULT GETDATE(),
        updated_at DATETIME DEFAULT GETDATE()
    );
    CREATE UNIQUE INDEX IX_Client_Code ON M_Client(client_code);
    PRINT N'M_Clientテーブルを作成しました。';
END
ELSE
BEGIN
    PRINT N'M_Clientテーブルは既に存在します。';
END
GO

-- 2. M_Projectテーブルが存在するか確認し、なければ作成
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'M_Project')
BEGIN
    CREATE TABLE M_Project (
        project_id INT IDENTITY(1,1) PRIMARY KEY,
        client_id INT NOT NULL,
        project_code NVARCHAR(50) NOT NULL,
        project_name NVARCHAR(200) NOT NULL,
        is_active BIT DEFAULT 1,
        created_at DATETIME DEFAULT GETDATE(),
        updated_at DATETIME DEFAULT GETDATE(),
        CONSTRAINT FK_Project_Client FOREIGN KEY (client_id) REFERENCES M_Client(client_id)
    );
    CREATE UNIQUE INDEX IX_Project_Code ON M_Project(project_code);
    CREATE INDEX IX_Project_ClientId ON M_Project(client_id);
    PRINT N'M_Projectテーブルを作成しました。';
END
ELSE
BEGIN
    PRINT N'M_Projectテーブルは既に存在します。';
END
GO

-- 3. T_Requestテーブルにclient_idカラムがあるか確認
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('T_Request') AND name = 'client_id')
BEGIN
    -- client_codeカラムがあれば削除
    IF EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('T_Request') AND name = 'client_code')
    BEGIN
        ALTER TABLE T_Request DROP COLUMN client_code;
        PRINT N'client_codeカラムを削除しました。';
    END

    -- client_idカラムを追加
    ALTER TABLE T_Request ADD client_id INT NULL;
    PRINT N'client_idカラムを追加しました。';
END
ELSE
BEGIN
    PRINT N'client_idカラムは既に存在します。';
END
GO

-- 4. T_Requestテーブルにproject_idカラムがあるか確認
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('T_Request') AND name = 'project_id')
BEGIN
    -- project_codeカラムがあれば削除
    IF EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('T_Request') AND name = 'project_code')
    BEGIN
        ALTER TABLE T_Request DROP COLUMN project_code;
        PRINT N'project_codeカラムを削除しました。';
    END

    -- project_idカラムを追加
    ALTER TABLE T_Request ADD project_id INT NULL;
    PRINT N'project_idカラムを追加しました。';
END
ELSE
BEGIN
    PRINT N'project_idカラムは既に存在します。';
END
GO

-- 5. 不要なclient_code, project_codeカラムが残っていれば削除
IF EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('T_Request') AND name = 'client_code')
BEGIN
    ALTER TABLE T_Request DROP COLUMN client_code;
    PRINT N'残っていたclient_codeカラムを削除しました。';
END
GO

IF EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('T_Request') AND name = 'project_code')
BEGIN
    ALTER TABLE T_Request DROP COLUMN project_code;
    PRINT N'残っていたproject_codeカラムを削除しました。';
END
GO

-- 6. 外部キー制約の追加（存在しない場合のみ）
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE name = 'FK_Request_Client')
BEGIN
    ALTER TABLE T_Request ADD CONSTRAINT FK_Request_Client FOREIGN KEY (client_id) REFERENCES M_Client(client_id);
    PRINT N'FK_Request_Client制約を追加しました。';
END
GO

IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE name = 'FK_Request_Project')
BEGIN
    ALTER TABLE T_Request ADD CONSTRAINT FK_Request_Project FOREIGN KEY (project_id) REFERENCES M_Project(project_id);
    PRINT N'FK_Request_Project制約を追加しました。';
END
GO

-- 7. インデックスの追加（存在しない場合のみ）
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_Request_ClientId' AND object_id = OBJECT_ID('T_Request'))
BEGIN
    CREATE INDEX IX_Request_ClientId ON T_Request(client_id);
    PRINT N'IX_Request_ClientIdインデックスを追加しました。';
END
GO

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_Request_ProjectId' AND object_id = OBJECT_ID('T_Request'))
BEGIN
    CREATE INDEX IX_Request_ProjectId ON T_Request(project_id);
    PRINT N'IX_Request_ProjectIdインデックスを追加しました。';
END
GO

-- 8. サンプルデータの追加（データがない場合のみ）
IF NOT EXISTS (SELECT * FROM M_Client)
BEGIN
    INSERT INTO M_Client (client_code, client_name) VALUES
    (N'C001', N'株式会社ABC商事'),
    (N'C002', N'有限会社XYZ工業'),
    (N'C003', N'合同会社テスト');
    PRINT N'取引先サンプルデータを追加しました。';
END
GO

IF NOT EXISTS (SELECT * FROM M_Project)
BEGIN
    INSERT INTO M_Project (client_id, project_code, project_name) VALUES
    (1, N'P2026-001', N'基幹システム更新'),
    (1, N'P2026-002', N'Webサイトリニューアル'),
    (2, N'P2026-003', N'業務効率化支援'),
    (3, N'P2026-004', N'新規事業立ち上げ');
    PRINT N'案件サンプルデータを追加しました。';
END
GO

-- 9. 現在のテーブル構造を確認
PRINT N'';
PRINT N'=== T_Requestテーブルのカラム一覧 ===';
SELECT COLUMN_NAME, DATA_TYPE, IS_NULLABLE
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = 'T_Request'
ORDER BY ORDINAL_POSITION;

PRINT N'';
PRINT N'=== 修正完了 ===';
GO
