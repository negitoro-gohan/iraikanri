-- ============================================================
-- migrate_remove_client_project_ids.sql
-- M_Client から client_id を削除
-- M_Project から project_id, client_id を削除（client_code に移行）
-- T_Request から client_id, project_id を削除
-- ============================================================

USE IraiKanri;
GO

-- ============================================================
-- 1. T_Request から client_id, project_id を削除
-- ============================================================

-- 既存の制約を自動検出して削除（DEFAULT制約）
DECLARE @sql NVARCHAR(MAX);
DECLARE @constraintName NVARCHAR(200);

SELECT @constraintName = dc.name
FROM sys.default_constraints dc
JOIN sys.columns c ON dc.parent_object_id = c.object_id AND dc.parent_column_id = c.column_id
JOIN sys.tables t ON c.object_id = t.object_id
JOIN sys.schemas s ON t.schema_id = s.schema_id
WHERE s.name = 'IRAI' AND t.name = 'T_Request' AND c.name = 'client_id';

IF @constraintName IS NOT NULL BEGIN
    SET @sql = 'ALTER TABLE IRAI.T_Request DROP CONSTRAINT ' + QUOTENAME(@constraintName);
    EXEC sp_executesql @sql;
    PRINT 'Dropped DEFAULT constraint on T_Request.client_id: ' + @constraintName;
END

SET @constraintName = NULL;
SELECT @constraintName = dc.name
FROM sys.default_constraints dc
JOIN sys.columns c ON dc.parent_object_id = c.object_id AND dc.parent_column_id = c.column_id
JOIN sys.tables t ON c.object_id = t.object_id
JOIN sys.schemas s ON t.schema_id = s.schema_id
WHERE s.name = 'IRAI' AND t.name = 'T_Request' AND c.name = 'project_id';

IF @constraintName IS NOT NULL BEGIN
    SET @sql = 'ALTER TABLE IRAI.T_Request DROP CONSTRAINT ' + QUOTENAME(@constraintName);
    EXEC sp_executesql @sql;
    PRINT 'Dropped DEFAULT constraint on T_Request.project_id: ' + @constraintName;
END

-- T_Request.client_id を削除
IF EXISTS (
    SELECT 1 FROM sys.columns c
    JOIN sys.tables t ON c.object_id = t.object_id
    JOIN sys.schemas s ON t.schema_id = s.schema_id
    WHERE s.name = 'IRAI' AND t.name = 'T_Request' AND c.name = 'client_id'
) BEGIN
    ALTER TABLE IRAI.T_Request DROP COLUMN client_id;
    PRINT 'Dropped column T_Request.client_id';
END ELSE BEGIN
    PRINT 'Column T_Request.client_id does not exist, skipping.';
END

-- T_Request.project_id を削除
IF EXISTS (
    SELECT 1 FROM sys.columns c
    JOIN sys.tables t ON c.object_id = t.object_id
    JOIN sys.schemas s ON t.schema_id = s.schema_id
    WHERE s.name = 'IRAI' AND t.name = 'T_Request' AND c.name = 'project_id'
) BEGIN
    ALTER TABLE IRAI.T_Request DROP COLUMN project_id;
    PRINT 'Dropped column T_Request.project_id';
END ELSE BEGIN
    PRINT 'Column T_Request.project_id does not exist, skipping.';
END

GO

-- ============================================================
-- 2. Client DB: M_Client から client_id を削除
--    M_Project から project_id, client_id を削除
-- ============================================================

USE ClientDB;
GO

-- M_Project に client_code カラムを追加（まだない場合）
IF NOT EXISTS (
    SELECT 1 FROM sys.columns c
    JOIN sys.tables t ON c.object_id = t.object_id
    JOIN sys.schemas s ON t.schema_id = s.schema_id
    WHERE s.name = 'IRAI' AND t.name = 'M_Project' AND c.name = 'client_code'
) BEGIN
    ALTER TABLE IRAI.M_Project ADD client_code NVARCHAR(50) NULL;
    PRINT 'Added column M_Project.client_code';
END ELSE BEGIN
    PRINT 'Column M_Project.client_code already exists, skipping.';
END
GO

-- M_Project.client_code にデータを移行（client_id → M_Client.client_code）
UPDATE p
SET p.client_code = c.client_code
FROM IRAI.M_Project p
JOIN IRAI.M_Client c ON p.client_id = c.client_id
WHERE p.client_code IS NULL OR p.client_code = '';
PRINT 'Migrated client_code data from M_Client to M_Project';
GO

-- M_Project の制約を自動検出して削除してから client_id カラムを削除
DECLARE @sql NVARCHAR(MAX);
DECLARE @constraintName NVARCHAR(200);

-- DEFAULT制約
SELECT @constraintName = dc.name
FROM sys.default_constraints dc
JOIN sys.columns c ON dc.parent_object_id = c.object_id AND dc.parent_column_id = c.column_id
JOIN sys.tables t ON c.object_id = t.object_id
JOIN sys.schemas s ON t.schema_id = s.schema_id
WHERE s.name = 'IRAI' AND t.name = 'M_Project' AND c.name = 'client_id';

IF @constraintName IS NOT NULL BEGIN
    SET @sql = 'ALTER TABLE IRAI.M_Project DROP CONSTRAINT ' + QUOTENAME(@constraintName);
    EXEC sp_executesql @sql;
    PRINT 'Dropped DEFAULT constraint on M_Project.client_id: ' + @constraintName;
END

-- M_Project.client_id を削除
IF EXISTS (
    SELECT 1 FROM sys.columns c
    JOIN sys.tables t ON c.object_id = t.object_id
    JOIN sys.schemas s ON t.schema_id = s.schema_id
    WHERE s.name = 'IRAI' AND t.name = 'M_Project' AND c.name = 'client_id'
) BEGIN
    ALTER TABLE IRAI.M_Project DROP COLUMN client_id;
    PRINT 'Dropped column M_Project.client_id';
END ELSE BEGIN
    PRINT 'Column M_Project.client_id does not exist, skipping.';
END

-- M_Project.project_id のDEFAULT制約
SET @constraintName = NULL;
SELECT @constraintName = dc.name
FROM sys.default_constraints dc
JOIN sys.columns c ON dc.parent_object_id = c.object_id AND dc.parent_column_id = c.column_id
JOIN sys.tables t ON c.object_id = t.object_id
JOIN sys.schemas s ON t.schema_id = s.schema_id
WHERE s.name = 'IRAI' AND t.name = 'M_Project' AND c.name = 'project_id';

IF @constraintName IS NOT NULL BEGIN
    SET @sql = 'ALTER TABLE IRAI.M_Project DROP CONSTRAINT ' + QUOTENAME(@constraintName);
    EXEC sp_executesql @sql;
    PRINT 'Dropped DEFAULT constraint on M_Project.project_id: ' + @constraintName;
END

-- M_Project.project_id を削除（project_code がPKになる想定）
IF EXISTS (
    SELECT 1 FROM sys.columns c
    JOIN sys.tables t ON c.object_id = t.object_id
    JOIN sys.schemas s ON t.schema_id = s.schema_id
    WHERE s.name = 'IRAI' AND t.name = 'M_Project' AND c.name = 'project_id'
) BEGIN
    -- PKがproject_idにある場合は先に削除
    DECLARE @pkName NVARCHAR(200);
    SELECT @pkName = kc.name
    FROM sys.key_constraints kc
    JOIN sys.tables t ON kc.parent_object_id = t.object_id
    JOIN sys.schemas s ON t.schema_id = s.schema_id
    WHERE s.name = 'IRAI' AND t.name = 'M_Project' AND kc.type = 'PK';

    IF @pkName IS NOT NULL BEGIN
        SET @sql = 'ALTER TABLE IRAI.M_Project DROP CONSTRAINT ' + QUOTENAME(@pkName);
        EXEC sp_executesql @sql;
        PRINT 'Dropped PK constraint on M_Project: ' + @pkName;
    END

    ALTER TABLE IRAI.M_Project DROP COLUMN project_id;
    PRINT 'Dropped column M_Project.project_id';
END ELSE BEGIN
    PRINT 'Column M_Project.project_id does not exist, skipping.';
END

-- M_Project.project_code をNOT NULLに変更してPKに設定
IF NOT EXISTS (
    SELECT 1 FROM sys.key_constraints kc
    JOIN sys.tables t ON kc.parent_object_id = t.object_id
    JOIN sys.schemas s ON t.schema_id = s.schema_id
    WHERE s.name = 'IRAI' AND t.name = 'M_Project' AND kc.type = 'PK'
) BEGIN
    ALTER TABLE IRAI.M_Project ALTER COLUMN project_code NVARCHAR(50) NOT NULL;
    ALTER TABLE IRAI.M_Project ADD CONSTRAINT PK_M_Project PRIMARY KEY (project_code);
    PRINT 'Set project_code as PRIMARY KEY of M_Project';
END ELSE BEGIN
    PRINT 'M_Project already has a PK, skipping.';
END
GO

-- M_Client から client_id を削除
DECLARE @sql NVARCHAR(MAX);
DECLARE @constraintName NVARCHAR(200);

-- DEFAULT制約
SELECT @constraintName = dc.name
FROM sys.default_constraints dc
JOIN sys.columns c ON dc.parent_object_id = c.object_id AND dc.parent_column_id = c.column_id
JOIN sys.tables t ON c.object_id = t.object_id
JOIN sys.schemas s ON t.schema_id = s.schema_id
WHERE s.name = 'IRAI' AND t.name = 'M_Client' AND c.name = 'client_id';

IF @constraintName IS NOT NULL BEGIN
    SET @sql = 'ALTER TABLE IRAI.M_Client DROP CONSTRAINT ' + QUOTENAME(@constraintName);
    EXEC sp_executesql @sql;
    PRINT 'Dropped DEFAULT constraint on M_Client.client_id: ' + @constraintName;
END

-- M_Client.client_id を削除（client_codeがPKの場合のみ）
IF EXISTS (
    SELECT 1 FROM sys.columns c
    JOIN sys.tables t ON c.object_id = t.object_id
    JOIN sys.schemas s ON t.schema_id = s.schema_id
    WHERE s.name = 'IRAI' AND t.name = 'M_Client' AND c.name = 'client_id'
) BEGIN
    -- PKがclient_idにある場合は先に削除してclient_codeをPKに
    DECLARE @pkName NVARCHAR(200);
    SELECT @pkName = kc.name
    FROM sys.key_constraints kc
    JOIN sys.index_columns ic ON kc.unique_index_id = ic.index_id AND kc.parent_object_id = ic.object_id
    JOIN sys.columns col ON ic.object_id = col.object_id AND ic.column_id = col.column_id
    JOIN sys.tables t ON kc.parent_object_id = t.object_id
    JOIN sys.schemas s ON t.schema_id = s.schema_id
    WHERE s.name = 'IRAI' AND t.name = 'M_Client' AND kc.type = 'PK' AND col.name = 'client_id';

    IF @pkName IS NOT NULL BEGIN
        SET @sql = 'ALTER TABLE IRAI.M_Client DROP CONSTRAINT ' + QUOTENAME(@pkName);
        EXEC sp_executesql @sql;
        PRINT 'Dropped PK constraint on M_Client (was on client_id): ' + @pkName;

        -- client_code を NOT NULL + PK に
        ALTER TABLE IRAI.M_Client ALTER COLUMN client_code NVARCHAR(50) NOT NULL;
        ALTER TABLE IRAI.M_Client ADD CONSTRAINT PK_M_Client PRIMARY KEY (client_code);
        PRINT 'Set client_code as PRIMARY KEY of M_Client';
    END

    ALTER TABLE IRAI.M_Client DROP COLUMN client_id;
    PRINT 'Dropped column M_Client.client_id';
END ELSE BEGIN
    PRINT 'Column M_Client.client_id does not exist, skipping.';
END
GO

PRINT '=== Migration completed ===';
