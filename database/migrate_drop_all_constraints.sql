-- ============================================
-- 全制約削除 + employee_id 廃止マイグレーション
--
-- 対象:
--   IraiKanri DB: T_Request の全制約を削除し
--                 requester_id/assignee_id を削除
--                 requester_code/assignee_code を追加
--   Employee DB : M_Employee の全制約を削除し
--                 employee_id を削除
--                 employee_code を PRIMARY KEY に変更
-- ============================================


-- ============================================
-- IraiKanri DB
-- ============================================
USE IraiKanri;
GO

-- T_Request の全制約（FK・PK・UNIQUE・CHECK・DEFAULT）を削除
DECLARE @sql NVARCHAR(MAX) = N'';

SELECT @sql = @sql +
    N'ALTER TABLE ' + QUOTENAME(s.name) + N'.' + QUOTENAME(t.name) +
    N' DROP CONSTRAINT ' + QUOTENAME(c.name) + N';' + CHAR(13)
FROM sys.objects c
INNER JOIN sys.tables t ON c.parent_object_id = t.object_id
INNER JOIN sys.schemas s ON t.schema_id = s.schema_id
WHERE t.name = 'T_Request'
  AND s.name = 'IRAI'
  AND c.type IN ('F','PK','UQ','C','D');  -- FK, PK, UNIQUE, CHECK, DEFAULT

IF @sql <> N''
BEGIN
    PRINT N'T_Request の制約を削除します:';
    PRINT @sql;
    EXEC sp_executesql @sql;
END
ELSE
    PRINT N'T_Request に削除すべき制約はありませんでした。';
GO

-- requester_code / assignee_code が未追加なら追加
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS
               WHERE TABLE_SCHEMA = 'IRAI' AND TABLE_NAME = 'T_Request' AND COLUMN_NAME = 'requester_code')
    ALTER TABLE IRAI.T_Request ADD requester_code NVARCHAR(20) NULL;
GO
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS
               WHERE TABLE_SCHEMA = 'IRAI' AND TABLE_NAME = 'T_Request' AND COLUMN_NAME = 'assignee_code')
    ALTER TABLE IRAI.T_Request ADD assignee_code NVARCHAR(20) NULL;
GO

-- requester_id / assignee_id が残っていれば削除
IF EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS
           WHERE TABLE_SCHEMA = 'IRAI' AND TABLE_NAME = 'T_Request' AND COLUMN_NAME = 'requester_id')
    ALTER TABLE IRAI.T_Request DROP COLUMN requester_id;
GO
IF EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS
           WHERE TABLE_SCHEMA = 'IRAI' AND TABLE_NAME = 'T_Request' AND COLUMN_NAME = 'assignee_id')
    ALTER TABLE IRAI.T_Request DROP COLUMN assignee_id;
GO

PRINT N'IraiKanri DB の変更が完了しました。';
GO


-- ============================================
-- Employee DB
-- ============================================
USE Employee;
GO

-- M_Employee の全制約を削除
DECLARE @sql NVARCHAR(MAX) = N'';

SELECT @sql = @sql +
    N'ALTER TABLE ' + QUOTENAME(s.name) + N'.' + QUOTENAME(t.name) +
    N' DROP CONSTRAINT ' + QUOTENAME(c.name) + N';' + CHAR(13)
FROM sys.objects c
INNER JOIN sys.tables t ON c.parent_object_id = t.object_id
INNER JOIN sys.schemas s ON t.schema_id = s.schema_id
WHERE t.name = 'M_Employee'
  AND s.name = 'IRAI'
  AND c.type IN ('F','PK','UQ','C','D');

IF @sql <> N''
BEGIN
    PRINT N'M_Employee の制約を削除します:';
    PRINT @sql;
    EXEC sp_executesql @sql;
END
ELSE
    PRINT N'M_Employee に削除すべき制約はありませんでした。';
GO

-- M_Employee の全インデックスを削除（PK以外）
DECLARE @sql NVARCHAR(MAX) = N'';

SELECT @sql = @sql +
    N'DROP INDEX ' + QUOTENAME(i.name) + N' ON IRAI.M_Employee;' + CHAR(13)
FROM sys.indexes i
INNER JOIN sys.tables t ON i.object_id = t.object_id
INNER JOIN sys.schemas s ON t.schema_id = s.schema_id
WHERE t.name = 'M_Employee'
  AND s.name = 'IRAI'
  AND i.type > 0          -- 0 = ヒープ（インデックスなし）を除外
  AND i.is_primary_key = 0;

IF @sql <> N''
BEGIN
    PRINT N'M_Employee のインデックスを削除します:';
    PRINT @sql;
    EXEC sp_executesql @sql;
END
ELSE
    PRINT N'M_Employee に削除すべきインデックスはありませんでした。';
GO

-- employee_id カラムを削除
IF EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS
           WHERE TABLE_SCHEMA = 'IRAI' AND TABLE_NAME = 'M_Employee' AND COLUMN_NAME = 'employee_id')
    ALTER TABLE IRAI.M_Employee DROP COLUMN employee_id;
GO

-- employee_code を PRIMARY KEY に設定
IF NOT EXISTS (SELECT 1 FROM sys.key_constraints
               WHERE type = 'PK' AND parent_object_id = OBJECT_ID('IRAI.M_Employee'))
    ALTER TABLE IRAI.M_Employee ADD CONSTRAINT PK_Employee PRIMARY KEY (employee_code);
GO

PRINT N'Employee DB の変更が完了しました。';
GO
