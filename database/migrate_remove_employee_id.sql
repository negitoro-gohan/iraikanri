-- ============================================
-- employee_id 廃止マイグレーション
--
-- 変更内容:
--   1. IraiKanri DB: T_Request の requester_id/assignee_id (INT) を削除し
--      requester_code/assignee_code (NVARCHAR) を追加する
--   2. Employee DB: M_Employee の employee_id を削除し
--      employee_code を PRIMARY KEY にする
--
-- 注意: 手順3のデータ移行は IraiKanri DB と Employee DB をまたぐため、
--       あらかじめ Employee DB で employee_id→employee_code のマッピングを
--       確認してから実行してください。
-- ============================================

-- ============================================
-- 手順1: IraiKanri DB に新カラム追加
-- ============================================
USE IraiKanri;
GO

-- FK制約を削除（存在する場合）
IF EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = 'FK_Request_Requester')
    ALTER TABLE IRAI.T_Request DROP CONSTRAINT FK_Request_Requester;
GO
IF EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = 'FK_Request_Assignee')
    ALTER TABLE IRAI.T_Request DROP CONSTRAINT FK_Request_Assignee;
GO

-- 新カラム追加
ALTER TABLE IRAI.T_Request ADD requester_code NVARCHAR(20) NULL;
GO
ALTER TABLE IRAI.T_Request ADD assignee_code  NVARCHAR(20) NULL;
GO

-- ============================================
-- 手順2: 既存データのコード値を移行する
--
-- Employee DB の employee_id→employee_code 対応を確認したうえで
-- 以下の UPDATE を繰り返し実行してください。
-- 例)
--   UPDATE IRAI.T_Request SET requester_code = 'yamada'  WHERE requester_id = 1;
--   UPDATE IRAI.T_Request SET requester_code = 'suzuki'  WHERE requester_id = 2;
--   UPDATE IRAI.T_Request SET assignee_code  = 'sato'    WHERE assignee_id  = 3;
-- ============================================

-- ============================================
-- 手順3: 旧カラム削除（手順2完了後に実行）
-- ============================================
ALTER TABLE IRAI.T_Request DROP COLUMN requester_id;
GO
ALTER TABLE IRAI.T_Request DROP COLUMN assignee_id;
GO

PRINT N'IraiKanri DB の変更が完了しました。';
GO

-- ============================================
-- 手順4: Employee DB の employee_id を削除し
--         employee_code を PRIMARY KEY にする
-- ============================================
USE Employee;
GO

-- employee_id が PK の場合、制約を削除
DECLARE @pkName NVARCHAR(200);
SELECT @pkName = name FROM sys.key_constraints
WHERE type = 'PK' AND parent_object_id = OBJECT_ID('IRAI.M_Employee');
IF @pkName IS NOT NULL
    EXEC('ALTER TABLE IRAI.M_Employee DROP CONSTRAINT ' + @pkName);
GO

-- employee_id カラムを削除
IF EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS
           WHERE TABLE_SCHEMA = 'IRAI' AND TABLE_NAME = 'M_Employee' AND COLUMN_NAME = 'employee_id')
    ALTER TABLE IRAI.M_Employee DROP COLUMN employee_id;
GO

-- employee_code の UNIQUE INDEX を削除し PRIMARY KEY に変更
IF EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_Employee_Code' AND object_id = OBJECT_ID('IRAI.M_Employee'))
    DROP INDEX IX_Employee_Code ON IRAI.M_Employee;
GO

ALTER TABLE IRAI.M_Employee ADD CONSTRAINT PK_Employee PRIMARY KEY (employee_code);
GO

PRINT N'Employee DB の変更が完了しました。';
GO
