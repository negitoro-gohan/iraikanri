-- ============================================
-- T_Request テーブルに登録者・更新者カラムを追加
-- ============================================

USE IraiKanri;
GO

-- 登録者（ドメインユーザー名）カラムを追加
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'T_Request' AND COLUMN_NAME = 'created_by')
BEGIN
    ALTER TABLE T_Request ADD created_by NVARCHAR(100) NULL;  -- 登録者（ドメインユーザー名）
    PRINT N'created_by カラムを追加しました。';
END
ELSE
BEGIN
    PRINT N'created_by カラムは既に存在します。';
END
GO

-- 更新者（ドメインユーザー名）カラムを追加
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'T_Request' AND COLUMN_NAME = 'updated_by')
BEGIN
    ALTER TABLE T_Request ADD updated_by NVARCHAR(100) NULL;  -- 更新者（ドメインユーザー名）
    PRINT N'updated_by カラムを追加しました。';
END
ELSE
BEGIN
    PRINT N'updated_by カラムは既に存在します。';
END
GO

PRINT N'登録者・更新者カラムの追加が完了しました。';
GO
