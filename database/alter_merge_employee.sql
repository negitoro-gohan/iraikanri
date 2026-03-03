-- ============================================
-- M_Requester / M_Assignee を M_Employee に統合する
-- マイグレーションスクリプト
-- ============================================

USE IraiKanri;
GO

-- 1. 社員マスターテーブル作成
CREATE TABLE M_Employee (
    employee_id INT IDENTITY(1,1) PRIMARY KEY,   -- 社員ID
    employee_code NVARCHAR(20) NOT NULL,          -- 社員コード
    employee_name NVARCHAR(100) NOT NULL,         -- 社員名
    email NVARCHAR(256),                          -- メールアドレス
    is_active BIT DEFAULT 1,                      -- 有効フラグ
    created_at DATETIME DEFAULT GETDATE(),        -- 作成日時
    updated_at DATETIME DEFAULT GETDATE()         -- 更新日時
);
GO

-- 社員コードにユニークインデックス
CREATE UNIQUE INDEX IX_Employee_Code ON M_Employee(employee_code);
GO

-- 2. 既存データを移行（依頼元 → 社員）
INSERT INTO M_Employee (employee_code, employee_name, email, is_active, created_at, updated_at)
SELECT
    COALESCE(employee_code, N'REQ' + CAST(requester_id AS NVARCHAR(10))),
    requester_name, email, is_active, created_at, updated_at
FROM M_Requester;

-- 依頼先で、依頼元と名前・メールが重複しないもののみ追加
INSERT INTO M_Employee (employee_code, employee_name, email, is_active, created_at, updated_at)
SELECT
    COALESCE(a.employee_code, N'ASG' + CAST(a.assignee_id AS NVARCHAR(10))),
    a.assignee_name, a.email, a.is_active, a.created_at, a.updated_at
FROM M_Assignee a
WHERE NOT EXISTS (
    SELECT 1 FROM M_Requester r WHERE r.requester_name = a.assignee_name
);
GO

-- 3. T_Request の外部キーを変更
-- 既存の外部キー制約を削除
ALTER TABLE T_Request DROP CONSTRAINT FK_Request_Requester;
ALTER TABLE T_Request DROP CONSTRAINT FK_Request_Assignee;
GO

-- requester_id を新しい employee_id にマッピング
UPDATE t SET t.requester_id = e.employee_id
FROM T_Request t
INNER JOIN M_Requester r ON t.requester_id = r.requester_id
INNER JOIN M_Employee e ON e.employee_name = r.requester_name;

UPDATE t SET t.assignee_id = e.employee_id
FROM T_Request t
INNER JOIN M_Assignee a ON t.assignee_id = a.assignee_id
INNER JOIN M_Employee e ON e.employee_name = a.assignee_name;
GO

-- 新しい外部キー制約を追加
ALTER TABLE T_Request ADD CONSTRAINT FK_Request_Requester FOREIGN KEY (requester_id) REFERENCES M_Employee(employee_id);
ALTER TABLE T_Request ADD CONSTRAINT FK_Request_Assignee FOREIGN KEY (assignee_id) REFERENCES M_Employee(employee_id);
GO

-- 4. 旧テーブルを削除
DROP TABLE M_Requester;
DROP TABLE M_Assignee;
GO

PRINT N'M_Employee 統合が完了しました。';
GO
