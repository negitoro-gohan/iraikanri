-- ============================================
-- 依頼事項 期限管理WEBアプリ
-- データベース作成スクリプト
-- ============================================

-- データベース作成
CREATE DATABASE IraiKanri;
GO

USE IraiKanri;
GO

-- ============================================
-- マスターテーブル
-- ============================================

-- 社員マスター（依頼元・依頼先の共通マスター）
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

-- ステータスマスター
CREATE TABLE M_Status (
    status_id INT PRIMARY KEY,                   -- ステータスID
    status_name NVARCHAR(50) NOT NULL,           -- ステータス名
    display_order INT DEFAULT 0                  -- 表示順
);
GO

-- ステータス初期データ
INSERT INTO M_Status (status_id, status_name, display_order) VALUES
(1, N'未着手', 1),
(2, N'着手済', 2),
(3, N'対応完了', 3),
(4, N'対象外', 4);
GO

-- ============================================
-- トランザクションテーブル
-- ============================================

-- 依頼事項テーブル
CREATE TABLE T_Request (
    request_id INT IDENTITY(1,1) PRIMARY KEY,    -- 依頼ID
    requester_id INT NOT NULL,                   -- 依頼元ID
    assignee_id INT NOT NULL,                    -- 依頼先ID
    request_title NVARCHAR(200) NOT NULL,        -- 依頼件名
    request_content NVARCHAR(MAX),               -- 依頼内容
    deadline_date DATE NOT NULL,                 -- 期限日
    end_date DATE,                               -- 終了日
    status_id INT DEFAULT 1,                     -- ステータスID
    response_content NVARCHAR(MAX),              -- 対応内容
    is_deleted BIT DEFAULT 0,                    -- 削除フラグ
    created_by NVARCHAR(100),                    -- 登録者（ドメインユーザー名）
    updated_by NVARCHAR(100),                    -- 更新者（ドメインユーザー名）
    created_at DATETIME DEFAULT GETDATE(),       -- 作成日時
    updated_at DATETIME DEFAULT GETDATE(),       -- 更新日時

    -- 外部キー制約（どちらも社員マスターを参照）
    CONSTRAINT FK_Request_Requester FOREIGN KEY (requester_id) REFERENCES M_Employee(employee_id),
    CONSTRAINT FK_Request_Assignee FOREIGN KEY (assignee_id) REFERENCES M_Employee(employee_id),
    CONSTRAINT FK_Request_Status FOREIGN KEY (status_id) REFERENCES M_Status(status_id)
);
GO

-- メール送信履歴テーブル
CREATE TABLE T_MailLog (
    mail_log_id INT IDENTITY(1,1) PRIMARY KEY,   -- メールログID
    request_id INT NOT NULL,                     -- 依頼ID
    mail_type INT NOT NULL,                      -- メール種別（1:新規、2:リマインド、3:期限切れ）
    to_email NVARCHAR(256) NOT NULL,             -- 送信先
    subject NVARCHAR(200) NOT NULL,              -- 件名
    body NVARCHAR(MAX),                          -- 本文
    sent_at DATETIME DEFAULT GETDATE(),          -- 送信日時
    is_success BIT DEFAULT 1,                    -- 送信成功フラグ
    error_message NVARCHAR(500),                 -- エラーメッセージ

    CONSTRAINT FK_MailLog_Request FOREIGN KEY (request_id) REFERENCES T_Request(request_id)
);
GO

-- ============================================
-- インデックス
-- ============================================

-- 依頼事項の検索用インデックス
CREATE INDEX IX_Request_Deadline ON T_Request(deadline_date);
CREATE INDEX IX_Request_Status ON T_Request(status_id);
CREATE INDEX IX_Request_Requester ON T_Request(requester_id);
CREATE INDEX IX_Request_Assignee ON T_Request(assignee_id);
CREATE INDEX IX_Request_IsDeleted ON T_Request(is_deleted);
GO

-- ============================================
-- サンプルデータ（開発用）
-- ============================================

-- 社員サンプルデータ
INSERT INTO M_Employee (employee_code, employee_name, email) VALUES
(N'EMP001', N'山田太郎', 'yamada@example.com'),
(N'EMP002', N'鈴木花子', 'suzuki@example.com'),
(N'EMP003', N'佐藤一郎', 'sato@example.com'),
(N'EMP004', N'田中次郎', 'tanaka@example.com'),
(N'EMP005', N'高橋美咲', 'takahashi@example.com'),
(N'EMP006', N'伊藤健太', 'ito@example.com');
GO

PRINT N'データベース作成が完了しました。';
GO
