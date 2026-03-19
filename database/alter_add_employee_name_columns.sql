-- ============================================
-- T_Request に社員名カラムを追加
--
-- 背景:
--   社員情報を別DB（Employee）で管理するようになったため、
--   JOINで社員名を取得するのではなく、依頼登録・編集時に
--   名前をT_Requestに保存するように変更する。
-- ============================================

USE IraiKanri;
GO

ALTER TABLE IRAI.T_Request ADD requester_name NVARCHAR(100) NULL; -- 依頼元社員名
GO

ALTER TABLE IRAI.T_Request ADD assignee_name  NVARCHAR(100) NULL; -- 依頼先社員名
GO

-- 既存レコードの名前は別途バッチで更新するか、NULLのままとする
-- （画面表示時は NULL の場合に「（不明）」等の代替文字列を表示するよう各画面で対応済み）
