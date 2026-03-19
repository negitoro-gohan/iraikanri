<%@ Language="VBScript" CodePage=65001 %>
<%
Response.CodePage = 65001
Session.CodePage = 65001
Response.Charset = "UTF-8"
Response.ContentType = "text/html; charset=UTF-8"
%>
<!--#include file="include/config.asp"-->
<!--#include file="include/db_connection.asp"-->
<!--#include file="include/functions.asp"-->
<%
' ============================================
' ダッシュボード（トップページ）
' 依頼事項の概況を表示
' ============================================

Dim conn, rs, sql
Dim countOverdue, countSoon, countInProgress, countTotal
Dim currentUser, myEmployeeId

Set conn = GetDBConnection()

' ログインユーザーの employee_id を取得（Employee DBへ接続）
' employee_code = ドメインユーザー名（DOMAIN\ 除去後）で照合する
currentUser = GetCurrentUser()
myEmployeeId = 0
Dim connEmpMe, rsEmpMe
Set connEmpMe = GetEmployeeDBConnection()
Set rsEmpMe = connEmpMe.Execute("SELECT employee_id FROM IRAI.M_Employee WHERE employee_code = N'" & EscapeSQL(currentUser) & "' AND is_active = 1")
If Not rsEmpMe.EOF Then myEmployeeId = CLng(rsEmpMe("employee_id"))
CloseRecordset rsEmpMe
CloseDBConnection connEmpMe

' ユーザーフィルタ（依頼先 = ログインユーザー）
Dim userFilter
If myEmployeeId > 0 Then
    userFilter = "is_deleted = 0 AND assignee_id = " & myEmployeeId
Else
    userFilter = "is_deleted = 0 AND 1 = 0" ' 社員マスターに未登録のため0件扱い
End If

' 期限切れ件数（未完了・自分が依頼先）
sql = "SELECT COUNT(*) FROM IRAI.T_Request WHERE " & userFilter & " AND status_id NOT IN (" & STATUS_COMPLETED & "," & STATUS_NOT_APPLICABLE & ") AND deadline_date < CONVERT(DATE, GETDATE())"
Set rs = conn.Execute(sql)
countOverdue = rs(0)
CloseRecordset rs

' 期限間近件数（未完了で期限が今日から3日以内・自分が依頼先）
sql = "SELECT COUNT(*) FROM IRAI.T_Request WHERE " & userFilter & " AND status_id NOT IN (" & STATUS_COMPLETED & "," & STATUS_NOT_APPLICABLE & ") AND deadline_date >= CONVERT(DATE, GETDATE()) AND deadline_date <= DATEADD(day, " & REMIND_DAYS_BEFORE & ", CONVERT(DATE, GETDATE()))"
Set rs = conn.Execute(sql)
countSoon = rs(0)
CloseRecordset rs

' 着手済件数（自分が依頼先）
sql = "SELECT COUNT(*) FROM IRAI.T_Request WHERE " & userFilter & " AND status_id = " & STATUS_IN_PROGRESS
Set rs = conn.Execute(sql)
countInProgress = rs(0)
CloseRecordset rs

' 全件数（自分が依頼先）
sql = "SELECT COUNT(*) FROM IRAI.T_Request WHERE " & userFilter
Set rs = conn.Execute(sql)
countTotal = rs(0)
CloseRecordset rs
%>
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title><%= APP_TITLE %></title>
    <link rel="stylesheet" href="css/style.css">
</head>
<body>
    <header class="site-header">
        <div class="header-container">
            <h1 class="site-title">
                <a href="index.asp"><%= APP_TITLE %></a>
            </h1>
            <nav class="main-nav">
                <ul>
                    <li><a href="index.asp">ホーム</a></li>
                    <li><a href="request_list.asp">依頼一覧</a></li>
                    <li><a href="request_new.asp">新規登録</a></li>
                    <li><a href="master_employee.asp">社員管理</a></li>
                    <li><a href="master_client.asp">取引先管理</a></li>
                    <li><a href="master_project.asp">案件管理</a></li>
                    <li><a href="export.asp">エクスポート</a></li>
                    <li><a href="import.asp">インポート</a></li>
                </ul>
            </nav>
        </div>
    </header>
    <main class="main-content">
        <div class="container">
            <%= GetMessage() %>

<h2 class="page-title">ダッシュボード <small style="font-size:0.6em;font-weight:normal;color:#666;">（<%= HtmlEncode(currentUser) %> の依頼）</small></h2>

<% If myEmployeeId = 0 Then %>
<div class="message message-warning">社員マスターに login_name が設定されていないため、件数を表示できません。管理者に連絡してください。</div>
<% End If %>

<!-- サマリーカード -->
<div class="dashboard-cards">
    <div class="dashboard-card overdue">
        <div class="count"><%= countOverdue %></div>
        <div class="label">期限切れ</div>
        <a href="request_list.asp?filter=overdue&my_assignee=1" class="btn btn-sm btn-danger">一覧を見る</a>
    </div>
    <div class="dashboard-card soon">
        <div class="count"><%= countSoon %></div>
        <div class="label">期限間近（<%= REMIND_DAYS_BEFORE %>日以内）</div>
        <a href="request_list.asp?filter=soon&my_assignee=1" class="btn btn-sm btn-warning">一覧を見る</a>
    </div>
    <div class="dashboard-card in-progress">
        <div class="count"><%= countInProgress %></div>
        <div class="label">着手済</div>
        <a href="request_list.asp?filter=in_progress&my_assignee=1" class="btn btn-sm btn-primary">一覧を見る</a>
    </div>
    <div class="dashboard-card">
        <div class="count"><%= countTotal %></div>
        <div class="label">全依頼件数</div>
        <a href="request_list.asp?my_assignee=1" class="btn btn-sm btn-secondary">一覧を見る</a>
    </div>
</div>

<%
CloseDBConnection conn
%>
        </div>
    </main>
    <script src="js/common.js"></script>
</body>
</html>
