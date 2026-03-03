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
' 依頼詳細画面
' ============================================

Dim conn, rs, sql
Dim requestId
requestId = SafeInt(Request.QueryString("id"), 0)

If requestId = 0 Then
    SetMessage "error", "依頼IDが指定されていません。"
    Response.Redirect "request_list.asp"
    Response.End
End If

Set conn = GetDBConnection()

' 依頼データ取得
sql = "SELECT r.*, req.employee_name AS requester_name, req.email AS requester_email, " & _
      "ass.employee_name AS assignee_name, ass.email AS assignee_email, " & _
      "c.client_code, c.client_name, p.project_code, p.project_name " & _
      "FROM T_Request r " & _
      "INNER JOIN M_Employee req ON r.requester_id = req.employee_id " & _
      "INNER JOIN M_Employee ass ON r.assignee_id = ass.employee_id " & _
      "LEFT JOIN M_Client c ON r.client_id = c.client_id " & _
      "LEFT JOIN M_Project p ON r.project_id = p.project_id " & _
      "WHERE r.request_id = " & requestId & " AND r.is_deleted = 0"
Set rs = conn.Execute(sql)

If rs.EOF Then
    CloseRecordset rs
    CloseDBConnection conn
    SetMessage "error", "指定された依頼が見つかりません。"
    Response.Redirect "request_list.asp"
    Response.End
End If

Dim dlClass
dlClass = GetDeadlineClass(rs("deadline_date"), rs("status_id"))
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
            <h1 class="site-title"><a href="index.asp"><%= APP_TITLE %></a></h1>
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

<h2 class="page-title">依頼詳細</h2>

<div class="card">
    <table class="detail-table">
        <tr>
            <th>依頼ID</th>
            <td><%= rs("request_id") %></td>
        </tr>
        <tr>
            <th>依頼件名</th>
            <td><%= HtmlEncode(rs("request_title")) %></td>
        </tr>
        <tr>
            <th>依頼元</th>
            <td>
                <%= HtmlEncode(rs("requester_name")) %>
                <% If Not IsNull(rs("requester_email")) And rs("requester_email") <> "" Then %>
                (<%= HtmlEncode(rs("requester_email")) %>)
                <% End If %>
            </td>
        </tr>
        <tr>
            <th>依頼先</th>
            <td>
                <%= HtmlEncode(rs("assignee_name")) %>
                <% If Not IsNull(rs("assignee_email")) And rs("assignee_email") <> "" Then %>
                (<%= HtmlEncode(rs("assignee_email")) %>)
                <% End If %>
            </td>
        </tr>
        <tr>
            <th>取引先</th>
            <td>
            <% If Not IsNull(rs("client_code")) And rs("client_code") <> "" Then %>
                <%= HtmlEncode(rs("client_code")) %> : <%= HtmlEncode(rs("client_name")) %>
            <% Else %>
                -
            <% End If %>
            </td>
        </tr>
        <tr>
            <th>案件</th>
            <td>
            <% If Not IsNull(rs("project_code")) And rs("project_code") <> "" Then %>
                <%= HtmlEncode(rs("project_code")) %> : <%= HtmlEncode(rs("project_name")) %>
            <% Else %>
                -
            <% End If %>
            </td>
        </tr>
        <tr>
            <th>期限日</th>
            <td class="<%= dlClass %>"><%= FormatDateJP(rs("deadline_date")) %></td>
        </tr>
        <tr>
            <th>ステータス</th>
            <td><span class="status-badge <%= GetStatusClass(rs("status_id")) %>"><%= GetStatusName(rs("status_id")) %></span></td>
        </tr>
        <tr>
            <th>依頼内容</th>
            <td><%= NL2BR(HtmlEncode(SafeValue(rs("request_content"), "-"))) %></td>
        </tr>
        <tr>
            <th>対応内容</th>
            <td><%= NL2BR(HtmlEncode(SafeValue(rs("response_content"), "-"))) %></td>
        </tr>
        <tr>
            <th>終了日</th>
            <td><%= SafeValue(FormatDateJP(rs("end_date")), "-") %></td>
        </tr>
        <tr>
            <th>登録日時</th>
            <td><%= rs("created_at") %></td>
        </tr>
        <tr>
            <th>更新日時</th>
            <td><%= rs("updated_at") %></td>
        </tr>
    </table>

    <div class="btn-group">
        <a href="request_edit.asp?id=<%= requestId %>" class="btn btn-warning">編集</a>
        <a href="request_response.asp?id=<%= requestId %>" class="btn btn-success">対応入力</a>
        <a href="request_delete.asp?id=<%= requestId %>" class="btn btn-danger" onclick="return confirmDelete('この依頼を削除しますか？');">削除</a>
        <a href="request_list.asp" class="btn btn-secondary">一覧に戻る</a>
    </div>
</div>

<%
CloseRecordset rs
CloseDBConnection conn
%>
        </div>
    </main>
    <script src="js/common.js"></script>
</body>
</html>
