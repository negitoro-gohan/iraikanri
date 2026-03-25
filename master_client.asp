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
' 取引先一覧画面（参照のみ）
' ============================================

Dim conn, rs, sql
Dim filterKeyword
filterKeyword = Trim(Request.QueryString("keyword") & "")

Set conn = GetClientDBConnection()

' 取引先一覧取得（キーワードフィルタ・案件数付き）
sql = "SELECT c.*, (SELECT COUNT(*) FROM IRAI.M_Project p WHERE p.client_code = c.client_code) AS project_count FROM IRAI.M_Client c"
If filterKeyword <> "" Then
    sql = sql & " WHERE c.client_code LIKE N'%" & EscapeSQL(filterKeyword) & "%' OR c.client_name LIKE N'%" & EscapeSQL(filterKeyword) & "%'"
End If
sql = sql & " ORDER BY c.client_code"
Set rs = conn.Execute(sql)
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
                    <li><a href="export.asp">エクスポート</a></li>
                    <li><a href="import.asp">インポート</a></li>
                </ul>
            </nav>
        </div>
    </header>
    <main class="main-content">
        <div class="container">
            <%= GetMessage() %>

<h2 class="page-title">取引先一覧</h2>

<!-- 絞り込み -->
<div class="card">
    <h3 class="card-title">絞り込み</h3>
    <form method="get" action="master_client.asp">
        <div class="form-group">
            <label>キーワード</label>
            <div style="display:flex;flex-direction:row;align-items:center;gap:10px;">
                <input type="text" name="keyword" class="form-control" style="width:300px;"
                       placeholder="取引先コードまたは取引先名"
                       value="<%= HtmlEncode(filterKeyword) %>">
                <button type="submit" class="btn btn-secondary">検索</button>
                <% If filterKeyword <> "" Then %>
                <a href="master_client.asp" class="btn btn-secondary">クリア</a>
                <% End If %>
            </div>
        </div>
    </form>
</div>

<!-- 取引先一覧 -->
<div class="card" style="padding:0; overflow-x:auto;">
    <table class="data-table">
        <thead>
            <tr>
                <th>取引先コード</th>
                <th>取引先名</th>
                <th>案件数</th>
                <th>有効</th>
            </tr>
        </thead>
        <tbody>
        <% If rs.EOF Then %>
            <tr><td colspan="4" style="text-align:center; padding:20px;">取引先が登録されていません。</td></tr>
        <% Else %>
            <% Do While Not rs.EOF %>
            <tr>
                <td><%= HtmlEncode(rs("client_code")) %></td>
                <td><%= HtmlEncode(rs("client_name")) %></td>
                <td><%= rs("project_count") %></td>
                <td><% If rs("is_active") = 1 Then %><span style="color:green;">有効</span><% Else %><span style="color:gray;">無効</span><% End If %></td>
            </tr>
            <% rs.MoveNext : Loop %>
        <% End If %>
        </tbody>
    </table>
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
