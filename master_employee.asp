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
' 社員一覧画面（参照のみ）
' ============================================

Dim conn, rs, sql
Set conn = GetEmployeeDBConnection()

sql = "SELECT * FROM IRAI.M_Employee ORDER BY employee_code"
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

<h2 class="page-title">社員一覧</h2>

<div class="card" style="padding: 0;">
    <table class="data-table">
        <thead>
            <tr>
                <th>社員コード</th>
                <th>社員名</th>
                <th>メールアドレス</th>
                <th>有効</th>
                <th>登録日時</th>
            </tr>
        </thead>
        <tbody>
        <%
        If rs.EOF Then
        %>
            <tr>
                <td colspan="5" style="text-align: center; padding: 30px;">登録されている社員がありません。</td>
            </tr>
        <%
        Else
            Do While Not rs.EOF
        %>
            <tr>
                <td><%= HtmlEncode(rs("employee_code")) %></td>
                <td><%= HtmlEncode(rs("employee_name")) %></td>
                <td><%= HtmlEncode(SafeValue(rs("email"), "-")) %></td>
                <td><% If rs("is_active") = 1 Then %><span style="color:green;">有効</span><% Else %><span style="color:gray;">無効</span><% End If %></td>
                <td><%= FormatDateJP(rs("created_at")) %></td>
            </tr>
        <%
                rs.MoveNext
            Loop
        End If
        CloseRecordset rs
        %>
        </tbody>
    </table>
</div>

<%
CloseDBConnection conn
%>
        </div>
    </main>
    <script src="js/common.js"></script>
</body>
</html>
