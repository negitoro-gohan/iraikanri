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
' 取引先マスター管理画面
' ============================================

Dim conn, rs, sql
Dim errorMsg, mode, clientId, filterKeyword
errorMsg = ""
mode = Trim(Request.QueryString("mode") & "")
clientId = SafeInt(Request.QueryString("id"), 0)
filterKeyword = Trim(Request.QueryString("keyword") & "")

Set conn = GetDBConnection()

' POST処理
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim action, postClientId, clientCode, clientName, isActive
    action = Trim(Request.Form("action") & "")
    postClientId = SafeInt(Request.Form("client_id"), 0)
    clientCode = Trim(Request.Form("client_code") & "")
    clientName = Trim(Request.Form("client_name") & "")
    isActive = SafeInt(Request.Form("is_active"), 1)

    If action = "add" Then
        ' 新規登録
        If clientCode = "" Then
            errorMsg = "取引先コードを入力してください。"
        ElseIf clientName = "" Then
            errorMsg = "取引先名を入力してください。"
        Else
            ' 重複チェック
            sql = "SELECT COUNT(*) FROM M_Client WHERE client_code = N'" & EscapeSQL(clientCode) & "'"
            Set rs = conn.Execute(sql)
            If rs(0) > 0 Then
                errorMsg = "この取引先コードは既に登録されています。"
            Else
                sql = "INSERT INTO M_Client (client_code, client_name, is_active) VALUES (N'" & EscapeSQL(clientCode) & "', N'" & EscapeSQL(clientName) & "', " & isActive & ")"
                On Error Resume Next
                conn.Execute sql
                If Err.Number <> 0 Then
                    errorMsg = "登録に失敗しました：" & Err.Description
                    Err.Clear
                Else
                    SetMessage "success", "取引先を登録しました。"
                    Response.Redirect "master_client.asp"
                    Response.End
                End If
                On Error GoTo 0
            End If
            CloseRecordset rs
        End If

    ElseIf action = "edit" And postClientId > 0 Then
        ' 更新
        If clientCode = "" Then
            errorMsg = "取引先コードを入力してください。"
        ElseIf clientName = "" Then
            errorMsg = "取引先名を入力してください。"
        Else
            ' 重複チェック（自分以外）
            sql = "SELECT COUNT(*) FROM M_Client WHERE client_code = N'" & EscapeSQL(clientCode) & "' AND client_id <> " & postClientId
            Set rs = conn.Execute(sql)
            If rs(0) > 0 Then
                errorMsg = "この取引先コードは既に使用されています。"
            Else
                sql = "UPDATE M_Client SET client_code = N'" & EscapeSQL(clientCode) & "', client_name = N'" & EscapeSQL(clientName) & "', is_active = " & isActive & ", updated_at = GETDATE() WHERE client_id = " & postClientId
                On Error Resume Next
                conn.Execute sql
                If Err.Number <> 0 Then
                    errorMsg = "更新に失敗しました：" & Err.Description
                    Err.Clear
                Else
                    SetMessage "success", "取引先を更新しました。"
                    Response.Redirect "master_client.asp"
                    Response.End
                End If
                On Error GoTo 0
            End If
            CloseRecordset rs
        End If

    ElseIf action = "delete" And postClientId > 0 Then
        ' 削除（案件が紐づいている場合は削除不可）
        sql = "SELECT COUNT(*) FROM M_Project WHERE client_id = " & postClientId
        Set rs = conn.Execute(sql)
        If rs(0) > 0 Then
            errorMsg = "この取引先には案件が登録されているため削除できません。"
        Else
            sql = "DELETE FROM M_Client WHERE client_id = " & postClientId
            On Error Resume Next
            conn.Execute sql
            If Err.Number <> 0 Then
                errorMsg = "削除に失敗しました：" & Err.Description
                Err.Clear
            Else
                SetMessage "success", "取引先を削除しました。"
                Response.Redirect "master_client.asp"
                Response.End
            End If
            On Error GoTo 0
        End If
        CloseRecordset rs
    End If
End If

' 編集モードの場合、データ取得
Dim editClient
If mode = "edit" And clientId > 0 Then
    sql = "SELECT * FROM M_Client WHERE client_id = " & clientId
    Set editClient = conn.Execute(sql)
    If editClient.EOF Then
        mode = ""
        clientId = 0
    End If
End If

' 取引先一覧取得（キーワードフィルタ）
sql = "SELECT c.*, (SELECT COUNT(*) FROM M_Project p WHERE p.client_id = c.client_id) AS project_count FROM M_Client c"
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

<h2 class="page-title">取引先マスター管理</h2>

<% If errorMsg <> "" Then %>
<div class="message message-error"><%= HtmlEncode(errorMsg) %></div>
<% End If %>

<!-- 登録・編集フォーム -->
<div class="card">
    <h3 class="card-title"><% If mode = "edit" Then %>取引先編集<% Else %>取引先新規登録<% End If %></h3>
    <form method="post" action="master_client.asp">
        <input type="hidden" name="action" value="<% If mode = "edit" Then %>edit<% Else %>add<% End If %>">
        <% If mode = "edit" Then %>
        <input type="hidden" name="client_id" value="<%= clientId %>">
        <% End If %>

        <div class="form-group">
            <label>取引先コード<span class="required">*</span></label>
            <input type="text" name="client_code" class="form-control" maxlength="50" required style="width:200px;"
                   value="<% If mode = "edit" Then %><%= HtmlEncode(editClient("client_code") & "") %><% Else %><%= HtmlEncode(Request.Form("client_code") & "") %><% End If %>">
        </div>

        <div class="form-group">
            <label>取引先名<span class="required">*</span></label>
            <input type="text" name="client_name" class="form-control" maxlength="200" required
                   value="<% If mode = "edit" Then %><%= HtmlEncode(editClient("client_name") & "") %><% Else %><%= HtmlEncode(Request.Form("client_name") & "") %><% End If %>">
        </div>

        <div class="form-group">
            <label>有効</label>
            <select name="is_active" class="form-control" style="width:100px;">
                <%
                Dim editIsActive
                If mode = "edit" Then
                    editIsActive = editClient("is_active")
                Else
                    editIsActive = 1
                End If
                %>
                <option value="1" <% If editIsActive = 1 Then %>selected<% End If %>>有効</option>
                <option value="0" <% If editIsActive = 0 Then %>selected<% End If %>>無効</option>
            </select>
        </div>

        <div class="btn-group">
            <button type="submit" class="btn btn-primary"><% If mode = "edit" Then %>更新<% Else %>登録<% End If %></button>
            <% If mode = "edit" Then %>
            <a href="master_client.asp" class="btn btn-secondary">キャンセル</a>
            <% End If %>
        </div>
    </form>
</div>

<!-- 絞り込み -->
<div class="card">
    <h3 class="card-title">絞り込み</h3>
    <form method="get" action="master_client.asp">
        <div class="form-group">
            <label>キーワード</label>
            <div class="employee-code-group" style="display:flex;flex-direction:row;flex-wrap:nowrap;align-items:center;gap:10px;">
                <input type="text" name="keyword" id="filterKeyword" class="form-control" style="width:300px;min-width:200px;max-width:300px;flex:0 0 300px;"
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
                <th>ID</th>
                <th>取引先コード</th>
                <th>取引先名</th>
                <th>案件数</th>
                <th>有効</th>
                <th>操作</th>
            </tr>
        </thead>
        <tbody>
        <% If rs.EOF Then %>
            <tr><td colspan="6" style="text-align:center; padding:20px;">取引先が登録されていません。</td></tr>
        <% Else %>
            <% Do While Not rs.EOF %>
            <tr>
                <td><%= rs("client_id") %></td>
                <td><%= HtmlEncode(rs("client_code")) %></td>
                <td><%= HtmlEncode(rs("client_name")) %></td>
                <td><%= rs("project_count") %></td>
                <td><% If rs("is_active") = 1 Then %><span style="color:green;">有効</span><% Else %><span style="color:gray;">無効</span><% End If %></td>
                <td class="action-links">
                    <a href="master_client.asp?mode=edit&id=<%= rs("client_id") %>" class="btn btn-sm btn-warning">編集</a>
                    <a href="master_project.asp?client_id=<%= rs("client_id") %>" class="btn btn-sm btn-primary">案件</a>
                    <form method="post" action="master_client.asp" style="display:inline;">
                        <input type="hidden" name="action" value="delete">
                        <input type="hidden" name="client_id" value="<%= rs("client_id") %>">
                        <button type="submit" class="btn btn-sm btn-danger" onclick="return confirm('この取引先を削除しますか？');">削除</button>
                    </form>
                </td>
            </tr>
            <% rs.MoveNext : Loop %>
        <% End If %>
        </tbody>
    </table>
</div>

<%
If IsObject(editClient) Then CloseRecordset editClient
CloseRecordset rs
CloseDBConnection conn
%>
        </div>
    </main>
    <script src="js/common.js"></script>
</body>
</html>
