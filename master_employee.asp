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
' 社員マスター管理画面
' ============================================

' IIf関数
Function IIf(condition, trueValue, falseValue)
    If condition Then
        IIf = trueValue
    Else
        IIf = falseValue
    End If
End Function

Dim conn, rs, sql
Dim errorMsg, mode, editId
errorMsg = ""
mode = Request.QueryString("mode") & ""
If mode = "" Then mode = "list"
editId = SafeInt(Request.QueryString("id"), 0)

Set conn = GetEmployeeDBConnection()

' POST処理
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim action, id, employeeCode, employeeName, email, isActive

    action = Request.Form("action") & ""
    id = SafeInt(Request.Form("id"), 0)
    employeeCode = Trim(Request.Form("employee_code") & "")
    employeeName = Trim(Request.Form("employee_name") & "")
    email = Trim(Request.Form("email") & "")
    isActive = SafeInt(Request.Form("is_active"), 1)

    Select Case action
        Case "add"
            If employeeCode = "" Then
                errorMsg = "社員コードを入力してください。"
                mode = "add"
            ElseIf employeeName = "" Then
                errorMsg = "社員名を入力してください。"
                mode = "add"
            Else
                sql = "INSERT INTO IRAI.M_Employee (employee_code, employee_name, email, is_active) VALUES (N'" & EscapeSQL(employeeCode) & "', N'" & EscapeSQL(employeeName) & "', N'" & EscapeSQL(email) & "', " & isActive & ")"
                On Error Resume Next
                conn.Execute sql
                If Err.Number <> 0 Then
                    errorMsg = "登録に失敗しました：" & Err.Description
                    mode = "add"
                    Err.Clear
                Else
                    SetMessage "success", "社員を登録しました。"
                    Response.Redirect "master_employee.asp"
                    Response.End
                End If
                On Error GoTo 0
            End If

        Case "edit"
            If employeeCode = "" Then
                errorMsg = "社員コードを入力してください。"
                mode = "edit"
                editId = id
            ElseIf employeeName = "" Then
                errorMsg = "社員名を入力してください。"
                mode = "edit"
                editId = id
            Else
                sql = "UPDATE IRAI.M_Employee SET employee_code = N'" & EscapeSQL(employeeCode) & "', employee_name = N'" & EscapeSQL(employeeName) & "', email = N'" & EscapeSQL(email) & "', is_active = " & isActive & ", updated_at = GETDATE() WHERE employee_id = " & id
                On Error Resume Next
                conn.Execute sql
                If Err.Number <> 0 Then
                    errorMsg = "更新に失敗しました：" & Err.Description
                    mode = "edit"
                    editId = id
                    Err.Clear
                Else
                    SetMessage "success", "社員を更新しました。"
                    Response.Redirect "master_employee.asp"
                    Response.End
                End If
                On Error GoTo 0
            End If

        Case "delete"
            ' 依頼元または依頼先として使用中かチェック（依頼管理DBへ別接続）
            Dim connMain
            Set connMain = GetDBConnection()
            sql = "SELECT COUNT(*) FROM IRAI.T_Request WHERE (requester_id = " & id & " OR assignee_id = " & id & ") AND is_deleted = 0"
            Set rs = connMain.Execute(sql)
            If rs(0) > 0 Then
                SetMessage "error", "この社員は依頼で使用中のため削除できません。"
            Else
                CloseRecordset rs
                CloseDBConnection connMain
                sql = "DELETE FROM IRAI.M_Employee WHERE employee_id = " & id
                On Error Resume Next
                conn.Execute sql
                If Err.Number <> 0 Then
                    SetMessage "error", "削除に失敗しました：" & Err.Description
                    Err.Clear
                Else
                    SetMessage "success", "社員を削除しました。"
                End If
                On Error GoTo 0
            End If
            If Not rs Is Nothing Then CloseRecordset rs
            If Not connMain Is Nothing Then CloseDBConnection connMain
            Response.Redirect "master_employee.asp"
            Response.End
    End Select
End If

' 編集データ取得
Dim editData
If mode = "edit" And editId > 0 Then
    sql = "SELECT * FROM IRAI.M_Employee WHERE employee_id = " & editId
    Set editData = conn.Execute(sql)
    If editData.EOF Then
        mode = "list"
    End If
End If
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

<h2 class="page-title">社員マスター管理</h2>

<% If errorMsg <> "" Then %>
<div class="message message-error"><%= HtmlEncode(errorMsg) %></div>
<% End If %>

<% If mode = "add" Or mode = "edit" Then %>
<!-- 登録・編集フォーム -->
<div class="card">
    <h3 class="card-title"><%= IIf(mode = "add", "新規登録", "編集") %></h3>
    <form method="post" action="master_employee.asp">
        <input type="hidden" name="action" value="<%= mode %>">
        <% If mode = "edit" Then %>
        <input type="hidden" name="id" value="<%= editId %>">
        <% End If %>

        <div class="form-group">
            <label>社員コード<span class="required">*</span></label>
            <%
            Dim formCode
            If Request.Form("employee_code") <> "" Then
                formCode = Request.Form("employee_code")
            ElseIf mode = "edit" Then
                formCode = editData("employee_code") & ""
            Else
                formCode = ""
            End If
            %>
            <input type="text" name="employee_code" class="form-control" maxlength="20" required value="<%= HtmlEncode(formCode) %>">
        </div>

        <div class="form-group">
            <label>社員名<span class="required">*</span></label>
            <%
            Dim formName
            If Request.Form("employee_name") <> "" Then
                formName = Request.Form("employee_name")
            ElseIf mode = "edit" Then
                formName = editData("employee_name")
            Else
                formName = ""
            End If
            %>
            <input type="text" name="employee_name" class="form-control" maxlength="100" required value="<%= HtmlEncode(formName) %>">
        </div>

        <div class="form-group">
            <label>メールアドレス</label>
            <%
            Dim formEmail
            If Request.Form("email") <> "" Then
                formEmail = Request.Form("email")
            ElseIf mode = "edit" Then
                formEmail = editData("email") & ""
            Else
                formEmail = ""
            End If
            %>
            <input type="email" name="email" class="form-control" maxlength="256" value="<%= HtmlEncode(formEmail) %>">
        </div>

        <div class="form-group">
            <label>有効</label>
            <%
            Dim formActive
            If Request.Form("is_active") <> "" Then
                formActive = SafeInt(Request.Form("is_active"), 1)
            ElseIf mode = "edit" Then
                formActive = editData("is_active")
            Else
                formActive = 1
            End If
            %>
            <select name="is_active" class="form-control">
                <option value="1" <% If formActive = 1 Then Response.Write " selected" End If %>>有効</option>
                <option value="0" <% If formActive = 0 Then Response.Write " selected" End If %>>無効</option>
            </select>
        </div>

        <div class="btn-group">
            <button type="submit" class="btn btn-primary"><%= IIf(mode = "add", "登録する", "更新する") %></button>
            <a href="master_employee.asp" class="btn btn-secondary">キャンセル</a>
        </div>
    </form>
</div>
<%
    If mode = "edit" Then
         CloseRecordset editData
    End If
%>

<% Else %>
<!-- 一覧表示 -->
<div style="margin-bottom: 15px;">
    <a href="master_employee.asp?mode=add" class="btn btn-primary">新規登録</a>
</div>

<div class="card" style="padding: 0;">
    <table class="data-table">
        <thead>
            <tr>
                <th>ID</th>
                <th>社員コード</th>
                <th>社員名</th>
                <th>メールアドレス</th>
                <th>有効</th>
                <th>登録日時</th>
                <th>操作</th>
            </tr>
        </thead>
        <tbody>
        <%
        sql = "SELECT * FROM IRAI.M_Employee ORDER BY employee_code"
        Set rs = conn.Execute(sql)

        If rs.EOF Then
        %>
            <tr>
                <td colspan="7" style="text-align: center; padding: 30px;">登録されている社員がありません。</td>
            </tr>
        <%
        Else
            Do While Not rs.EOF
        %>
            <tr>
                <td><%= rs("employee_id") %></td>
                <td><%= HtmlEncode(rs("employee_code")) %></td>
                <td><%= HtmlEncode(rs("employee_name")) %></td>
                <td><%= HtmlEncode(SafeValue(rs("email"), "-")) %></td>
                <td><%= IIf(rs("is_active") = 1, "有効", "無効") %></td>
                <td><%= FormatDateJP(rs("created_at")) %></td>
                <td class="action-links">
                    <a href="master_employee.asp?mode=edit&id=<%= rs("employee_id") %>" class="btn btn-sm btn-warning">編集</a>
                    <form method="post" action="master_employee.asp" style="display: inline;">
                        <input type="hidden" name="action" value="delete">
                        <input type="hidden" name="id" value="<%= rs("employee_id") %>">
                        <button type="submit" class="btn btn-sm btn-danger" onclick="return confirmDelete('この社員を削除しますか？');">削除</button>
                    </form>
                </td>
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
<% End If %>

<%
CloseDBConnection conn
%>
        </div>
    </main>
    <script src="js/common.js"></script>
</body>
</html>
