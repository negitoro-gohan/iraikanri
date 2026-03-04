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
' 対応状況入力画面
' ============================================

Dim conn, rs, sql
Dim requestId
Dim errorMsg
errorMsg = ""

requestId = SafeInt(Request.QueryString("id"), 0)
If requestId = 0 Then requestId = SafeInt(Request.Form("request_id"), 0)

If requestId = 0 Then
    SetMessage "error", "依頼IDが指定されていません。"
    Response.Redirect "request_list.asp"
    Response.End
End If

Set conn = GetDBConnection()

' POST処理（更新実行）
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim statusId, responseContent, endDate

    statusId = SafeInt(Request.Form("status_id"), 0)
    responseContent = Trim(Request.Form("response_content") & "")
    endDate = Trim(Request.Form("end_date") & "")

    ' 入力チェック
    If statusId = 0 Then
        errorMsg = "ステータスを選択してください。"
    ElseIf statusId = STATUS_COMPLETED And endDate = "" Then
        errorMsg = "対応完了の場合は終了日を入力してください。"
    ElseIf endDate <> "" And Not IsDate(endDate) Then
        errorMsg = "終了日の形式が正しくありません。"
    Else
        ' ドメインユーザー名を取得
        Dim domainUser
        domainUser = Request.ServerVariables("LOGON_USER")
        If domainUser = "" Then domainUser = Request.ServerVariables("AUTH_USER")

        ' 更新処理
        sql = "UPDATE IRAI.T_Request SET " & _
              "status_id = " & statusId & ", " & _
              "response_content = N'" & EscapeSQL(responseContent) & "', "

        If endDate <> "" Then
            sql = sql & "end_date = '" & endDate & "', "
        Else
            sql = sql & "end_date = NULL, "
        End If

        sql = sql & "updated_by = N'" & EscapeSQL(domainUser) & "', " & _
              "updated_at = GETDATE() " & _
              "WHERE request_id = " & requestId

        On Error Resume Next
        conn.Execute sql
        If Err.Number <> 0 Then
            errorMsg = "更新に失敗しました：" & Err.Description
            Err.Clear
        Else
            SetMessage "success", "対応状況を更新しました。"
            Response.Redirect "request_detail.asp?id=" & requestId
            Response.End
        End If
        On Error GoTo 0
    End If
End If

' 依頼データ取得
sql = "SELECT r.*, req.employee_name AS requester_name, ass.employee_name AS assignee_name " & _
      "FROM IRAI.T_Request r " & _
      "INNER JOIN IRAI.M_Employee req ON r.requester_id = req.employee_id " & _
      "INNER JOIN IRAI.M_Employee ass ON r.assignee_id = ass.employee_id " & _
      "WHERE r.request_id = " & requestId & " AND r.is_deleted = 0"
Set rs = conn.Execute(sql)

If rs.EOF Then
    CloseRecordset rs
    CloseDBConnection conn
    SetMessage "error", "指定された依頼が見つかりません。"
    Response.Redirect "request_list.asp"
    Response.End
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

<h2 class="page-title">対応状況入力</h2>

<% If errorMsg <> "" Then %>
<div class="message message-error"><%= HtmlEncode(errorMsg) %></div>
<% End If %>

<!-- 依頼情報（参照用） -->
<div class="card">
    <h3 class="card-title">依頼情報</h3>
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
            <td><%= HtmlEncode(rs("requester_name")) %></td>
        </tr>
        <tr>
            <th>依頼先</th>
            <td><%= HtmlEncode(rs("assignee_name")) %></td>
        </tr>
        <tr>
            <th>期限日</th>
            <td><%= FormatDateJP(rs("deadline_date")) %></td>
        </tr>
        <tr>
            <th>依頼内容</th>
            <td><%= NL2BR(HtmlEncode(SafeValue(rs("request_content"), "-"))) %></td>
        </tr>
    </table>
</div>

<!-- 対応入力フォーム -->
<div class="card">
    <h3 class="card-title">対応入力</h3>
    <form method="post" action="request_response.asp">
        <input type="hidden" name="request_id" value="<%= requestId %>">

        <div class="form-group">
            <label>ステータス<span class="required">*</span></label>
            <select name="status_id" class="form-control" required>
                <%
                Dim selectedStatusId
                If Request.Form("status_id") <> "" Then
                    selectedStatusId = SafeInt(Request.Form("status_id"), STATUS_NOT_STARTED)
                Else
                    selectedStatusId = rs("status_id")
                End If
                %>
                <option value="<%= STATUS_NOT_STARTED %>" <% If selectedStatusId = STATUS_NOT_STARTED Then Response.Write " selected" End If %>>未着手</option>
                <option value="<%= STATUS_IN_PROGRESS %>" <% If selectedStatusId = STATUS_IN_PROGRESS Then Response.Write " selected" End If %>>着手済</option>
                <option value="<%= STATUS_COMPLETED %>" <% If selectedStatusId = STATUS_COMPLETED Then Response.Write " selected" End If %>>対応完了</option>
                <option value="<%= STATUS_NOT_APPLICABLE %>" <% If selectedStatusId = STATUS_NOT_APPLICABLE Then Response.Write " selected" End If %>>対象外</option>
            </select>
        </div>

        <div class="form-group">
            <label>終了日</label>
            <%
            Dim respEndDate
            If Request.Form("end_date") <> "" Then
                respEndDate = Request.Form("end_date")
            Else
                respEndDate = FormatDateISO(rs("end_date"))
            End If
            %>
            <input type="date" name="end_date" class="form-control" value="<%= respEndDate %>">
            <small style="color: #666;">※対応完了の場合は必須</small>
        </div>

        <div class="form-group">
            <label>対応内容</label>
            <%
            Dim respContent
            If Request.Form("response_content") <> "" Then
                respContent = Request.Form("response_content")
            Else
                respContent = rs("response_content") & ""
            End If
            %>
            <textarea name="response_content" class="form-control" rows="5"><%= HtmlEncode(respContent) %></textarea>
        </div>

        <div class="btn-group">
            <button type="submit" class="btn btn-success">更新する</button>
            <a href="request_detail.asp?id=<%= requestId %>" class="btn btn-secondary">キャンセル</a>
        </div>
    </form>
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
