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
' 依頼編集画面
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
    Dim requesterId, assigneeId, requestTitle, requestContent, deadlineDate
    Dim requesterCode, assigneeCode, clientId, projectId

    requesterCode = Trim(Request.Form("requester_code") & "")
    assigneeCode = Trim(Request.Form("assignee_code") & "")
    requestTitle = Trim(Request.Form("request_title") & "")
    requestContent = Trim(Request.Form("request_content") & "")
    deadlineDate = Trim(Request.Form("deadline_date") & "")
    clientId = SafeInt(Request.Form("client_id"), 0)
    projectId = SafeInt(Request.Form("project_id"), 0)

    ' 社員コードからIDを逆引き
    requesterId = 0
    assigneeId = 0
    If requesterCode <> "" Then
        Dim rsReqLookup
        sql = "SELECT employee_id FROM IRAI.M_Employee WHERE employee_code = N'" & EscapeSQL(requesterCode) & "' AND is_active = 1"
        Set rsReqLookup = conn.Execute(sql)
        If Not rsReqLookup.EOF Then requesterId = CLng(rsReqLookup("employee_id"))
        CloseRecordset rsReqLookup
    End If
    If assigneeCode <> "" Then
        Dim rsAsgLookup
        sql = "SELECT employee_id FROM IRAI.M_Employee WHERE employee_code = N'" & EscapeSQL(assigneeCode) & "' AND is_active = 1"
        Set rsAsgLookup = conn.Execute(sql)
        If Not rsAsgLookup.EOF Then assigneeId = CLng(rsAsgLookup("employee_id"))
        CloseRecordset rsAsgLookup
    End If

    ' 入力チェック
    If requesterCode = "" Then
        errorMsg = "依頼元の社員コードを入力してください。"
    ElseIf requesterId = 0 Then
        errorMsg = "依頼元の社員コードが見つかりません。"
    ElseIf assigneeCode = "" Then
        errorMsg = "依頼先の社員コードを入力してください。"
    ElseIf assigneeId = 0 Then
        errorMsg = "依頼先の社員コードが見つかりません。"
    ElseIf requestTitle = "" Then
        errorMsg = "依頼件名を入力してください。"
    ElseIf deadlineDate = "" Then
        errorMsg = "期限日を入力してください。"
    ElseIf Not IsDate(deadlineDate) Then
        errorMsg = "期限日の形式が正しくありません。"
    Else
        ' ドメインユーザー名を取得
        Dim domainUser
        domainUser = Request.ServerVariables("LOGON_USER")
        If domainUser = "" Then domainUser = Request.ServerVariables("AUTH_USER")

        ' client_id, project_id が0の場合はNULLとして扱う
        Dim clientIdSQL, projectIdSQL
        If clientId > 0 Then
            clientIdSQL = CStr(clientId)
        Else
            clientIdSQL = "NULL"
        End If
        If projectId > 0 Then
            projectIdSQL = CStr(projectId)
        Else
            projectIdSQL = "NULL"
        End If

        ' 更新処理
        sql = "UPDATE IRAI.T_Request SET " & _
              "requester_id = " & requesterId & ", " & _
              "assignee_id = " & assigneeId & ", " & _
              "request_title = N'" & EscapeSQL(requestTitle) & "', " & _
              "request_content = N'" & EscapeSQL(requestContent) & "', " & _
              "deadline_date = '" & deadlineDate & "', " & _
              "client_id = " & clientIdSQL & ", " & _
              "project_id = " & projectIdSQL & ", " & _
              "updated_by = N'" & EscapeSQL(domainUser) & "', " & _
              "updated_at = GETDATE() " & _
              "WHERE request_id = " & requestId

        On Error Resume Next
        conn.Execute sql
        If Err.Number <> 0 Then
            errorMsg = "更新に失敗しました：" & Err.Description
            Err.Clear
        Else
            SetMessage "success", "依頼を更新しました。"
            Response.Redirect "request_detail.asp?id=" & requestId
            Response.End
        End If
        On Error GoTo 0
    End If
End If

' 依頼データ取得
sql = "SELECT * FROM IRAI.T_Request WHERE request_id = " & requestId & " AND is_deleted = 0"
Set rs = conn.Execute(sql)

If rs.EOF Then
    CloseRecordset rs
    CloseDBConnection conn
    SetMessage "error", "指定された依頼が見つかりません。"
    Response.Redirect "request_list.asp"
    Response.End
End If

' 現在の依頼元・依頼先の社員コードを取得
Dim requesterEmployeeCode, assigneeEmployeeCode
Dim rsEmpReq, rsEmpAsg
Set rsEmpReq = conn.Execute("SELECT employee_code FROM IRAI.M_Employee WHERE employee_id = " & rs("requester_id"))
If Not rsEmpReq.EOF Then requesterEmployeeCode = rsEmpReq("employee_code") & ""
CloseRecordset rsEmpReq
Set rsEmpAsg = conn.Execute("SELECT employee_code FROM IRAI.M_Employee WHERE employee_id = " & rs("assignee_id"))
If Not rsEmpAsg.EOF Then assigneeEmployeeCode = rsEmpAsg("employee_code") & ""
CloseRecordset rsEmpAsg

' 現在の取引先・案件ID
Dim currentClientId, currentProjectId
currentClientId = 0
currentProjectId = 0
If Not IsNull(rs("client_id")) Then currentClientId = rs("client_id")
If Not IsNull(rs("project_id")) Then currentProjectId = rs("project_id")

' 取引先一覧取得
Dim rsClients
sql = "SELECT client_id, client_code, client_name FROM IRAI.M_Client WHERE is_active = 1 ORDER BY client_code"
Set rsClients = conn.Execute(sql)
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

<h2 class="page-title">依頼編集</h2>

<% If errorMsg <> "" Then %>
<div class="message message-error"><%= HtmlEncode(errorMsg) %></div>
<% End If %>

<div class="card">
    <form method="post" action="request_edit.asp" onsubmit="return validateRequired(this);">
        <input type="hidden" name="request_id" value="<%= requestId %>">

        <div class="form-group">
            <label>依頼ID</label>
            <input type="text" class="form-control" value="<%= requestId %>" disabled>
        </div>

        <div class="form-group">
            <label>依頼元（社員コード）<span class="required">*</span></label>
            <div class="employee-code-group" style="display:flex;flex-direction:row;flex-wrap:nowrap;align-items:center;gap:10px;">
                <%
                Dim editRequesterCode
                If Request.Form("requester_code") <> "" Then
                    editRequesterCode = Request.Form("requester_code")
                Else
                    editRequesterCode = requesterEmployeeCode
                End If
                %>
                <input type="text" name="requester_code" id="requester_code" class="form-control employee-code-input" style="width:200px;min-width:200px;max-width:200px;flex:0 0 200px;"
                       maxlength="20" required placeholder="社員コードを入力"
                       value="<%= HtmlEncode(editRequesterCode) %>"
                       onblur="lookupEmployee('requester')">
                <button type="button" class="btn btn-secondary btn-sm btn-emp-search" onclick="openEmployeeSearch('requester')">検索</button>
                <span id="requester_name_display" class="employee-name-display"></span>
            </div>
        </div>

        <div class="form-group">
            <label>依頼先（社員コード）<span class="required">*</span></label>
            <div class="employee-code-group" style="display:flex;flex-direction:row;flex-wrap:nowrap;align-items:center;gap:10px;">
                <%
                Dim editAssigneeCode
                If Request.Form("assignee_code") <> "" Then
                    editAssigneeCode = Request.Form("assignee_code")
                Else
                    editAssigneeCode = assigneeEmployeeCode
                End If
                %>
                <input type="text" name="assignee_code" id="assignee_code" class="form-control employee-code-input" style="width:200px;min-width:200px;max-width:200px;flex:0 0 200px;"
                       maxlength="20" required placeholder="社員コードを入力"
                       value="<%= HtmlEncode(editAssigneeCode) %>"
                       onblur="lookupEmployee('assignee')">
                <button type="button" class="btn btn-secondary btn-sm btn-emp-search" onclick="openEmployeeSearch('assignee')">検索</button>
                <span id="assignee_name_display" class="employee-name-display"></span>
            </div>
        </div>

        <div class="form-group">
            <label>取引先</label>
            <%
            Dim editClientId
            If Request.Form("client_id") <> "" Then
                editClientId = SafeInt(Request.Form("client_id"), 0)
            Else
                editClientId = currentClientId
            End If
            %>
            <select name="client_id" id="client_id" class="form-control" style="width:400px;" onchange="loadProjects();">
                <option value="">-- 選択してください --</option>
                <% Do While Not rsClients.EOF %>
                <option value="<%= rsClients("client_id") %>" <% If editClientId = CLng(rsClients("client_id")) Then %>selected<% End If %>><%= HtmlEncode(rsClients("client_code")) %> : <%= HtmlEncode(rsClients("client_name")) %></option>
                <% rsClients.MoveNext : Loop %>
            </select>
        </div>

        <div class="form-group">
            <label>案件</label>
            <select name="project_id" id="project_id" class="form-control" style="width:400px;">
                <option value="">-- 取引先を選択してください --</option>
            </select>
        </div>

        <div class="form-group">
            <label>依頼件名<span class="required">*</span></label>
            <%
            Dim editTitle
            If Request.Form("request_title") <> "" Then
                editTitle = Request.Form("request_title")
            Else
                editTitle = rs("request_title")
            End If
            %>
            <input type="text" name="request_title" class="form-control" maxlength="200" required
                   value="<%= HtmlEncode(editTitle) %>">
        </div>

        <div class="form-group">
            <label>期限日<span class="required">*</span></label>
            <%
            Dim editDeadline
            If Request.Form("deadline_date") <> "" Then
                editDeadline = Request.Form("deadline_date")
            Else
                editDeadline = FormatDateISO(rs("deadline_date"))
            End If
            %>
            <input type="date" name="deadline_date" class="form-control" required
                   value="<%= editDeadline %>">
        </div>

        <div class="form-group">
            <label>依頼内容</label>
            <%
            Dim editContent
            If Request.Form("request_content") <> "" Then
                editContent = Request.Form("request_content")
            Else
                editContent = rs("request_content") & ""
            End If
            %>
            <textarea name="request_content" class="form-control" rows="5"><%= HtmlEncode(editContent) %></textarea>
        </div>

        <div class="btn-group">
            <button type="submit" class="btn btn-primary">更新する</button>
            <a href="request_detail.asp?id=<%= requestId %>" class="btn btn-secondary">キャンセル</a>
        </div>
    </form>
</div>

<%
' 現在の案件IDを保持
Dim editProjectId
If Request.Form("project_id") <> "" Then
    editProjectId = SafeInt(Request.Form("project_id"), 0)
Else
    editProjectId = currentProjectId
End If

CloseRecordset rsClients
CloseRecordset rs
CloseDBConnection conn
%>
        </div>
    </main>
    <script src="js/common.js"></script>
    <script src="js/employee_search.js"></script>
    <script>
    // 取引先選択時に案件一覧を読み込む
    function loadProjects(selectedProjectId) {
        var clientId = document.getElementById('client_id').value;
        var projectSelect = document.getElementById('project_id');

        // クリア
        projectSelect.innerHTML = '<option value="">-- 案件を選択してください --</option>';

        if (!clientId) {
            projectSelect.innerHTML = '<option value="">-- 取引先を選択してください --</option>';
            return;
        }

        var xhr = new XMLHttpRequest();
        xhr.open('GET', 'api_project_lookup.asp?action=projects&client_id=' + clientId, true);
        xhr.onreadystatechange = function() {
            if (xhr.readyState === 4 && xhr.status === 200) {
                try {
                    var projects = JSON.parse(xhr.responseText);
                    projectSelect.innerHTML = '<option value="">-- 案件を選択してください --</option>';
                    for (var i = 0; i < projects.length; i++) {
                        var opt = document.createElement('option');
                        opt.value = projects[i].id;
                        opt.textContent = projects[i].code + ' : ' + projects[i].name;
                        if (selectedProjectId && projects[i].id == selectedProjectId) {
                            opt.selected = true;
                        }
                        projectSelect.appendChild(opt);
                    }
                } catch (e) {
                    console.error('Failed to parse projects', e);
                }
            }
        };
        xhr.send();
    }

    // ページ読み込み時に取引先が選択されていれば案件を読み込む
    document.addEventListener('DOMContentLoaded', function() {
        var clientId = document.getElementById('client_id').value;
        if (clientId) {
            loadProjects(<%= editProjectId %>);
        }
    });
    </script>
</body>
</html>
