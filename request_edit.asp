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
    Dim requesterFound, assigneeFound, requestTitle, requestContent, deadlineDate, importance, folderAddress
    Dim requesterCode, assigneeCode, clientCode, projectCode

    requesterCode  = Trim(Request.Form("requester_code")  & "")
    assigneeCode   = Trim(Request.Form("assignee_code")   & "")
    requestTitle   = Trim(Request.Form("request_title")   & "")
    requestContent = Trim(Request.Form("request_content") & "")
    deadlineDate   = Trim(Request.Form("deadline_date")   & "")
    clientCode     = Trim(Request.Form("client_code")     & "")
    projectCode    = Trim(Request.Form("project_code")    & "")
    importance     = Trim(Request.Form("importance")      & "")
    folderAddress  = Trim(Request.Form("folder_address")  & "")
    ' 不正値はデフォルトのB（中）に補正
    If importance <> "A" And importance <> "B" And importance <> "C" Then importance = "B"

    ' 社員コードから名前を取得（Employee DBへ接続）
    requesterFound = False
    assigneeFound  = False
    Dim requesterName, assigneeName
    requesterName = ""
    assigneeName  = ""
    Dim connEmp
    Set connEmp = GetEmployeeDBConnection()
    If requesterCode <> "" Then
        Dim rsReqLookup
        sql = "SELECT employee_name FROM IRAI.M_Employee WHERE employee_code = N'" & EscapeSQL(requesterCode) & "' AND is_active = 1"
        Set rsReqLookup = connEmp.Execute(sql)
        If Not rsReqLookup.EOF Then
            requesterFound = True
            requesterName  = rsReqLookup("employee_name") & ""
        End If
        CloseRecordset rsReqLookup
    End If
    If assigneeCode <> "" Then
        Dim rsAsgLookup
        sql = "SELECT employee_name FROM IRAI.M_Employee WHERE employee_code = N'" & EscapeSQL(assigneeCode) & "' AND is_active = 1"
        Set rsAsgLookup = connEmp.Execute(sql)
        If Not rsAsgLookup.EOF Then
            assigneeFound = True
            assigneeName  = rsAsgLookup("employee_name") & ""
        End If
        CloseRecordset rsAsgLookup
    End If
    CloseDBConnection connEmp

    ' 取引先コードから名称を取得（Client DBへ接続）
    Dim clientName, projectName
    clientName  = ""
    projectName = ""
    Dim connClient
    Set connClient = GetClientDBConnection()
    If clientCode <> "" Then
        Dim rsClientLookup
        sql = "SELECT client_name FROM IRAI.M_Client WHERE client_code = N'" & EscapeSQL(clientCode) & "' AND is_active = 1"
        Set rsClientLookup = connClient.Execute(sql)
        If Not rsClientLookup.EOF Then clientName = rsClientLookup("client_name") & ""
        CloseRecordset rsClientLookup
    End If
    If projectCode <> "" Then
        Dim rsProjectLookup
        sql = "SELECT project_name FROM IRAI.M_Project WHERE project_code = N'" & EscapeSQL(projectCode) & "' AND is_active = 1"
        Set rsProjectLookup = connClient.Execute(sql)
        If Not rsProjectLookup.EOF Then projectName = rsProjectLookup("project_name") & ""
        CloseRecordset rsProjectLookup
    End If
    CloseDBConnection connClient

    ' 入力チェック
    If requesterCode = "" Then
        errorMsg = "依頼元の社員コードを入力してください。"
    ElseIf Not requesterFound Then
        errorMsg = "依頼元の社員コードが見つかりません。"
    ElseIf assigneeCode = "" Then
        errorMsg = "依頼先の社員コードを入力してください。"
    ElseIf Not assigneeFound Then
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
        domainUser = GetCurrentUser()

        ' 更新処理（社員名・取引先・案件・重要度・フォルダアドレスも一緒に更新）
        sql = "UPDATE IRAI.T_Request SET " & _
              "requester_code = N'" & EscapeSQL(requesterCode) & "', " & _
              "assignee_code = N'" & EscapeSQL(assigneeCode) & "', " & _
              "requester_name = N'" & EscapeSQL(requesterName) & "', " & _
              "assignee_name = N'" & EscapeSQL(assigneeName) & "', " & _
              "request_title = N'" & EscapeSQL(requestTitle) & "', " & _
              "request_content = N'" & EscapeSQL(requestContent) & "', " & _
              "deadline_date = '" & deadlineDate & "', " & _
              "importance = '" & EscapeSQL(importance) & "', " & _
              "folder_address = N'" & EscapeSQL(folderAddress) & "', " & _
              "client_code = N'" & EscapeSQL(clientCode) & "', " & _
              "client_name = N'" & EscapeSQL(clientName) & "', " & _
              "project_code = N'" & EscapeSQL(projectCode) & "', " & _
              "project_name = N'" & EscapeSQL(projectName) & "', " & _
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

' 依頼元・依頼先の社員コードをT_Requestから直接取得
Dim requesterEmployeeCode, assigneeEmployeeCode
requesterEmployeeCode = rs("requester_code") & ""
assigneeEmployeeCode  = rs("assignee_code")  & ""

' 取引先・案件の表示用名称を決定
' POST-backエラー時はフォーム値＋DB参照、初期表示時はT_Requestの保存値を使用
Dim editClientCode, editClientName, editProjectCode, editProjectName
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    ' POST-back: フォームのコードから名称を再取得
    editClientCode  = Trim(Request.Form("client_code")  & "")
    editProjectCode = Trim(Request.Form("project_code") & "")
    editClientName  = ""
    editProjectName = ""
    If editClientCode <> "" Or editProjectCode <> "" Then
        Dim connClientDisp
        Set connClientDisp = GetClientDBConnection()
        If editClientCode <> "" Then
            Dim rsClientDisp
            Set rsClientDisp = connClientDisp.Execute("SELECT client_name FROM IRAI.M_Client WHERE client_code = N'" & EscapeSQL(editClientCode) & "' AND is_active = 1")
            If Not rsClientDisp.EOF Then editClientName = rsClientDisp("client_name") & ""
            CloseRecordset rsClientDisp
        End If
        If editProjectCode <> "" Then
            Dim rsProjectDisp
            Set rsProjectDisp = connClientDisp.Execute("SELECT project_name FROM IRAI.M_Project WHERE project_code = N'" & EscapeSQL(editProjectCode) & "' AND is_active = 1")
            If Not rsProjectDisp.EOF Then editProjectName = rsProjectDisp("project_name") & ""
            CloseRecordset rsProjectDisp
        End If
        CloseDBConnection connClientDisp
    End If
Else
    ' 初期表示: T_Requestに保存済みの値をそのまま使用
    editClientCode  = rs("client_code")  & ""
    editClientName  = rs("client_name")  & ""
    editProjectCode = rs("project_code") & ""
    editProjectName = rs("project_name") & ""
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
            <label>最終更新者</label>
            <input type="text" class="form-control" value="<%= HtmlEncode(SafeValue(rs("updated_by"), "-")) %>" disabled>
        </div>

        <!-- 依頼元 -->
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

        <!-- 依頼先 -->
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

        <!-- 取引先 -->
        <div class="form-group">
            <label>取引先</label>
            <div class="employee-code-group" style="display:flex;flex-direction:row;flex-wrap:nowrap;align-items:center;gap:10px;">
                <input type="text" name="client_code" id="client_code" class="form-control employee-code-input" style="width:200px;min-width:200px;max-width:200px;flex:0 0 200px;"
                       maxlength="50" placeholder="取引先コードを入力"
                       value="<%= HtmlEncode(editClientCode) %>"
                       onblur="lookupClient()">
                <button type="button" class="btn btn-secondary btn-sm btn-emp-search" onclick="openClientSearch()">検索</button>
                <span id="client_name_display" class="employee-name-display<% If editClientName <> "" Then %> employee-name-found<% End If %>"><%= HtmlEncode(editClientName) %></span>
            </div>
        </div>

        <!-- 案件 -->
        <div class="form-group">
            <label>案件</label>
            <div class="employee-code-group" style="display:flex;flex-direction:row;flex-wrap:nowrap;align-items:center;gap:10px;">
                <input type="text" name="project_code" id="project_code" class="form-control employee-code-input" style="width:200px;min-width:200px;max-width:200px;flex:0 0 200px;"
                       maxlength="50" placeholder="案件コードを入力"
                       value="<%= HtmlEncode(editProjectCode) %>"
                       onblur="lookupProject()">
                <button type="button" class="btn btn-secondary btn-sm btn-emp-search" onclick="openProjectSearch()">検索</button>
                <span id="project_name_display" class="employee-name-display<% If editProjectName <> "" Then %> employee-name-found<% End If %>"><%= HtmlEncode(editProjectName) %></span>
            </div>
        </div>

        <!-- 依頼件名 -->
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

        <!-- 期限日 -->
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

        <!-- 重要度 -->
        <div class="form-group">
            <label>重要度<span class="required">*</span></label>
            <%
            Dim selImp
            If Request.Form("importance") <> "" Then
                selImp = Request.Form("importance")
            Else
                selImp = rs("importance") & ""
            End If
            If selImp <> "A" And selImp <> "B" And selImp <> "C" Then selImp = "B"
            %>
            <div style="display:flex;gap:24px;align-items:center;padding:6px 0;">
                <label style="display:flex;align-items:center;gap:6px;font-weight:normal;cursor:pointer;">
                    <input type="radio" name="importance" value="A" <% If selImp = "A" Then Response.Write "checked" End If %>>
                    <span style="color:#c0392b;font-weight:bold;">A（高）</span>
                </label>
                <label style="display:flex;align-items:center;gap:6px;font-weight:normal;cursor:pointer;">
                    <input type="radio" name="importance" value="B" <% If selImp = "B" Then Response.Write "checked" End If %>>
                    <span style="color:#e67e22;font-weight:bold;">B（中）</span>
                </label>
                <label style="display:flex;align-items:center;gap:6px;font-weight:normal;cursor:pointer;">
                    <input type="radio" name="importance" value="C" <% If selImp = "C" Then Response.Write "checked" End If %>>
                    <span style="color:#7f8c8d;font-weight:bold;">C（低）</span>
                </label>
            </div>
        </div>

        <!-- フォルダアドレス -->
        <div class="form-group">
            <label>フォルダアドレス</label>
            <%
            Dim editFolderAddress
            If Request.Form("folder_address") <> "" Then
                editFolderAddress = Request.Form("folder_address")
            Else
                editFolderAddress = rs("folder_address") & ""
            End If
            %>
            <div style="display:flex;gap:8px;align-items:center;">
                <input type="text" name="folder_address" id="folder_address" class="form-control" maxlength="500"
                       placeholder="例: \\server\share\project"
                       value="<%= HtmlEncode(editFolderAddress) %>">
                <button type="button" class="btn btn-secondary btn-sm" onclick="copyFolderAddress()" style="white-space:nowrap;">コピー</button>
            </div>
        </div>

        <!-- 依頼内容 -->
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
CloseRecordset rs
CloseDBConnection conn
%>
        </div>
    </main>

<!-- 取引先検索モーダル -->
<div id="clientSearchOverlay" style="display:none;position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.5);z-index:1000;justify-content:center;align-items:flex-start;padding-top:60px;box-sizing:border-box;" onclick="if(event.target===this)closeClientSearch()">
    <div style="background:#fff;border-radius:8px;padding:24px;width:620px;max-height:75vh;display:flex;flex-direction:column;box-shadow:0 4px 16px rgba(0,0,0,0.3);">
        <h3 style="margin:0 0 14px;font-size:1.1rem;">取引先検索</h3>
        <input type="text" id="clientSearchKeyword" class="form-control" placeholder="コード・名称で検索" oninput="filterClients()" style="margin-bottom:12px;">
        <div id="clientSearchResult" style="overflow-y:auto;flex:1;min-height:120px;">
            <p style="color:#999;text-align:center;padding:30px 0;">読み込み中...</p>
        </div>
        <div style="margin-top:14px;text-align:right;">
            <button type="button" class="btn btn-secondary" onclick="closeClientSearch()">閉じる</button>
        </div>
    </div>
</div>

<!-- 案件検索モーダル -->
<div id="projectSearchOverlay" style="display:none;position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.5);z-index:1000;justify-content:center;align-items:flex-start;padding-top:60px;box-sizing:border-box;" onclick="if(event.target===this)closeProjectSearch()">
    <div style="background:#fff;border-radius:8px;padding:24px;width:620px;max-height:75vh;display:flex;flex-direction:column;box-shadow:0 4px 16px rgba(0,0,0,0.3);">
        <h3 style="margin:0 0 14px;font-size:1.1rem;">案件検索</h3>
        <input type="text" id="projectSearchKeyword" class="form-control" placeholder="コード・案件名で検索" oninput="filterProjects()" style="margin-bottom:12px;">
        <div id="projectSearchResult" style="overflow-y:auto;flex:1;min-height:120px;">
            <p style="color:#999;text-align:center;padding:30px 0;">読み込み中...</p>
        </div>
        <div style="margin-top:14px;text-align:right;">
            <button type="button" class="btn btn-secondary" onclick="closeProjectSearch()">閉じる</button>
        </div>
    </div>
</div>

    <script src="js/common.js"></script>
    <script src="js/employee_search.js"></script>
    <script>
    // ============================================================
    // 取引先: コード入力 onblur → API参照して名称を表示
    // ============================================================
    var currentClientCode = null; // 案件検索モーダルで取引先を絞り込むために保持

    function lookupClient() {
        var codeInput   = document.getElementById('client_code');
        var nameDisplay = document.getElementById('client_name_display');
        var code = codeInput.value.trim();

        if (code === '') {
            nameDisplay.textContent = '';
            nameDisplay.className = 'employee-name-display';
            currentClientCode = null;
            // 取引先が消えたら案件もクリア
            document.getElementById('project_code').value = '';
            document.getElementById('project_name_display').textContent = '';
            document.getElementById('project_name_display').className = 'employee-name-display';
            allProjects = null;
            currentProjectClientCode = null;
            return;
        }

        var xhr = new XMLHttpRequest();
        xhr.open('GET', 'api_project_lookup.asp?action=lookup_client&code=' + encodeURIComponent(code), true);
        xhr.onreadystatechange = function() {
            if (xhr.readyState === 4 && xhr.status === 200) {
                try {
                    var result = JSON.parse(xhr.responseText);
                    if (result.found) {
                        nameDisplay.textContent = result.name;
                        nameDisplay.className = 'employee-name-display employee-name-found';
                        currentClientCode = code;
                    } else {
                        nameDisplay.textContent = '該当なし';
                        nameDisplay.className = 'employee-name-display employee-name-notfound';
                        currentClientCode = null;
                    }
                } catch (e) {
                    nameDisplay.textContent = '';
                    currentClientCode = null;
                }
            }
        };
        xhr.send();
    }

    // ============================================================
    // 案件: コード入力 onblur → API参照して名称を表示
    // ============================================================
    function lookupProject() {
        var codeInput   = document.getElementById('project_code');
        var nameDisplay = document.getElementById('project_name_display');
        var code = codeInput.value.trim();

        if (code === '') {
            nameDisplay.textContent = '';
            nameDisplay.className = 'employee-name-display';
            return;
        }

        var xhr = new XMLHttpRequest();
        xhr.open('GET', 'api_project_lookup.asp?action=lookup_project&code=' + encodeURIComponent(code), true);
        xhr.onreadystatechange = function() {
            if (xhr.readyState === 4 && xhr.status === 200) {
                try {
                    var result = JSON.parse(xhr.responseText);
                    if (result.found) {
                        nameDisplay.textContent = result.name;
                        nameDisplay.className = 'employee-name-display employee-name-found';
                    } else {
                        nameDisplay.textContent = '該当なし';
                        nameDisplay.className = 'employee-name-display employee-name-notfound';
                    }
                } catch (e) {
                    nameDisplay.textContent = '';
                }
            }
        };
        xhr.send();
    }

    // ============================================================
    // 取引先検索モーダル
    // ============================================================
    var allClients = null; // ページ内キャッシュ

    function openClientSearch() {
        document.getElementById('clientSearchKeyword').value = '';
        document.getElementById('clientSearchOverlay').style.display = 'flex';
        setTimeout(function() { document.getElementById('clientSearchKeyword').focus(); }, 100);

        if (allClients === null) {
            var xhr = new XMLHttpRequest();
            xhr.open('GET', 'api_project_lookup.asp?action=clients', true);
            xhr.onreadystatechange = function() {
                if (xhr.readyState === 4 && xhr.status === 200) {
                    try {
                        allClients = JSON.parse(xhr.responseText);
                        renderClientList('');
                    } catch (e) {
                        document.getElementById('clientSearchResult').innerHTML =
                            '<p style="color:red;text-align:center;padding:20px;">読み込みに失敗しました。</p>';
                    }
                }
            };
            xhr.send();
        } else {
            renderClientList('');
        }
    }

    function closeClientSearch() {
        document.getElementById('clientSearchOverlay').style.display = 'none';
    }

    function filterClients() {
        renderClientList(document.getElementById('clientSearchKeyword').value);
    }

    function renderClientList(keyword) {
        var result = document.getElementById('clientSearchResult');
        if (!allClients) {
            result.innerHTML = '<p style="color:#999;text-align:center;padding:20px;">読み込み中...</p>';
            return;
        }
        keyword = keyword.toLowerCase();
        var filtered = allClients.filter(function(c) {
            return keyword === '' ||
                   c.code.toLowerCase().indexOf(keyword) >= 0 ||
                   c.name.toLowerCase().indexOf(keyword) >= 0;
        });
        if (filtered.length === 0) {
            result.innerHTML = '<p style="color:#999;text-align:center;padding:20px;">該当する取引先がありません。</p>';
            return;
        }
        var html = '<table class="data-table" style="width:100%;"><thead><tr><th>コード</th><th>取引先名</th></tr></thead><tbody>';
        for (var i = 0; i < filtered.length; i++) {
            var c = filtered[i];
            html += '<tr style="cursor:pointer;" onclick="selectClient(\'' + escJS(c.code) + '\',\'' + escJS(c.name) + '\')">' +
                    '<td>' + escHtml(c.code) + '</td><td>' + escHtml(c.name) + '</td></tr>';
        }
        html += '</tbody></table>';
        result.innerHTML = html;
    }

    function selectClient(code, name) {
        document.getElementById('client_code').value = code;
        var nameDisplay = document.getElementById('client_name_display');
        nameDisplay.textContent = name;
        nameDisplay.className = 'employee-name-display employee-name-found';
        currentClientCode = code;
        closeClientSearch();
        // 取引先が変わったので案件をクリア
        document.getElementById('project_code').value = '';
        document.getElementById('project_name_display').textContent = '';
        document.getElementById('project_name_display').className = 'employee-name-display';
        allProjects = null;
        currentProjectClientCode = null;
    }

    // ============================================================
    // 案件検索モーダル（選択中の取引先で絞り込み）
    // ============================================================
    var allProjects = null;
    var currentProjectClientCode = null;

    function openProjectSearch() {
        if (!currentClientCode) {
            alert('先に取引先コードを入力するか、検索ボタンで取引先を選択してください。');
            return;
        }
        document.getElementById('projectSearchKeyword').value = '';
        document.getElementById('projectSearchOverlay').style.display = 'flex';
        setTimeout(function() { document.getElementById('projectSearchKeyword').focus(); }, 100);

        // 取引先が変わった場合は再取得
        if (allProjects === null || currentProjectClientCode !== currentClientCode) {
            currentProjectClientCode = currentClientCode;
            allProjects = null;
            var xhr = new XMLHttpRequest();
            xhr.open('GET', 'api_project_lookup.asp?action=projects&client_code=' + encodeURIComponent(currentClientCode), true);
            xhr.onreadystatechange = function() {
                if (xhr.readyState === 4 && xhr.status === 200) {
                    try {
                        allProjects = JSON.parse(xhr.responseText);
                        renderProjectList('');
                    } catch (e) {
                        document.getElementById('projectSearchResult').innerHTML =
                            '<p style="color:red;text-align:center;padding:20px;">読み込みに失敗しました。</p>';
                    }
                }
            };
            xhr.send();
        } else {
            renderProjectList('');
        }
    }

    function closeProjectSearch() {
        document.getElementById('projectSearchOverlay').style.display = 'none';
    }

    function filterProjects() {
        renderProjectList(document.getElementById('projectSearchKeyword').value);
    }

    function renderProjectList(keyword) {
        var result = document.getElementById('projectSearchResult');
        if (!allProjects) {
            result.innerHTML = '<p style="color:#999;text-align:center;padding:20px;">読み込み中...</p>';
            return;
        }
        keyword = keyword.toLowerCase();
        var filtered = allProjects.filter(function(p) {
            return keyword === '' ||
                   p.code.toLowerCase().indexOf(keyword) >= 0 ||
                   p.name.toLowerCase().indexOf(keyword) >= 0;
        });
        if (filtered.length === 0) {
            result.innerHTML = '<p style="color:#999;text-align:center;padding:20px;">該当する案件がありません。</p>';
            return;
        }
        var html = '<table class="data-table" style="width:100%;"><thead><tr><th>コード</th><th>案件名</th></tr></thead><tbody>';
        for (var i = 0; i < filtered.length; i++) {
            var p = filtered[i];
            html += '<tr style="cursor:pointer;" onclick="selectProject(\'' + escJS(p.code) + '\',\'' + escJS(p.name) + '\')">' +
                    '<td>' + escHtml(p.code) + '</td><td>' + escHtml(p.name) + '</td></tr>';
        }
        html += '</tbody></table>';
        result.innerHTML = html;
    }

    function selectProject(code, name) {
        document.getElementById('project_code').value = code;
        var nameDisplay = document.getElementById('project_name_display');
        nameDisplay.textContent = name;
        nameDisplay.className = 'employee-name-display employee-name-found';
        closeProjectSearch();
    }

    // ============================================================
    // ユーティリティ
    // ============================================================
    function escHtml(str) {
        return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
    }
    function escJS(str) {
        return String(str).replace(/\\/g,'\\\\').replace(/'/g,"\\'");
    }

    // ============================================================
    // フォルダアドレス コピー
    // ============================================================
    function copyFolderAddress() {
        var input = document.getElementById('folder_address');
        var text = input.value.trim();
        if (text === '') {
            alert('フォルダアドレスが入力されていません。');
            return;
        }
        if (navigator.clipboard && window.isSecureContext) {
            navigator.clipboard.writeText(text).then(function() {
                showCopyToast();
            });
        } else {
            // フォールバック（HTTP環境など）
            input.select();
            document.execCommand('copy');
            showCopyToast();
        }
    }

    function showCopyToast() {
        var toast = document.getElementById('copyToast');
        if (!toast) {
            toast = document.createElement('div');
            toast.id = 'copyToast';
            toast.textContent = 'コピーしました';
            toast.style.cssText = 'position:fixed;bottom:30px;left:50%;transform:translateX(-50%);background:#333;color:#fff;padding:8px 20px;border-radius:4px;font-size:0.9rem;z-index:9999;';
            document.body.appendChild(toast);
        }
        toast.style.display = 'block';
        setTimeout(function() { toast.style.display = 'none'; }, 2000);
    }

    // ESCキーでモーダルを閉じる
    document.addEventListener('keydown', function(e) {
        if (e.key === 'Escape') { closeClientSearch(); closeProjectSearch(); }
    });

    // ページ読み込み時: コードが入力済みなら名称を表示し currentClientId をセット
    document.addEventListener('DOMContentLoaded', function() {
        if (document.getElementById('client_code').value.trim() !== '') {
            lookupClient();
        }
        if (document.getElementById('project_code').value.trim() !== '') {
            lookupProject();
        }
    });
    </script>
</body>
</html>
