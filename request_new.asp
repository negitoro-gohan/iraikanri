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
<!--#include file="include/mail.asp"-->
<%
' ============================================
' 依頼新規登録画面
' ============================================

Dim conn, rs, sql
Dim errorMsg
errorMsg = ""

Set conn = GetDBConnection()

' POST処理（登録実行）
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim requesterId, assigneeId, requestTitle, requestContent, deadlineDate
    Dim requesterCode, assigneeCode, clientId, projectId, clientCode, projectCode

    requesterCode = Trim(Request.Form("requester_code") & "")
    assigneeCode  = Trim(Request.Form("assignee_code")  & "")
    requestTitle  = Trim(Request.Form("request_title")  & "")
    requestContent = Trim(Request.Form("request_content") & "")
    deadlineDate  = Trim(Request.Form("deadline_date")  & "")
    clientCode    = Trim(Request.Form("client_code")    & "")
    projectCode   = Trim(Request.Form("project_code")   & "")

    ' 社員コードからIDを逆引き
    requesterId = 0
    assigneeId  = 0
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

    ' 取引先コードからIDを逆引き
    clientId = 0
    If clientCode <> "" Then
        Dim rsClientLookup
        sql = "SELECT client_id FROM IRAI.M_Client WHERE client_code = N'" & EscapeSQL(clientCode) & "' AND is_active = 1"
        Set rsClientLookup = conn.Execute(sql)
        If Not rsClientLookup.EOF Then clientId = CLng(rsClientLookup("client_id"))
        CloseRecordset rsClientLookup
    End If

    ' 案件コードからIDを逆引き
    projectId = 0
    If projectCode <> "" Then
        Dim rsProjectLookup
        sql = "SELECT project_id FROM IRAI.M_Project WHERE project_code = N'" & EscapeSQL(projectCode) & "' AND is_active = 1"
        Set rsProjectLookup = conn.Execute(sql)
        If Not rsProjectLookup.EOF Then projectId = CLng(rsProjectLookup("project_id"))
        CloseRecordset rsProjectLookup
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
        clientIdSQL  = IIf(clientId  > 0, CStr(clientId),  "NULL")
        projectIdSQL = IIf(projectId > 0, CStr(projectId), "NULL")

        ' 登録処理
        sql = "INSERT INTO IRAI.T_Request (requester_id, assignee_id, request_title, request_content, deadline_date, status_id, client_id, project_id, created_by, updated_by) " & _
              "VALUES (" & requesterId & ", " & assigneeId & ", N'" & EscapeSQL(requestTitle) & "', N'" & EscapeSQL(requestContent) & "', '" & deadlineDate & "', " & STATUS_NOT_STARTED & ", " & clientIdSQL & ", " & projectIdSQL & ", N'" & EscapeSQL(domainUser) & "', N'" & EscapeSQL(domainUser) & "'); SELECT SCOPE_IDENTITY() AS NewID"

        On Error Resume Next
        Set rs = conn.Execute(sql)
        If Err.Number <> 0 Then
            errorMsg = "登録に失敗しました：" & Err.Description
            Err.Clear
        Else
            ' 新規登録されたIDを取得
            Dim newRequestId
            Set rs = rs.NextRecordset
            If Not rs Is Nothing Then
                If Not rs.EOF Then newRequestId = rs("NewID")
            End If
            CloseRecordset rs

            ' 新規依頼通知メールを送信
            If newRequestId > 0 Then SendNewRequestMail conn, newRequestId

            ' 成功メッセージをセットしてリダイレクト
            SetMessage "success", "依頼を登録しました。"
            Response.Redirect "request_list.asp"
            Response.End
        End If
        On Error GoTo 0
    End If
End If

' POST-backエラー時: 入力されたコードから名称を取得してスパンに表示する
Dim clientDisplayName, projectDisplayName
clientDisplayName  = ""
projectDisplayName = ""

If clientCode <> "" Then
    Dim rsClientDisp
    sql = "SELECT client_name FROM IRAI.M_Client WHERE client_code = N'" & EscapeSQL(clientCode) & "' AND is_active = 1"
    Set rsClientDisp = conn.Execute(sql)
    If Not rsClientDisp.EOF Then clientDisplayName = rsClientDisp("client_name") & ""
    CloseRecordset rsClientDisp
End If

If projectCode <> "" Then
    Dim rsProjectDisp
    sql = "SELECT project_name FROM IRAI.M_Project WHERE project_code = N'" & EscapeSQL(projectCode) & "' AND is_active = 1"
    Set rsProjectDisp = conn.Execute(sql)
    If Not rsProjectDisp.EOF Then projectDisplayName = rsProjectDisp("project_name") & ""
    CloseRecordset rsProjectDisp
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

<h2 class="page-title">依頼新規登録</h2>

<% If errorMsg <> "" Then %>
<div class="message message-error"><%= HtmlEncode(errorMsg) %></div>
<% End If %>

<div class="card">
    <form method="post" action="request_new.asp" onsubmit="return validateRequired(this);">

        <!-- 依頼元 -->
        <div class="form-group">
            <label>依頼元（社員コード）<span class="required">*</span></label>
            <div class="employee-code-group" style="display:flex;flex-direction:row;flex-wrap:nowrap;align-items:center;gap:10px;">
                <input type="text" name="requester_code" id="requester_code" class="form-control employee-code-input" style="width:200px;min-width:200px;max-width:200px;flex:0 0 200px;"
                       maxlength="20" required placeholder="社員コードを入力"
                       value="<%= HtmlEncode(Request.Form("requester_code") & "") %>"
                       onblur="lookupEmployee('requester')">
                <button type="button" class="btn btn-secondary btn-sm btn-emp-search" onclick="openEmployeeSearch('requester')">検索</button>
                <span id="requester_name_display" class="employee-name-display"></span>
            </div>
        </div>

        <!-- 依頼先 -->
        <div class="form-group">
            <label>依頼先（社員コード）<span class="required">*</span></label>
            <div class="employee-code-group" style="display:flex;flex-direction:row;flex-wrap:nowrap;align-items:center;gap:10px;">
                <input type="text" name="assignee_code" id="assignee_code" class="form-control employee-code-input" style="width:200px;min-width:200px;max-width:200px;flex:0 0 200px;"
                       maxlength="20" required placeholder="社員コードを入力"
                       value="<%= HtmlEncode(Request.Form("assignee_code") & "") %>"
                       onblur="lookupEmployee('assignee')">
                <button type="button" class="btn btn-secondary btn-sm btn-emp-search" onclick="openEmployeeSearch('assignee')">検索</button>
                <span id="assignee_name_display" class="employee-name-display"></span>
            </div>
        </div>

        <!-- 取引先（社員コードフィールドと同じデザイン） -->
        <div class="form-group">
            <label>取引先</label>
            <div class="employee-code-group" style="display:flex;flex-direction:row;flex-wrap:nowrap;align-items:center;gap:10px;">
                <input type="text" name="client_code" id="client_code" class="form-control employee-code-input" style="width:200px;min-width:200px;max-width:200px;flex:0 0 200px;"
                       maxlength="50" placeholder="取引先コードを入力"
                       value="<%= HtmlEncode(Request.Form("client_code") & "") %>"
                       onblur="lookupClient()">
                <button type="button" class="btn btn-secondary btn-sm btn-emp-search" onclick="openClientSearch()">検索</button>
                <span id="client_name_display" class="employee-name-display<% If clientDisplayName <> "" Then %> employee-name-found<% End If %>"><%= HtmlEncode(clientDisplayName) %></span>
            </div>
        </div>

        <!-- 案件（社員コードフィールドと同じデザイン） -->
        <div class="form-group">
            <label>案件</label>
            <div class="employee-code-group" style="display:flex;flex-direction:row;flex-wrap:nowrap;align-items:center;gap:10px;">
                <input type="text" name="project_code" id="project_code" class="form-control employee-code-input" style="width:200px;min-width:200px;max-width:200px;flex:0 0 200px;"
                       maxlength="50" placeholder="案件コードを入力"
                       value="<%= HtmlEncode(Request.Form("project_code") & "") %>"
                       onblur="lookupProject()">
                <button type="button" class="btn btn-secondary btn-sm btn-emp-search" onclick="openProjectSearch()">検索</button>
                <span id="project_name_display" class="employee-name-display<% If projectDisplayName <> "" Then %> employee-name-found<% End If %>"><%= HtmlEncode(projectDisplayName) %></span>
            </div>
        </div>

        <!-- 依頼件名 -->
        <div class="form-group">
            <label>依頼件名<span class="required">*</span></label>
            <input type="text" name="request_title" class="form-control" maxlength="200" required
                   value="<%= HtmlEncode(Request.Form("request_title") & "") %>">
        </div>

        <!-- 期限日 -->
        <div class="form-group">
            <label>期限日<span class="required">*</span></label>
            <input type="date" name="deadline_date" class="form-control" required
                   value="<%= HtmlEncode(Request.Form("deadline_date") & "") %>">
        </div>

        <!-- 依頼内容 -->
        <div class="form-group">
            <label>依頼内容</label>
            <textarea name="request_content" class="form-control" rows="5"><%= HtmlEncode(Request.Form("request_content") & "") %></textarea>
        </div>

        <div class="btn-group">
            <button type="submit" class="btn btn-primary">登録する</button>
            <a href="request_list.asp" class="btn btn-secondary">キャンセル</a>
        </div>
    </form>
</div>

<%
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
    var currentClientId = null; // 案件検索モーダルで取引先を絞り込むために保持

    function lookupClient() {
        var codeInput   = document.getElementById('client_code');
        var nameDisplay = document.getElementById('client_name_display');
        var code = codeInput.value.trim();

        if (code === '') {
            nameDisplay.textContent = '';
            nameDisplay.className = 'employee-name-display';
            currentClientId = null;
            // 取引先が消えたら案件もクリア
            document.getElementById('project_code').value = '';
            document.getElementById('project_name_display').textContent = '';
            document.getElementById('project_name_display').className = 'employee-name-display';
            allProjects = null;
            currentProjectClientId = null;
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
                        currentClientId = result.id;
                    } else {
                        nameDisplay.textContent = '該当なし';
                        nameDisplay.className = 'employee-name-display employee-name-notfound';
                        currentClientId = null;
                    }
                } catch (e) {
                    nameDisplay.textContent = '';
                    currentClientId = null;
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
            html += '<tr style="cursor:pointer;" onclick="selectClient(' + c.id + ',\'' + escJS(c.code) + '\',\'' + escJS(c.name) + '\')">' +
                    '<td>' + escHtml(c.code) + '</td><td>' + escHtml(c.name) + '</td></tr>';
        }
        html += '</tbody></table>';
        result.innerHTML = html;
    }

    function selectClient(id, code, name) {
        document.getElementById('client_code').value = code;
        var nameDisplay = document.getElementById('client_name_display');
        nameDisplay.textContent = name;
        nameDisplay.className = 'employee-name-display employee-name-found';
        currentClientId = id;
        closeClientSearch();
        // 取引先が変わったので案件をクリア
        document.getElementById('project_code').value = '';
        document.getElementById('project_name_display').textContent = '';
        document.getElementById('project_name_display').className = 'employee-name-display';
        allProjects = null;
        currentProjectClientId = null;
    }

    // ============================================================
    // 案件検索モーダル（選択中の取引先で絞り込み）
    // ============================================================
    var allProjects = null;
    var currentProjectClientId = null;

    function openProjectSearch() {
        if (!currentClientId) {
            alert('先に取引先コードを入力するか、検索ボタンで取引先を選択してください。');
            return;
        }
        document.getElementById('projectSearchKeyword').value = '';
        document.getElementById('projectSearchOverlay').style.display = 'flex';
        setTimeout(function() { document.getElementById('projectSearchKeyword').focus(); }, 100);

        // 取引先が変わった場合は再取得
        if (allProjects === null || currentProjectClientId !== String(currentClientId)) {
            currentProjectClientId = String(currentClientId);
            allProjects = null;
            var xhr = new XMLHttpRequest();
            xhr.open('GET', 'api_project_lookup.asp?action=projects&client_id=' + currentClientId, true);
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
            html += '<tr style="cursor:pointer;" onclick="selectProject(' + p.id + ',\'' + escJS(p.code) + '\',\'' + escJS(p.name) + '\')">' +
                    '<td>' + escHtml(p.code) + '</td><td>' + escHtml(p.name) + '</td></tr>';
        }
        html += '</tbody></table>';
        result.innerHTML = html;
    }

    function selectProject(id, code, name) {
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

    // ESCキーでモーダルを閉じる
    document.addEventListener('keydown', function(e) {
        if (e.key === 'Escape') { closeClientSearch(); closeProjectSearch(); }
    });

    // POST-backエラー時: コードが残っていれば名称を再表示して currentClientId をセット
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
