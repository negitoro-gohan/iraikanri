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
' 案件マスター管理画面
' ============================================

Dim conn, rs, sql
Dim errorMsg, mode, projectId
Dim filterClientCode, filterClientId, filterClientName
Dim formClientCode, formClientId, formClientName
errorMsg = ""
mode = Trim(Request.QueryString("mode") & "")
projectId = SafeInt(Request.QueryString("id"), 0)
filterClientCode = Trim(Request.QueryString("client_code") & "")
filterClientId = 0
filterClientName = ""
formClientCode = ""
formClientId = 0
formClientName = ""

Set conn = GetDBConnection()

' フィルタ用取引先コードからIDと名称を取得
If filterClientCode <> "" Then
    Dim rsFilter
    Set rsFilter = conn.Execute("SELECT client_id, client_name FROM M_Client WHERE client_code = N'" & EscapeSQL(filterClientCode) & "' AND is_active = 1")
    If Not rsFilter.EOF Then
        filterClientId = CLng(rsFilter("client_id"))
        filterClientName = rsFilter("client_name") & ""
    End If
    CloseRecordset rsFilter
End If

' POST処理
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim action, postProjectId, projectCode, projectName, isActive
    action = Trim(Request.Form("action") & "")
    postProjectId = SafeInt(Request.Form("project_id"), 0)
    formClientCode = Trim(Request.Form("client_code") & "")
    projectCode = Trim(Request.Form("project_code") & "")
    projectName = Trim(Request.Form("project_name") & "")
    isActive = SafeInt(Request.Form("is_active"), 1)

    ' 取引先コードからIDと名称を取得
    If formClientCode <> "" Then
        Dim rsFormClient
        Set rsFormClient = conn.Execute("SELECT client_id, client_name FROM M_Client WHERE client_code = N'" & EscapeSQL(formClientCode) & "' AND is_active = 1")
        If Not rsFormClient.EOF Then
            formClientId = CLng(rsFormClient("client_id"))
            formClientName = rsFormClient("client_name") & ""
        End If
        CloseRecordset rsFormClient
    End If

    If action = "add" Then
        ' 新規登録
        If formClientId = 0 Then
            errorMsg = "有効な取引先コードを入力してください。"
        ElseIf projectCode = "" Then
            errorMsg = "案件コードを入力してください。"
        ElseIf projectName = "" Then
            errorMsg = "案件名を入力してください。"
        Else
            ' 重複チェック
            sql = "SELECT COUNT(*) FROM M_Project WHERE project_code = N'" & EscapeSQL(projectCode) & "'"
            Set rs = conn.Execute(sql)
            If rs(0) > 0 Then
                errorMsg = "この案件コードは既に登録されています。"
            Else
                sql = "INSERT INTO M_Project (client_id, project_code, project_name, is_active) VALUES (" & formClientId & ", N'" & EscapeSQL(projectCode) & "', N'" & EscapeSQL(projectName) & "', " & isActive & ")"
                On Error Resume Next
                conn.Execute sql
                If Err.Number <> 0 Then
                    errorMsg = "登録に失敗しました：" & Err.Description
                    Err.Clear
                Else
                    SetMessage "success", "案件を登録しました。"
                    Response.Redirect "master_project.asp?client_code=" & Server.URLEncode(formClientCode)
                    Response.End
                End If
                On Error GoTo 0
            End If
            CloseRecordset rs
        End If

    ElseIf action = "edit" And postProjectId > 0 Then
        ' 更新
        If formClientId = 0 Then
            errorMsg = "有効な取引先コードを入力してください。"
        ElseIf projectCode = "" Then
            errorMsg = "案件コードを入力してください。"
        ElseIf projectName = "" Then
            errorMsg = "案件名を入力してください。"
        Else
            ' 重複チェック（自分以外）
            sql = "SELECT COUNT(*) FROM M_Project WHERE project_code = N'" & EscapeSQL(projectCode) & "' AND project_id <> " & postProjectId
            Set rs = conn.Execute(sql)
            If rs(0) > 0 Then
                errorMsg = "この案件コードは既に使用されています。"
            Else
                sql = "UPDATE M_Project SET client_id = " & formClientId & ", project_code = N'" & EscapeSQL(projectCode) & "', project_name = N'" & EscapeSQL(projectName) & "', is_active = " & isActive & ", updated_at = GETDATE() WHERE project_id = " & postProjectId
                On Error Resume Next
                conn.Execute sql
                If Err.Number <> 0 Then
                    errorMsg = "更新に失敗しました：" & Err.Description
                    Err.Clear
                Else
                    SetMessage "success", "案件を更新しました。"
                    Response.Redirect "master_project.asp?client_code=" & Server.URLEncode(formClientCode)
                    Response.End
                End If
                On Error GoTo 0
            End If
            CloseRecordset rs
        End If

    ElseIf action = "delete" And postProjectId > 0 Then
        ' 削除（依頼が紐づいている場合は削除不可）
        sql = "SELECT COUNT(*) FROM T_Request WHERE project_id = " & postProjectId
        Set rs = conn.Execute(sql)
        If rs(0) > 0 Then
            errorMsg = "この案件には依頼が登録されているため削除できません。"
        Else
            ' リダイレクト用に取引先コードを取得
            Dim rsDelProj, delClientCode
            delClientCode = ""
            Set rsDelProj = conn.Execute("SELECT c.client_code FROM M_Project p INNER JOIN M_Client c ON p.client_id = c.client_id WHERE p.project_id = " & postProjectId)
            If Not rsDelProj.EOF Then delClientCode = rsDelProj("client_code") & ""
            CloseRecordset rsDelProj

            sql = "DELETE FROM M_Project WHERE project_id = " & postProjectId
            On Error Resume Next
            conn.Execute sql
            If Err.Number <> 0 Then
                errorMsg = "削除に失敗しました：" & Err.Description
                Err.Clear
            Else
                SetMessage "success", "案件を削除しました。"
                Response.Redirect "master_project.asp?client_code=" & Server.URLEncode(delClientCode)
                Response.End
            End If
            On Error GoTo 0
        End If
        CloseRecordset rs
    End If
End If

' 編集モードの場合、データ取得（取引先情報も結合）
Dim editProject, editClientCode, editClientName
editClientCode = ""
editClientName = ""
If mode = "edit" And projectId > 0 Then
    sql = "SELECT p.*, c.client_code, c.client_name FROM M_Project p " & _
          "INNER JOIN M_Client c ON p.client_id = c.client_id " & _
          "WHERE p.project_id = " & projectId
    Set editProject = conn.Execute(sql)
    If editProject.EOF Then
        mode = ""
        projectId = 0
    Else
        editClientCode = editProject("client_code") & ""
        editClientName = editProject("client_name") & ""
        ' フィルタが未指定の場合、編集対象の取引先でフィルタを設定
        If filterClientCode = "" Then
            filterClientCode = editClientCode
            filterClientId = CLng(editProject("client_id"))
            filterClientName = editClientName
        End If
    End If
End If

' フォーム表示用の取引先コード・名称を決定
Dim formDisplayClientCode, formDisplayClientName
If mode = "edit" Then
    formDisplayClientCode = editClientCode
    formDisplayClientName = editClientName
ElseIf Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    ' POST-backエラー時：送信済みのコードを表示
    formDisplayClientCode = formClientCode
    formDisplayClientName = formClientName
Else
    ' 新規登録：フィルタの取引先をプリセット
    formDisplayClientCode = filterClientCode
    formDisplayClientName = filterClientName
End If

' 案件一覧取得
sql = "SELECT p.*, c.client_code, c.client_name FROM M_Project p " & _
      "INNER JOIN M_Client c ON p.client_id = c.client_id"
If filterClientId > 0 Then
    sql = sql & " WHERE p.client_id = " & filterClientId
End If
sql = sql & " ORDER BY c.client_code, p.project_code"
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

<h2 class="page-title">案件マスター管理</h2>

<% If errorMsg <> "" Then %>
<div class="message message-error"><%= HtmlEncode(errorMsg) %></div>
<% End If %>

<!-- 登録・編集フォーム -->
<div class="card">
    <h3 class="card-title"><% If mode = "edit" Then %>案件編集<% Else %>案件新規登録<% End If %></h3>
    <form method="post" action="master_project.asp">
        <input type="hidden" name="action" value="<% If mode = "edit" Then %>edit<% Else %>add<% End If %>">
        <% If mode = "edit" Then %>
        <input type="hidden" name="project_id" value="<%= projectId %>">
        <% End If %>

        <div class="form-group">
            <label>取引先<span class="required">*</span></label>
            <div class="employee-code-group" style="display:flex;flex-direction:row;flex-wrap:nowrap;align-items:center;gap:10px;">
                <input type="text" name="client_code" id="formClientCode" class="form-control employee-code-input"
                       style="width:200px;min-width:200px;max-width:200px;flex:0 0 200px;"
                       maxlength="50" placeholder="取引先コードを入力"
                       value="<%= HtmlEncode(formDisplayClientCode) %>"
                       onblur="lookupFormClient()">
                <button type="button" class="btn btn-secondary btn-sm btn-emp-search" onclick="openClientSearch('form')">検索</button>
                <span id="formClientNameDisplay" class="employee-name-display<% If formDisplayClientName <> "" Then %> employee-name-found<% End If %>"><%= HtmlEncode(formDisplayClientName) %></span>
            </div>
        </div>

        <div class="form-group">
            <label>案件コード<span class="required">*</span></label>
            <input type="text" name="project_code" class="form-control" maxlength="50" required style="width:200px;"
                   value="<% If mode = "edit" Then %><%= HtmlEncode(editProject("project_code") & "") %><% Else %><%= HtmlEncode(Request.Form("project_code") & "") %><% End If %>">
        </div>

        <div class="form-group">
            <label>案件名<span class="required">*</span></label>
            <input type="text" name="project_name" class="form-control" maxlength="200" required
                   value="<% If mode = "edit" Then %><%= HtmlEncode(editProject("project_name") & "") %><% Else %><%= HtmlEncode(Request.Form("project_name") & "") %><% End If %>">
        </div>

        <div class="form-group">
            <label>有効</label>
            <select name="is_active" class="form-control" style="width:100px;">
                <%
                Dim editProjActive
                If mode = "edit" Then
                    editProjActive = editProject("is_active")
                Else
                    editProjActive = 1
                End If
                %>
                <option value="1" <% If editProjActive = 1 Then %>selected<% End If %>>有効</option>
                <option value="0" <% If editProjActive = 0 Then %>selected<% End If %>>無効</option>
            </select>
        </div>

        <div class="btn-group">
            <button type="submit" class="btn btn-primary"><% If mode = "edit" Then %>更新<% Else %>登録<% End If %></button>
            <% If mode = "edit" Then %>
            <a href="master_project.asp?client_code=<%= Server.URLEncode(filterClientCode) %>" class="btn btn-secondary">キャンセル</a>
            <% End If %>
        </div>
    </form>
</div>

<!-- 絞り込み -->
<div class="card">
    <h3 class="card-title">絞り込み</h3>
    <form method="get" action="master_project.asp" id="filterForm">
        <div class="form-group">
            <label>取引先</label>
            <input type="hidden" name="client_code" id="filterClientCodeHidden" value="<%= HtmlEncode(filterClientCode) %>">
            <div class="employee-code-group" style="display:flex;flex-direction:row;flex-wrap:nowrap;align-items:center;gap:10px;">
                <input type="text" id="filterClientCodeDisplay" class="form-control"
                       style="width:200px;min-width:200px;max-width:200px;flex:0 0 200px;"
                       placeholder="取引先コードを入力"
                       value="<%= HtmlEncode(filterClientCode) %>"
                       onblur="lookupFilterClient()">
                <button type="button" class="btn btn-secondary btn-sm btn-emp-search" onclick="openClientSearch('filter')">検索</button>
                <span id="filterClientNameDisplay" class="employee-name-display<% If filterClientName <> "" Then %> employee-name-found<% End If %>"><%= HtmlEncode(filterClientName) %></span>
                <button type="submit" class="btn btn-secondary" onclick="syncFilterCode()">絞込</button>
                <% If filterClientCode <> "" Then %>
                <a href="master_project.asp" class="btn btn-secondary">クリア</a>
                <% End If %>
            </div>
        </div>
    </form>
</div>

<!-- 案件一覧 -->
<div class="card" style="padding:0; overflow-x:auto;">
    <table class="data-table">
        <thead>
            <tr>
                <th>ID</th>
                <th>取引先</th>
                <th>案件コード</th>
                <th>案件名</th>
                <th>有効</th>
                <th>操作</th>
            </tr>
        </thead>
        <tbody>
        <% If rs.EOF Then %>
            <tr><td colspan="6" style="text-align:center; padding:20px;">案件が登録されていません。</td></tr>
        <% Else %>
            <% Do While Not rs.EOF %>
            <tr>
                <td><%= rs("project_id") %></td>
                <td><%= HtmlEncode(rs("client_code")) %> : <%= HtmlEncode(rs("client_name")) %></td>
                <td><%= HtmlEncode(rs("project_code")) %></td>
                <td><%= HtmlEncode(rs("project_name")) %></td>
                <td><% If rs("is_active") = 1 Then %><span style="color:green;">有効</span><% Else %><span style="color:gray;">無効</span><% End If %></td>
                <td class="action-links">
                    <a href="master_project.asp?mode=edit&id=<%= rs("project_id") %>" class="btn btn-sm btn-warning">編集</a>
                    <form method="post" action="master_project.asp" style="display:inline;">
                        <input type="hidden" name="action" value="delete">
                        <input type="hidden" name="project_id" value="<%= rs("project_id") %>">
                        <button type="submit" class="btn btn-sm btn-danger" onclick="return confirm('この案件を削除しますか？');">削除</button>
                    </form>
                </td>
            </tr>
            <% rs.MoveNext : Loop %>
        <% End If %>
        </tbody>
    </table>
</div>

<%
If mode = "edit" And Not IsEmpty(editProject) Then CloseRecordset editProject
CloseRecordset rs
CloseDBConnection conn
%>
        </div>
    </main>
    <script src="js/common.js"></script>
    <script>
    // ============================================
    // 取引先検索モーダル
    // ============================================

    var currentClientTarget = ''; // 'form' or 'filter'
    var clientSearchModalCreated = false;
    var allClients = null; // キャッシュ

    // 取引先コード（フォーム）のonblurルックアップ
    function lookupFormClient() {
        var code = document.getElementById('formClientCode').value.trim();
        var nameDisplay = document.getElementById('formClientNameDisplay');
        if (code === '') {
            nameDisplay.textContent = '';
            nameDisplay.className = 'employee-name-display';
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
                    } else {
                        nameDisplay.textContent = '該当なし';
                        nameDisplay.className = 'employee-name-display employee-name-notfound';
                    }
                } catch (e) {
                    nameDisplay.textContent = '検索エラー';
                    nameDisplay.className = 'employee-name-display employee-name-notfound';
                }
            }
        };
        xhr.send();
    }

    // 取引先コード（フィルタ）のonblurルックアップ
    function lookupFilterClient() {
        var code = document.getElementById('filterClientCodeDisplay').value.trim();
        var nameDisplay = document.getElementById('filterClientNameDisplay');
        document.getElementById('filterClientCodeHidden').value = code;
        if (code === '') {
            nameDisplay.textContent = '';
            nameDisplay.className = 'employee-name-display';
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
                    } else {
                        nameDisplay.textContent = '該当なし';
                        nameDisplay.className = 'employee-name-display employee-name-notfound';
                    }
                } catch (e) {
                    nameDisplay.textContent = '検索エラー';
                    nameDisplay.className = 'employee-name-display employee-name-notfound';
                }
            }
        };
        xhr.send();
    }

    // 絞込ボタン押下時に表示用テキストを hidden に同期
    function syncFilterCode() {
        var display = document.getElementById('filterClientCodeDisplay');
        document.getElementById('filterClientCodeHidden').value = display.value.trim();
    }

    // モーダルHTML生成（初回のみ）
    function ensureClientSearchModal() {
        if (clientSearchModalCreated) return;
        clientSearchModalCreated = true;

        var overlay = document.createElement('div');
        overlay.id = 'clientSearchOverlay';
        overlay.style.cssText = 'display:none;position:fixed;top:0;left:0;width:100%;height:100%;background-color:rgba(0,0,0,0.5);z-index:1000;justify-content:center;align-items:center;';
        overlay.onclick = function(e) {
            if (e.target === overlay) closeClientSearch();
        };

        var modal = document.createElement('div');
        modal.style.cssText = 'background-color:#fff;border-radius:8px;box-shadow:0 4px 20px rgba(0,0,0,0.3);width:600px;max-width:90vw;max-height:80vh;display:flex;flex-direction:column;';

        modal.innerHTML =
            '<div style="display:flex;justify-content:space-between;align-items:center;padding:16px 20px;border-bottom:1px solid #eee;">' +
                '<h3 style="margin:0;font-size:18px;">取引先検索</h3>' +
                '<button type="button" style="background:none;border:none;font-size:24px;cursor:pointer;color:#999;padding:0 5px;line-height:1;" onclick="closeClientSearch()">&times;</button>' +
            '</div>' +
            '<div style="padding:20px;overflow-y:auto;flex:1;">' +
                '<div style="display:flex;gap:10px;margin-bottom:15px;">' +
                    '<input type="text" id="clientSearchKeyword" class="form-control" placeholder="取引先コードまたは取引先名を入力" style="flex:1;">' +
                    '<button type="button" class="btn btn-primary" onclick="filterClientList()">検索</button>' +
                '</div>' +
                '<div id="clientSearchResult" style="min-height:100px;max-height:400px;overflow-y:auto;">' +
                    '<p style="color:#999;text-align:center;padding:30px 0;">読み込み中...</p>' +
                '</div>' +
            '</div>';

        overlay.appendChild(modal);
        document.body.appendChild(overlay);

        document.getElementById('clientSearchKeyword').addEventListener('keydown', function(e) {
            if (e.key === 'Enter') { e.preventDefault(); filterClientList(); }
        });
    }

    // モーダルを開く
    function openClientSearch(target) {
        currentClientTarget = target;
        ensureClientSearchModal();

        document.getElementById('clientSearchKeyword').value = '';
        document.getElementById('clientSearchOverlay').style.display = 'flex';

        // 取引先一覧を読み込む（キャッシュあれば再利用）
        if (allClients !== null) {
            renderClientList(allClients);
            setTimeout(function() { document.getElementById('clientSearchKeyword').focus(); }, 100);
        } else {
            var xhr = new XMLHttpRequest();
            xhr.open('GET', 'api_project_lookup.asp?action=clients', true);
            xhr.onreadystatechange = function() {
                if (xhr.readyState === 4 && xhr.status === 200) {
                    try {
                        allClients = JSON.parse(xhr.responseText);
                        renderClientList(allClients);
                    } catch (e) {
                        document.getElementById('clientSearchResult').innerHTML =
                            '<p style="color:#c62828;text-align:center;padding:30px 0;">読み込みエラーが発生しました。</p>';
                    }
                }
            };
            xhr.send();
            setTimeout(function() { document.getElementById('clientSearchKeyword').focus(); }, 100);
        }
    }

    // モーダルを閉じる
    function closeClientSearch() {
        document.getElementById('clientSearchOverlay').style.display = 'none';
    }

    // クライアントサイドでリストをフィルタリング
    function filterClientList() {
        if (allClients === null) return;
        var kw = document.getElementById('clientSearchKeyword').value.trim().toLowerCase();
        if (kw === '') {
            renderClientList(allClients);
            return;
        }
        var filtered = allClients.filter(function(c) {
            return c.code.toLowerCase().indexOf(kw) >= 0 || c.name.toLowerCase().indexOf(kw) >= 0;
        });
        renderClientList(filtered);
    }

    // 取引先リストを描画
    function renderClientList(clients) {
        var resultDiv = document.getElementById('clientSearchResult');
        if (clients.length === 0) {
            resultDiv.innerHTML = '<p style="color:#999;text-align:center;padding:30px 0;">該当する取引先が見つかりません。</p>';
            return;
        }
        var html = '<table class="data-table" style="cursor:pointer;">' +
            '<thead><tr><th>取引先コード</th><th>取引先名</th></tr></thead><tbody>';
        for (var i = 0; i < clients.length; i++) {
            var c = clients[i];
            html += '<tr style="cursor:pointer;" onclick="selectClient(\'' +
                escapeJsStr(c.code) + '\',\'' + escapeJsStr(c.name) + '\')">' +
                '<td>' + escapeHtml(c.code) + '</td>' +
                '<td>' + escapeHtml(c.name) + '</td>' +
                '</tr>';
        }
        html += '</tbody></table>';
        resultDiv.innerHTML = html;
    }

    // 取引先を選択
    function selectClient(code, name) {
        if (currentClientTarget === 'form') {
            document.getElementById('formClientCode').value = code;
            var nameDisplay = document.getElementById('formClientNameDisplay');
            nameDisplay.textContent = name;
            nameDisplay.className = 'employee-name-display employee-name-found';
        } else if (currentClientTarget === 'filter') {
            document.getElementById('filterClientCodeDisplay').value = code;
            document.getElementById('filterClientCodeHidden').value = code;
            var nameDisplay = document.getElementById('filterClientNameDisplay');
            nameDisplay.textContent = name;
            nameDisplay.className = 'employee-name-display employee-name-found';
            // 選択後に自動で絞込実行
            document.getElementById('filterForm').submit();
        }
        closeClientSearch();
    }

    // ページ読み込み時、プリセット済みコードがあればルックアップ
    document.addEventListener('DOMContentLoaded', function() {
        if (document.getElementById('formClientCode').value.trim() !== '') {
            lookupFormClient();
        }
        if (document.getElementById('filterClientCodeDisplay').value.trim() !== '') {
            lookupFilterClient();
        }
    });

    // エスケープユーティリティ
    function escapeHtml(str) {
        if (!str) return '';
        return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    }
    function escapeJsStr(str) {
        if (!str) return '';
        return str.replace(/\\/g, '\\\\').replace(/'/g, "\\'").replace(/"/g, '\\"');
    }
    </script>
</body>
</html>
