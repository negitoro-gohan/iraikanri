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
' 依頼一覧画面
' フィルタリング・ソート・ページネーション対応
' ============================================

Dim conn, rs, sql
Dim filterStatus, filterRequester, filterAssignee, filterKeyword, filterPreset, filterDeadlineFrom, filterDeadlineTo, filterClientCode, filterProjectCode, filterMyAssignee
Dim sortColumn, sortDir
Dim currentPage, totalRecords, totalPages, startRecord

Set conn = GetDBConnection()

' フィルタパラメータ取得
filterStatus = SafeInt(Request.QueryString("status"), 0)
filterRequester = Trim(Request.QueryString("requester") & "")
filterAssignee = Trim(Request.QueryString("assignee") & "")
filterKeyword = Trim(Request.QueryString("keyword") & "")
filterDeadlineFrom = Trim(Request.QueryString("deadline_from") & "")
filterDeadlineTo = Trim(Request.QueryString("deadline_to") & "")
filterClientCode = Trim(Request.QueryString("client_code") & "")
filterProjectCode = Trim(Request.QueryString("project_code") & "")
filterPreset    = Trim(Request.QueryString("filter")      & "")
filterMyAssignee = Trim(Request.QueryString("my_assignee") & "")

' ソートパラメータ取得
sortColumn = Request.QueryString("sort") & ""
If sortColumn = "" Then sortColumn = "deadline_date"
sortDir = Request.QueryString("dir") & ""
If sortDir = "" Then sortDir = "ASC"

' ソート方向の正規化
If UCase(sortDir) <> "DESC" Then sortDir = "ASC"

' ソートカラムの正規化（SQLインジェクション対策）とテーブル修飾マッピング
' ※ ROW_NUMBER の ORDER BY で複数テーブルに同名カラムがある場合のあいまい参照を防止
Dim sortColumnSQL
Select Case LCase(sortColumn)
    Case "request_id"
        sortColumnSQL = "r.request_id"
    Case "request_title"
        sortColumnSQL = "r.request_title"
    Case "requester_name"
        sortColumnSQL = "r.requester_name"
    Case "assignee_name"
        sortColumnSQL = "r.assignee_name"
    Case "deadline_date"
        sortColumnSQL = "r.deadline_date"
    Case "status_id"
        sortColumnSQL = "r.status_id"
    Case "created_at"
        sortColumnSQL = "r.created_at"        ' IRAI.M_Client / IRAI.M_Project にも同名カラムがあるため必ずテーブル修飾が必要
    Case Else
        sortColumn    = "deadline_date"
        sortColumnSQL = "r.deadline_date"
End Select

' ページネーション
currentPage = SafeInt(Request.QueryString("page"), 1)
If currentPage < 1 Then currentPage = 1

' WHERE句の構築
Dim whereClause
whereClause = "WHERE r.is_deleted = 0"

' プリセットフィルタ
Select Case filterPreset
    Case "overdue"
        whereClause = whereClause & " AND r.status_id NOT IN (" & STATUS_COMPLETED & "," & STATUS_NOT_APPLICABLE & ") AND r.deadline_date < CONVERT(DATE, GETDATE())"
    Case "soon"
        whereClause = whereClause & " AND r.status_id NOT IN (" & STATUS_COMPLETED & "," & STATUS_NOT_APPLICABLE & ") AND r.deadline_date >= CONVERT(DATE, GETDATE()) AND r.deadline_date <= DATEADD(day, " & REMIND_DAYS_BEFORE & ", CONVERT(DATE, GETDATE()))"
    Case "in_progress"
        whereClause = whereClause & " AND r.status_id = " & STATUS_IN_PROGRESS
End Select

' 個別フィルタ
If filterStatus > 0 Then
    whereClause = whereClause & " AND r.status_id = " & filterStatus
End If
' 社員コードフィルタ: Employee DBから employee_id を先に取得してWHEREに埋め込む
If filterRequester <> "" Then
    Dim connEmpFilter, rsEmpFilter, filterRequesterId
    filterRequesterId = 0
    Set connEmpFilter = GetEmployeeDBConnection()
    Set rsEmpFilter = connEmpFilter.Execute("SELECT employee_id FROM IRAI.M_Employee WHERE employee_code = N'" & EscapeSQL(filterRequester) & "' AND is_active = 1")
    If Not rsEmpFilter.EOF Then filterRequesterId = CLng(rsEmpFilter("employee_id"))
    CloseRecordset rsEmpFilter
    CloseDBConnection connEmpFilter
    If filterRequesterId > 0 Then
        whereClause = whereClause & " AND r.requester_id = " & filterRequesterId
    Else
        whereClause = whereClause & " AND 1 = 0" ' 該当なし → 0件
    End If
End If
If filterAssignee <> "" Then
    Dim connEmpFilter2, rsEmpFilter2, filterAssigneeId
    filterAssigneeId = 0
    Set connEmpFilter2 = GetEmployeeDBConnection()
    Set rsEmpFilter2 = connEmpFilter2.Execute("SELECT employee_id FROM IRAI.M_Employee WHERE employee_code = N'" & EscapeSQL(filterAssignee) & "' AND is_active = 1")
    If Not rsEmpFilter2.EOF Then filterAssigneeId = CLng(rsEmpFilter2("employee_id"))
    CloseRecordset rsEmpFilter2
    CloseDBConnection connEmpFilter2
    If filterAssigneeId > 0 Then
        whereClause = whereClause & " AND r.assignee_id = " & filterAssigneeId
    Else
        whereClause = whereClause & " AND 1 = 0" ' 該当なし → 0件
    End If
End If
If filterKeyword <> "" Then
    whereClause = whereClause & " AND (r.request_title LIKE N'%" & EscapeSQL(filterKeyword) & "%' OR r.request_content LIKE N'%" & EscapeSQL(filterKeyword) & "%')"
End If
If filterDeadlineFrom <> "" And IsDate(filterDeadlineFrom) Then
    whereClause = whereClause & " AND r.deadline_date >= '" & EscapeSQL(filterDeadlineFrom) & "'"
End If
If filterDeadlineTo <> "" And IsDate(filterDeadlineTo) Then
    whereClause = whereClause & " AND r.deadline_date <= '" & EscapeSQL(filterDeadlineTo) & "'"
End If
If filterClientCode <> "" Then
    whereClause = whereClause & " AND r.client_id = " & SafeInt(filterClientCode, 0)
End If
If filterProjectCode <> "" Then
    whereClause = whereClause & " AND r.project_id = " & SafeInt(filterProjectCode, 0)
End If
' my_assignee=1: ログインユーザーが依頼先の依頼のみ表示
If filterMyAssignee = "1" Then
    Dim connEmpMyMe, rsEmpMyMe, myAssigneeId
    myAssigneeId = 0
    Set connEmpMyMe = GetEmployeeDBConnection()
    Set rsEmpMyMe = connEmpMyMe.Execute("SELECT employee_id FROM IRAI.M_Employee WHERE employee_code = N'" & EscapeSQL(GetCurrentUser()) & "' AND is_active = 1")
    If Not rsEmpMyMe.EOF Then myAssigneeId = CLng(rsEmpMyMe("employee_id"))
    CloseRecordset rsEmpMyMe
    CloseDBConnection connEmpMyMe
    If myAssigneeId > 0 Then
        whereClause = whereClause & " AND r.assignee_id = " & myAssigneeId
    Else
        whereClause = whereClause & " AND 1 = 0" ' 社員マスターに未登録のため0件
    End If
End If

' 総件数取得
sql = "SELECT COUNT(*) FROM IRAI.T_Request r " & whereClause
Set rs = conn.Execute(sql)
totalRecords = rs(0)
CloseRecordset rs

' ページ数計算
totalPages = 0
If totalRecords > 0 Then
    totalPages = Int((totalRecords + ITEMS_PER_PAGE - 1) / ITEMS_PER_PAGE)
End If
If currentPage > totalPages And totalPages > 0 Then currentPage = totalPages
startRecord = (currentPage - 1) * ITEMS_PER_PAGE

' データ取得（ページネーション対応）
sql = "SELECT * FROM (" & _
      "SELECT r.request_id, r.request_title, r.deadline_date, r.status_id, r.created_at, " & _
      "r.client_id, r.project_id, r.client_code, r.client_name, r.project_code, r.project_name, " & _
      "r.requester_name, r.assignee_name, " & _
      "ROW_NUMBER() OVER (ORDER BY " & sortColumnSQL & " " & sortDir & ") AS RowNum " & _
      "FROM IRAI.T_Request r " & _
      whereClause & _
      ") AS t WHERE RowNum > " & startRecord & " AND RowNum <= " & (startRecord + ITEMS_PER_PAGE)
Set rs = conn.Execute(sql)

' 取引先一覧取得（フィルタ用、Client DBへ接続）
Dim rsClients, clientIds(), clientCodes(), clientNames(), clientCount
Dim connClientList
Set connClientList = GetClientDBConnection()
sql = "SELECT client_id, client_code, client_name FROM IRAI.M_Client WHERE is_active = 1 ORDER BY client_code"
Set rsClients = connClientList.Execute(sql)
clientCount = 0
Do While Not rsClients.EOF
    ReDim Preserve clientIds(clientCount)
    ReDim Preserve clientCodes(clientCount)
    ReDim Preserve clientNames(clientCount)
    clientIds(clientCount) = rsClients("client_id")
    clientCodes(clientCount) = rsClients("client_code") & ""
    clientNames(clientCount) = rsClients("client_name") & ""
    clientCount = clientCount + 1
    rsClients.MoveNext
Loop
CloseRecordset rsClients
CloseDBConnection connClientList

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
            <h1 class="site-title">
                <a href="index.asp"><%= APP_TITLE %></a>
            </h1>
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

<h2 class="page-title">依頼一覧
<%
Dim pageTitleSub
pageTitleSub = ""
Select Case filterPreset
    Case "overdue"      : pageTitleSub = "期限切れ"
    Case "soon"         : pageTitleSub = "期限間近"
    Case "in_progress"  : pageTitleSub = "着手済"
End Select
If filterMyAssignee = "1" Then
    If pageTitleSub <> "" Then pageTitleSub = pageTitleSub & " / "
    pageTitleSub = pageTitleSub & "自分が依頼先"
End If
If pageTitleSub <> "" Then
    Response.Write " <small style=""font-size:0.6em;font-weight:normal;color:#666;"">（" & HtmlEncode(pageTitleSub) & "）</small>"
End If
%>
</h2>

<!-- フィルタエリア -->
<form method="get" action="request_list.asp" class="filter-area">
    <div class="filter-group">
        <label>期限フィルタ</label>
        <select name="filter" class="form-control">
            <option value="">-- なし --</option>
            <option value="overdue"     <% If filterPreset = "overdue"      Then Response.Write "selected" End If %>>期限切れ</option>
            <option value="soon"        <% If filterPreset = "soon"         Then Response.Write "selected" End If %>>期限間近（<%= REMIND_DAYS_BEFORE %>日以内）</option>
            <option value="in_progress" <% If filterPreset = "in_progress"  Then Response.Write "selected" End If %>>着手済</option>
        </select>
    </div>

    <div class="filter-group">
        <label>対象</label>
        <label style="display:flex;align-items:center;gap:6px;font-weight:normal;cursor:pointer;padding-top:4px;">
            <input type="checkbox" name="my_assignee" value="1" <% If filterMyAssignee = "1" Then Response.Write "checked" End If %>>
            自分が依頼先のみ
        </label>
    </div>

    <div class="filter-group">
        <label>ステータス</label>
        <select name="status" class="form-control">
            <option value="0">-- すべて --</option>
            <option value="<%= STATUS_NOT_STARTED %>" <% If filterStatus = STATUS_NOT_STARTED Then Response.Write " selected" End If %>>未着手</option>
            <option value="<%= STATUS_IN_PROGRESS %>" <% If filterStatus = STATUS_IN_PROGRESS Then Response.Write " selected" End If %>>着手済</option>
            <option value="<%= STATUS_COMPLETED %>" <% If filterStatus = STATUS_COMPLETED Then Response.Write " selected" End If %>>対応終了</option>
            <option value="<%= STATUS_NOT_APPLICABLE %>" <% If filterStatus = STATUS_NOT_APPLICABLE Then Response.Write " selected" End If %>>対象外</option>
        </select>
    </div>

    <div class="filter-group">
        <label>依頼元（社員コード）</label>
        <div class="employee-code-group" style="display:flex;flex-direction:row;flex-wrap:nowrap;align-items:center;gap:6px;">
            <input type="text" name="requester" id="filter_requester_code" class="form-control" style="width:140px;min-width:140px;" maxlength="20" placeholder="社員コード"
                   value="<%= HtmlEncode(filterRequester) %>"
                   onblur="lookupFilterEmployee('requester')">
            <button type="button" class="btn btn-secondary btn-sm btn-emp-search" onclick="openFilterEmployeeSearch('requester')">検索</button>
            <span id="filter_requester_name" class="employee-name-display"></span>
        </div>
    </div>

    <div class="filter-group">
        <label>依頼先（社員コード）</label>
        <div class="employee-code-group" style="display:flex;flex-direction:row;flex-wrap:nowrap;align-items:center;gap:6px;">
            <input type="text" name="assignee" id="filter_assignee_code" class="form-control" style="width:140px;min-width:140px;" maxlength="20" placeholder="社員コード"
                   value="<%= HtmlEncode(filterAssignee) %>"
                   onblur="lookupFilterEmployee('assignee')">
            <button type="button" class="btn btn-secondary btn-sm btn-emp-search" onclick="openFilterEmployeeSearch('assignee')">検索</button>
            <span id="filter_assignee_name" class="employee-name-display"></span>
        </div>
    </div>

    <div class="filter-group">
        <label>期限日（開始）</label>
        <input type="date" name="deadline_from" class="form-control" value="<%= HtmlEncode(filterDeadlineFrom) %>">
    </div>

    <div class="filter-group">
        <label>期限日（終了）</label>
        <input type="date" name="deadline_to" class="form-control" value="<%= HtmlEncode(filterDeadlineTo) %>">
    </div>

    <div class="filter-group">
        <label>取引先</label>
        <select name="client_code" id="filter_client_id" class="form-control" style="width:200px;" onchange="loadFilterProjects();">
            <option value="">-- すべて --</option>
            <% Dim fci : For fci = 0 To clientCount - 1 %>
            <option value="<%= clientIds(fci) %>" <% If CStr(filterClientCode) = CStr(clientIds(fci)) Then %>selected<% End If %>><%= HtmlEncode(clientCodes(fci)) %> : <%= HtmlEncode(clientNames(fci)) %></option>
            <% Next %>
        </select>
    </div>

    <div class="filter-group">
        <label>案件</label>
        <select name="project_code" id="filter_project_id" class="form-control" style="width:200px;">
            <option value="">-- すべて --</option>
        </select>
    </div>

    <div class="filter-group">
        <label>キーワード</label>
        <input type="text" name="keyword" class="form-control" value="<%= HtmlEncode(filterKeyword) %>" placeholder="件名・内容で検索">
    </div>

    <div class="filter-group">
        <label>&nbsp;</label>
        <button type="submit" class="btn btn-primary">検索</button>
        <a href="request_list.asp" class="btn btn-secondary">クリア</a>
    </div>
</form>

<!-- 件数表示 -->
<p style="margin-bottom: 10px;">全 <%= totalRecords %> 件</p>

<!-- 依頼一覧テーブル -->
<div class="card" style="padding: 0; overflow-x: auto;">
    <table class="data-table" id="requestTable">
        <thead>
            <tr>
                <th><a href="request_list.asp?sort=request_id&dir=<% If sortColumn = "request_id" And sortDir = "ASC" Then Response.Write "DESC" Else Response.Write "ASC" End If %>">ID<% If sortColumn = "request_id" Then If sortDir = "ASC" Then Response.Write " ▲" Else Response.Write " ▼" End If End If %></a></th>
                <th><a href="request_list.asp?sort=request_title&dir=<% If sortColumn = "request_title" And sortDir = "ASC" Then Response.Write "DESC" Else Response.Write "ASC" End If %>">依頼件名<% If sortColumn = "request_title" Then If sortDir = "ASC" Then Response.Write " ▲" Else Response.Write " ▼" End If End If %></a></th>
                <th><a href="request_list.asp?sort=requester_name&dir=<% If sortColumn = "requester_name" And sortDir = "ASC" Then Response.Write "DESC" Else Response.Write "ASC" End If %>">依頼元<% If sortColumn = "requester_name" Then If sortDir = "ASC" Then Response.Write " ▲" Else Response.Write " ▼" End If End If %></a></th>
                <th><a href="request_list.asp?sort=assignee_name&dir=<% If sortColumn = "assignee_name" And sortDir = "ASC" Then Response.Write "DESC" Else Response.Write "ASC" End If %>">依頼先<% If sortColumn = "assignee_name" Then If sortDir = "ASC" Then Response.Write " ▲" Else Response.Write " ▼" End If End If %></a></th>
                <th>取引先</th>
                <th>案件</th>
                <th><a href="request_list.asp?sort=deadline_date&dir=<% If sortColumn = "deadline_date" And sortDir = "ASC" Then Response.Write "DESC" Else Response.Write "ASC" End If %>">期限日<% If sortColumn = "deadline_date" Then If sortDir = "ASC" Then Response.Write " ▲" Else Response.Write " ▼" End If End If %></a></th>
                <th><a href="request_list.asp?sort=status_id&dir=<% If sortColumn = "status_id" And sortDir = "ASC" Then Response.Write "DESC" Else Response.Write "ASC" End If %>">ステータス<% If sortColumn = "status_id" Then If sortDir = "ASC" Then Response.Write " ▲" Else Response.Write " ▼" End If End If %></a></th>
                <th><a href="request_list.asp?sort=created_at&dir=<% If sortColumn = "created_at" And sortDir = "ASC" Then Response.Write "DESC" Else Response.Write "ASC" End If %>">登録日<% If sortColumn = "created_at" Then If sortDir = "ASC" Then Response.Write " ▲" Else Response.Write " ▼" End If End If %></a></th>
                <th>操作</th>
            </tr>
        </thead>
        <tbody>
        <%
        If rs.EOF Then
        %>
            <tr>
                <td colspan="10" style="text-align: center; padding: 30px;">該当する依頼がありません。</td>
            </tr>
        <%
        Else
            Do While Not rs.EOF
                Dim dlClass
                dlClass = GetDeadlineClass(rs("deadline_date"), rs("status_id"))
        %>
            <tr>
                <td><%= rs("request_id") %></td>
                <td><a href="request_detail.asp?id=<%= rs("request_id") %>"><%= HtmlEncode(rs("request_title")) %></a></td>
                <td><%= HtmlEncode(rs("requester_name")) %></td>
                <td><%= HtmlEncode(rs("assignee_name")) %></td>
                <td><% If Not IsNull(rs("client_code")) And rs("client_code") <> "" Then %><%= HtmlEncode(rs("client_code")) %><% End If %></td>
                <td><% If Not IsNull(rs("project_code")) And rs("project_code") <> "" Then %><%= HtmlEncode(rs("project_code")) %><% End If %></td>
                <td class="<%= dlClass %>"><%= FormatDateJP(rs("deadline_date")) %></td>
                <td><span class="status-badge <%= GetStatusClass(rs("status_id")) %>"><%= GetStatusName(rs("status_id")) %></span></td>
                <td><%= FormatDateJP(rs("created_at")) %></td>
                <td class="action-links">
                    <a href="request_detail.asp?id=<%= rs("request_id") %>" class="btn btn-sm btn-primary">詳細</a>
                    <a href="request_edit.asp?id=<%= rs("request_id") %>" class="btn btn-sm btn-warning">編集</a>
                </td>
            </tr>
        <%
                rs.MoveNext
            Loop
        End If
        %>
        </tbody>
    </table>
</div>

<!-- ページネーション -->
<% If totalPages > 1 Then %>
<div class="pagination">
    <%
    ' 前へ
    If currentPage > 1 Then
    %>
    <a href="request_list.asp?page=<%= currentPage - 1 %><% If filterMyAssignee = "1" Then %>&my_assignee=1<% End If %>">&laquo; 前へ</a>
    <% End If %>

    <%
    ' ページ番号
    Dim i
    For i = 1 To totalPages
        If i = currentPage Then
    %>
    <span class="current"><%= i %></span>
    <%
        ElseIf Abs(i - currentPage) <= 2 Or i = 1 Or i = totalPages Then
    %>
    <a href="request_list.asp?page=<%= i %><% If filterMyAssignee = "1" Then %>&my_assignee=1<% End If %>"><%= i %></a>
    <%
        ElseIf Abs(i - currentPage) = 3 Then
    %>
    <span>...</span>
    <%
        End If
    Next
    %>

    <%
    ' 次へ
    If currentPage < totalPages Then
    %>
    <a href="request_list.asp?page=<%= currentPage + 1 %><% If filterMyAssignee = "1" Then %>&my_assignee=1<% End If %>">次へ &raquo;</a>
    <% End If %>
</div>
<% End If %>

<%
CloseRecordset rs
CloseDBConnection conn
%>
        </div>
    </main>
    <script src="js/common.js"></script>
    <script src="js/employee_search.js"></script>
    <script>
    // 一覧フィルタ用：社員コードから社員名を表示
    function lookupFilterEmployee(type) {
        var codeInput, nameDisplay;
        if (type === 'requester') {
            codeInput = document.getElementById('filter_requester_code');
            nameDisplay = document.getElementById('filter_requester_name');
        } else {
            codeInput = document.getElementById('filter_assignee_code');
            nameDisplay = document.getElementById('filter_assignee_name');
        }
        var code = codeInput.value.trim();
        if (code === '') {
            nameDisplay.textContent = '';
            nameDisplay.className = 'employee-name-display';
            return;
        }
        var xhr = new XMLHttpRequest();
        xhr.open('GET', 'api_employee_lookup.asp?code=' + encodeURIComponent(code), true);
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

    // 一覧フィルタ用：モーダル検索を開く
    function openFilterEmployeeSearch(type) {
        // employee_search.js のモーダルを流用するため、
        // 選択時のコールバックを一時的に差し替える
        currentTargetField = '__filter_' + type;
        ensureEmployeeSearchModal();
        document.getElementById('empSearchKeyword').value = '';
        document.getElementById('empSearchResult').innerHTML =
            '<p style="color:#999;text-align:center;padding:30px 0;">キーワードを入力して検索してください。</p>';
        document.getElementById('empSearchOverlay').style.display = 'flex';
        setTimeout(function() { document.getElementById('empSearchKeyword').focus(); }, 100);
    }

    // selectEmployee をオーバーライド（フィルタ用の分岐を追加）
    var _originalSelectEmployee = selectEmployee;
    selectEmployee = function(code, name) {
        if (currentTargetField === '__filter_requester') {
            document.getElementById('filter_requester_code').value = code;
            var d = document.getElementById('filter_requester_name');
            d.textContent = name;
            d.className = 'employee-name-display employee-name-found';
            closeEmployeeSearch();
        } else if (currentTargetField === '__filter_assignee') {
            document.getElementById('filter_assignee_code').value = code;
            var d2 = document.getElementById('filter_assignee_name');
            d2.textContent = name;
            d2.className = 'employee-name-display employee-name-found';
            closeEmployeeSearch();
        } else {
            _originalSelectEmployee(code, name);
        }
    };

    // 取引先選択時に案件一覧を読み込む（フィルタ用）
    function loadFilterProjects(selectedProjectId) {
        var clientId = document.getElementById('filter_client_id').value;
        var projectSelect = document.getElementById('filter_project_id');

        projectSelect.innerHTML = '<option value="">-- すべて --</option>';

        if (!clientId) {
            return;
        }

        var xhr = new XMLHttpRequest();
        xhr.open('GET', 'api_project_lookup.asp?action=projects&client_id=' + clientId, true);
        xhr.onreadystatechange = function() {
            if (xhr.readyState === 4 && xhr.status === 200) {
                try {
                    var projects = JSON.parse(xhr.responseText);
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

    // ページ読み込み時、フィルタに入力済みのコードがあれば名前を表示
    document.addEventListener('DOMContentLoaded', function() {
        if (document.getElementById('filter_requester_code') &&
            document.getElementById('filter_requester_code').value.trim() !== '') {
            lookupFilterEmployee('requester');
        }
        if (document.getElementById('filter_assignee_code') &&
            document.getElementById('filter_assignee_code').value.trim() !== '') {
            lookupFilterEmployee('assignee');
        }
        // 取引先が選択されていれば案件を読み込む
        var clientId = document.getElementById('filter_client_id').value;
        if (clientId) {
            loadFilterProjects(<%= SafeInt(filterProjectCode, 0) %>);
        }
    });
    </script>
</body>
</html>
