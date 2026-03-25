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
' 案件一覧画面（参照のみ）
' ============================================

Dim conn, rs, sql
Dim filterClientCode, filterClientName
filterClientCode = Trim(Request.QueryString("client_code") & "")
filterClientName = ""

Set conn = GetClientDBConnection()

' フィルタ用取引先名を取得
If filterClientCode <> "" Then
    Dim rsFilter
    Set rsFilter = conn.Execute("SELECT client_name FROM IRAI.M_Client WHERE client_code = N'" & EscapeSQL(filterClientCode) & "'")
    If Not rsFilter.EOF Then
        filterClientName = rsFilter("client_name") & ""
    End If
    CloseRecordset rsFilter
End If

' 案件一覧取得
sql = "SELECT p.*, c.client_name FROM IRAI.M_Project p " & _
      "INNER JOIN IRAI.M_Client c ON p.client_code = c.client_code"
If filterClientCode <> "" Then
    sql = sql & " WHERE p.client_code = N'" & EscapeSQL(filterClientCode) & "'"
End If
sql = sql & " ORDER BY p.client_code, p.project_code"
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

<h2 class="page-title">案件一覧</h2>

<!-- 取引先絞り込み -->
<div class="card">
    <h3 class="card-title">絞り込み</h3>
    <form method="get" action="master_project.asp" id="filterForm">
        <div class="form-group">
            <label>取引先</label>
            <input type="hidden" name="client_code" id="filterClientCodeHidden" value="<%= HtmlEncode(filterClientCode) %>">
            <div style="display:flex;flex-direction:row;align-items:center;gap:10px;">
                <input type="text" id="filterClientCodeDisplay" class="form-control"
                       style="width:200px;"
                       placeholder="取引先コードを入力"
                       value="<%= HtmlEncode(filterClientCode) %>"
                       onblur="lookupFilterClient()">
                <button type="button" class="btn btn-secondary btn-sm" onclick="openClientSearch()">検索</button>
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
                <th>取引先</th>
                <th>案件コード</th>
                <th>案件名</th>
                <th>有効</th>
            </tr>
        </thead>
        <tbody>
        <% If rs.EOF Then %>
            <tr><td colspan="4" style="text-align:center; padding:20px;">案件が登録されていません。</td></tr>
        <% Else %>
            <% Do While Not rs.EOF %>
            <tr>
                <td><%= HtmlEncode(rs("client_code")) %> : <%= HtmlEncode(rs("client_name")) %></td>
                <td><%= HtmlEncode(rs("project_code")) %></td>
                <td><%= HtmlEncode(rs("project_name")) %></td>
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
    <script>
    var allClients = null;

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
                    nameDisplay.textContent = result.found ? result.name : '該当なし';
                    nameDisplay.className = 'employee-name-display ' + (result.found ? 'employee-name-found' : 'employee-name-notfound');
                } catch (e) {
                    nameDisplay.textContent = '検索エラー';
                    nameDisplay.className = 'employee-name-display employee-name-notfound';
                }
            }
        };
        xhr.send();
    }

    function syncFilterCode() {
        document.getElementById('filterClientCodeHidden').value =
            document.getElementById('filterClientCodeDisplay').value.trim();
    }

    function openClientSearch() {
        if (!document.getElementById('clientSearchOverlay')) {
            var overlay = document.createElement('div');
            overlay.id = 'clientSearchOverlay';
            overlay.style.cssText = 'display:none;position:fixed;top:0;left:0;width:100%;height:100%;background-color:rgba(0,0,0,0.5);z-index:1000;justify-content:center;align-items:center;';
            overlay.onclick = function(e) { if (e.target === overlay) closeClientSearch(); };
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
                    '<div id="clientSearchResult" style="min-height:100px;max-height:400px;overflow-y:auto;"><p style="color:#999;text-align:center;padding:30px 0;">読み込み中...</p></div>' +
                '</div>';
            overlay.appendChild(modal);
            document.body.appendChild(overlay);
            document.getElementById('clientSearchKeyword').addEventListener('keydown', function(e) {
                if (e.key === 'Enter') { e.preventDefault(); filterClientList(); }
            });
        }

        document.getElementById('clientSearchKeyword').value = '';
        document.getElementById('clientSearchOverlay').style.display = 'flex';

        if (allClients !== null) {
            renderClientList(allClients);
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
        }
        setTimeout(function() { document.getElementById('clientSearchKeyword').focus(); }, 100);
    }

    function closeClientSearch() {
        document.getElementById('clientSearchOverlay').style.display = 'none';
    }

    function filterClientList() {
        if (allClients === null) return;
        var kw = document.getElementById('clientSearchKeyword').value.trim().toLowerCase();
        renderClientList(kw === '' ? allClients : allClients.filter(function(c) {
            return c.code.toLowerCase().indexOf(kw) >= 0 || c.name.toLowerCase().indexOf(kw) >= 0;
        }));
    }

    function renderClientList(clients) {
        var resultDiv = document.getElementById('clientSearchResult');
        if (clients.length === 0) {
            resultDiv.innerHTML = '<p style="color:#999;text-align:center;padding:30px 0;">該当する取引先が見つかりません。</p>';
            return;
        }
        var html = '<table class="data-table" style="cursor:pointer;"><thead><tr><th>取引先コード</th><th>取引先名</th></tr></thead><tbody>';
        for (var i = 0; i < clients.length; i++) {
            var c = clients[i];
            html += '<tr style="cursor:pointer;" onclick="selectClient(\'' + escapeJsStr(c.code) + '\',\'' + escapeJsStr(c.name) + '\')">' +
                '<td>' + escapeHtml(c.code) + '</td><td>' + escapeHtml(c.name) + '</td></tr>';
        }
        resultDiv.innerHTML = html + '</tbody></table>';
    }

    function selectClient(code, name) {
        document.getElementById('filterClientCodeDisplay').value = code;
        document.getElementById('filterClientCodeHidden').value = code;
        var nameDisplay = document.getElementById('filterClientNameDisplay');
        nameDisplay.textContent = name;
        nameDisplay.className = 'employee-name-display employee-name-found';
        closeClientSearch();
        document.getElementById('filterForm').submit();
    }

    function escapeHtml(str) {
        if (!str) return '';
        return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    }
    function escapeJsStr(str) {
        if (!str) return '';
        return str.replace(/\\/g, '\\\\').replace(/'/g, "\\'").replace(/"/g, '\\"');
    }

    document.addEventListener('DOMContentLoaded', function() {
        if (document.getElementById('filterClientCodeDisplay').value.trim() !== '') {
            lookupFilterClient();
        }
    });
    </script>
</body>
</html>
