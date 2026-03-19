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

    responseContent = Trim(Request.Form("response_content") & "")
    endDate         = Trim(Request.Form("end_date")         & "")

    ' 入力チェック
    If endDate <> "" And Not IsDate(endDate) Then
        errorMsg = "終了日の形式が正しくありません。"
    ElseIf endDate <> "" And responseContent = "" Then
        errorMsg = "終了日を入力する場合は、対応内容も入力してください。"
    Else
        ' ステータスを入力内容から自動決定
        ' 対応終了：終了日あり かつ 対応内容あり
        ' 着手済　：終了日なし かつ 対応内容あり
        ' 未着手　：終了日なし かつ 対応内容なし
        If endDate <> "" And responseContent <> "" Then
            statusId = STATUS_COMPLETED
        ElseIf responseContent <> "" Then
            statusId = STATUS_IN_PROGRESS
        Else
            statusId = STATUS_NOT_STARTED
        End If

        ' ドメインユーザー名を取得
        Dim domainUser
        domainUser = GetCurrentUser()

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

' 依頼データ取得（社員名はT_Requestに保存済みのカラムを使用）
sql = "SELECT r.* FROM IRAI.T_Request r " & _
      "WHERE r.request_id = " & requestId & " AND r.is_deleted = 0"
Set rs = conn.Execute(sql)

If rs.EOF Then
    CloseRecordset rs
    CloseDBConnection conn
    SetMessage "error", "指定された依頼が見つかりません。"
    Response.Redirect "request_list.asp"
    Response.End
End If

' 社員コードをEmployee DBから取得
Dim requesterCode, assigneeCode
requesterCode = ""
assigneeCode  = ""
Dim connEmpResp, rsEmpResp
Set connEmpResp = GetEmployeeDBConnection()
Set rsEmpResp = connEmpResp.Execute("SELECT employee_code FROM IRAI.M_Employee WHERE employee_id = " & rs("requester_id"))
If Not rsEmpResp.EOF Then requesterCode = rsEmpResp("employee_code") & ""
CloseRecordset rsEmpResp
Set rsEmpResp = connEmpResp.Execute("SELECT employee_code FROM IRAI.M_Employee WHERE employee_id = " & rs("assignee_id"))
If Not rsEmpResp.EOF Then assigneeCode = rsEmpResp("employee_code") & ""
CloseRecordset rsEmpResp
CloseDBConnection connEmpResp
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
            <td>
                <% If requesterCode <> "" Then %><%= HtmlEncode(requesterCode) %> : <% End If %>
                <%= HtmlEncode(SafeValue(rs("requester_name"), "（不明）")) %>
                <% If Not IsNull(rs("requester_email")) And rs("requester_email") <> "" Then %>
                （<%= HtmlEncode(rs("requester_email")) %>）
                <% End If %>
            </td>
        </tr>
        <tr>
            <th>依頼先</th>
            <td>
                <% If assigneeCode <> "" Then %><%= HtmlEncode(assigneeCode) %> : <% End If %>
                <%= HtmlEncode(SafeValue(rs("assignee_name"), "（不明）")) %>
                <% If Not IsNull(rs("assignee_email")) And rs("assignee_email") <> "" Then %>
                （<%= HtmlEncode(rs("assignee_email")) %>）
                <% End If %>
            </td>
        </tr>
        <tr>
            <th>取引先</th>
            <td>
            <% If Not IsNull(rs("client_name")) And rs("client_name") <> "" Then %>
                <%= HtmlEncode(rs("client_code")) %> : <%= HtmlEncode(rs("client_name")) %>
            <% Else %>
                -
            <% End If %>
            </td>
        </tr>
        <tr>
            <th>案件</th>
            <td>
            <% If Not IsNull(rs("project_name")) And rs("project_name") <> "" Then %>
                <%= HtmlEncode(rs("project_code")) %> : <%= HtmlEncode(rs("project_name")) %>
            <% Else %>
                -
            <% End If %>
            </td>
        </tr>
        <tr>
            <th>期限日</th>
            <td><%= FormatDateJP(rs("deadline_date")) %></td>
        </tr>
        <tr>
            <th>依頼内容</th>
            <td><%= NL2BR(HtmlEncode(SafeValue(rs("request_content"), "-"))) %></td>
        </tr>
        <tr>
            <th>最終更新者</th>
            <td><%= HtmlEncode(SafeValue(rs("updated_by"), "-")) %></td>
        </tr>
        <% If Not IsNull(rs("folder_address")) And rs("folder_address") <> "" Then %>
        <tr>
            <th>フォルダアドレス</th>
            <td>
                <div style="display:flex;gap:8px;align-items:center;">
                    <span id="folderAddressText"><%= HtmlEncode(rs("folder_address")) %></span>
                    <button type="button" class="btn btn-secondary btn-sm" onclick="copyFolderAddress()" style="white-space:nowrap;">コピー</button>
                </div>
            </td>
        </tr>
        <% End If %>
    </table>
</div>

<!-- 対応入力フォーム -->
<div class="card">
    <h3 class="card-title">対応入力</h3>
    <form method="post" action="request_response.asp">
        <input type="hidden" name="request_id" value="<%= requestId %>">

        <!-- 対応内容 -->
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
            <textarea name="response_content" id="response_content" class="form-control" rows="5" oninput="updateStatusPreview()"><%= HtmlEncode(respContent) %></textarea>
        </div>

        <!-- 終了日 -->
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
            <input type="date" name="end_date" id="end_date" class="form-control" value="<%= respEndDate %>" onchange="updateStatusPreview()">
        </div>

        <!-- ステータス（自動判定・表示のみ） -->
        <div class="form-group">
            <label>ステータス（自動判定）</label>
            <div id="statusPreview" style="padding:6px 0;font-weight:bold;"></div>
            <small style="color:#666;">対応内容・終了日の入力状況により自動で決まります</small>
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
    <script>
    // ============================================================
    // ステータス自動判定プレビュー
    // 終了日・対応内容の入力状況からリアルタイムにステータスを表示
    // ============================================================
    function updateStatusPreview() {
        var responseContent = document.getElementById('response_content').value.trim();
        var endDate         = document.getElementById('end_date').value.trim();
        var preview         = document.getElementById('statusPreview');

        var label, color;
        if (endDate !== '' && responseContent !== '') {
            label = '対応終了';
            color = '#27ae60';
        } else if (responseContent !== '') {
            label = '着手済';
            color = '#2980b9';
        } else {
            label = '未着手';
            color = '#7f8c8d';
        }
        preview.textContent = label;
        preview.style.color = color;
    }

    // 初期表示
    document.addEventListener('DOMContentLoaded', function() {
        updateStatusPreview();
    });

    // フォルダアドレス コピー
    function copyFolderAddress() {
        var text = document.getElementById('folderAddressText').textContent.trim();
        if (!text) return;
        if (navigator.clipboard && window.isSecureContext) {
            navigator.clipboard.writeText(text).then(function() {
                showCopyToast();
            });
        } else {
            // フォールバック（HTTP環境など）
            var tmp = document.createElement('textarea');
            tmp.value = text;
            tmp.style.position = 'fixed';
            tmp.style.opacity = '0';
            document.body.appendChild(tmp);
            tmp.select();
            document.execCommand('copy');
            document.body.removeChild(tmp);
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
    </script>
</body>
</html>
