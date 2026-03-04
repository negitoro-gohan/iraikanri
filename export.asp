<%@ Language="VBScript" CodePage=65001 %>
<%
Response.CodePage = 65001
Session.CodePage = 65001
%>
<!--#include file="include/config.asp"-->
<!--#include file="include/db_connection.asp"-->
<!--#include file="include/functions.asp"-->
<%
' ============================================
' エクスポート画面
' 依頼一覧をCSVでダウンロード
' ============================================

Dim conn, rs, sql

' CSVエスケープ関数
Function EscapeCSV(val)
    Dim str
    str = SafeValue(val, "")
    str = CStr(str)
    str = Replace(str, """", """""")
    str = Replace(Replace(str, vbCrLf, " "), vbLf, " ")
    If InStr(str, ",") > 0 Or InStr(str, """") > 0 Then
        str = """" & str & """"
    End If
    EscapeCSV = str
End Function

' ダウンロード処理
If Request.QueryString("action") = "download" Then
    Dim filterStatus, filterRequester, filterAssignee
    filterStatus = SafeInt(Request.QueryString("status"), 0)
    filterRequester = SafeInt(Request.QueryString("requester"), 0)
    filterAssignee = SafeInt(Request.QueryString("assignee"), 0)

    Set conn = GetDBConnection()

    Dim whereClause
    whereClause = "WHERE r.is_deleted = 0"

    If filterStatus > 0 Then
        whereClause = whereClause & " AND r.status_id = " & filterStatus
    End If
    If filterRequester > 0 Then
        whereClause = whereClause & " AND r.requester_id = " & filterRequester
    End If
    If filterAssignee > 0 Then
        whereClause = whereClause & " AND r.assignee_id = " & filterAssignee
    End If

    sql = "SELECT r.request_id, r.request_title, req.employee_name AS requester_name, ass.employee_name AS assignee_name, " & _
          "c.client_code, c.client_name, p.project_code, p.project_name, " & _
          "r.deadline_date, s.status_name, r.request_content, r.response_content, " & _
          "r.end_date, r.created_at, r.updated_at " & _
          "FROM IRAI.T_Request r " & _
          "INNER JOIN IRAI.M_Employee req ON r.requester_id = req.employee_id " & _
          "INNER JOIN IRAI.M_Employee ass ON r.assignee_id = ass.employee_id " & _
          "INNER JOIN M_Status s ON r.status_id = s.status_id " & _
          "LEFT JOIN IRAI.M_Client c ON r.client_id = c.client_id " & _
          "LEFT JOIN IRAI.M_Project p ON r.project_id = p.project_id " & _
          whereClause & " ORDER BY r.deadline_date ASC"
    Set rs = conn.Execute(sql)

    Response.Clear
    Response.ContentType = "text/csv"
    Response.Charset = "UTF-8"
    Response.AddHeader "Content-Disposition", "attachment; filename=requests_" & Year(Now()) & Right("0" & Month(Now()), 2) & Right("0" & Day(Now()), 2) & ".csv"

    Response.BinaryWrite ChrB(&HEF) & ChrB(&HBB) & ChrB(&HBF)

    Response.Write "依頼ID,依頼件名,依頼元,依頼先,取引先コード,取引先名,案件コード,案件名,期限日,ステータス,依頼内容,対応内容,終了日,登録日時,更新日時" & vbCrLf

    Do While Not rs.EOF
        Response.Write EscapeCSV(rs("request_id")) & ","
        Response.Write EscapeCSV(rs("request_title")) & ","
        Response.Write EscapeCSV(rs("requester_name")) & ","
        Response.Write EscapeCSV(rs("assignee_name")) & ","
        Response.Write EscapeCSV(rs("client_code")) & ","
        Response.Write EscapeCSV(rs("client_name")) & ","
        Response.Write EscapeCSV(rs("project_code")) & ","
        Response.Write EscapeCSV(rs("project_name")) & ","
        Response.Write EscapeCSV(FormatDateJP(rs("deadline_date"))) & ","
        Response.Write EscapeCSV(rs("status_name")) & ","
        Response.Write EscapeCSV(rs("request_content")) & ","
        Response.Write EscapeCSV(rs("response_content")) & ","
        Response.Write EscapeCSV(FormatDateJP(rs("end_date"))) & ","
        Response.Write EscapeCSV(rs("created_at")) & ","
        Response.Write EscapeCSV(rs("updated_at")) & vbCrLf
        rs.MoveNext
    Loop

    CloseRecordset rs
    CloseDBConnection conn
    Response.End
End If

Response.Charset = "UTF-8"
Response.ContentType = "text/html; charset=UTF-8"

Set conn = GetDBConnection()

Dim rsEmployee
Set rsEmployee = conn.Execute("SELECT employee_id, employee_name FROM IRAI.M_Employee WHERE is_active = 1 ORDER BY employee_name")

' 社員リストを配列に格納
Dim empIds(), empNames(), empCount
empCount = 0
Do While Not rsEmployee.EOF
    ReDim Preserve empIds(empCount)
    ReDim Preserve empNames(empCount)
    empIds(empCount) = rsEmployee("employee_id")
    empNames(empCount) = rsEmployee("employee_name") & ""
    empCount = empCount + 1
    rsEmployee.MoveNext
Loop
CloseRecordset rsEmployee
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

<h2 class="page-title">エクスポート</h2>

<div class="card">
    <h3 class="card-title">CSVダウンロード</h3>
    <p style="margin-bottom: 15px;">依頼一覧をCSV形式でダウンロードします。</p>

    <form method="get" action="export.asp">
        <input type="hidden" name="action" value="download">

        <div class="form-group">
            <label>ステータス</label>
            <select name="status" class="form-control">
                <option value="0">-- すべて --</option>
                <option value="<%= STATUS_NOT_STARTED %>">未着手</option>
                <option value="<%= STATUS_IN_PROGRESS %>">着手済</option>
                <option value="<%= STATUS_COMPLETED %>">対応完了</option>
                <option value="<%= STATUS_NOT_APPLICABLE %>">対象外</option>
            </select>
        </div>

        <div class="form-group">
            <label>依頼元</label>
            <select name="requester" class="form-control">
                <option value="0">-- すべて --</option>
                <% Dim ei : For ei = 0 To empCount - 1 %>
                <option value="<%= empIds(ei) %>"><%= HtmlEncode(empNames(ei)) %></option>
                <% Next %>
            </select>
        </div>

        <div class="form-group">
            <label>依頼先</label>
            <select name="assignee" class="form-control">
                <option value="0">-- すべて --</option>
                <% Dim ej : For ej = 0 To empCount - 1 %>
                <option value="<%= empIds(ej) %>"><%= HtmlEncode(empNames(ej)) %></option>
                <% Next %>
            </select>
        </div>

        <div class="btn-group">
            <button type="submit" class="btn btn-primary">CSVダウンロード</button>
        </div>
    </form>
</div>

<%
CloseDBConnection conn
%>
        </div>
    </main>
    <script src="js/common.js"></script>
</body>
</html>
