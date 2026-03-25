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
' インポート画面
' ============================================

Dim conn, rs, sql
Dim errorMsg, successCount, errorCount
errorMsg = ""
successCount = 0
errorCount = 0

' IIf関数
Function IIf(condition, trueValue, falseValue)
    If condition Then
        IIf = trueValue
    Else
        IIf = falseValue
    End If
End Function

' CSVパース関数
Function ParseCSVLine(line)
    Dim result(), field, inQuotes, i, c, fieldIndex
    ReDim result(0)
    field = ""
    inQuotes = False
    fieldIndex = 0

    For i = 1 To Len(line)
        c = Mid(line, i, 1)
        If c = """" Then
            If inQuotes And Mid(line, i + 1, 1) = """" Then
                field = field & """"
                i = i + 1
            Else
                inQuotes = Not inQuotes
            End If
        ElseIf c = "," And Not inQuotes Then
            ReDim Preserve result(fieldIndex)
            result(fieldIndex) = field
            field = ""
            fieldIndex = fieldIndex + 1
        Else
            field = field & c
        End If
    Next
    ReDim Preserve result(fieldIndex)
    result(fieldIndex) = field
    ParseCSVLine = result
End Function

' 社員コードから名前とメールを取得（Employee DBへ接続）
' 戻り値: True=見つかった / False=見つからない
Function GetEmployeeByCode(code, ByRef outName, ByRef outEmail)
    Dim connEmp, rsTemp, sqlTemp
    outName  = ""
    outEmail = ""
    Set connEmp = GetEmployeeDBConnection()
    sqlTemp = "SELECT employee_name, email FROM IRAI.M_Employee WHERE employee_code = N'" & EscapeSQL(code) & "' AND is_active = 1"
    Set rsTemp = connEmp.Execute(sqlTemp)
    If Not rsTemp.EOF Then
        outName  = rsTemp("employee_name") & ""
        outEmail = rsTemp("email") & ""
        GetEmployeeByCode = True
    Else
        GetEmployeeByCode = False
    End If
    rsTemp.Close
    Set rsTemp = Nothing
    CloseDBConnection connEmp
End Function

' 日付をSQL用にフォーマット
Function FormatDateForSQL(dt)
    Dim d
    d = CDate(dt)
    FormatDateForSQL = Year(d) & "-" & Right("0" & Month(d), 2) & "-" & Right("0" & Day(d), 2)
End Function

Set conn = GetDBConnection()

' POST処理
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim csvData, lines, j, cols
    Dim requesterFound, assigneeFound, requestTitle, requestContent, deadlineDate
    Dim errorDetails, requesterCode, assigneeCode, lineErrors

    csvData = Trim(Request.Form("csv_data") & "")
    errorDetails = ""

    If csvData = "" Then
        errorMsg = "CSVデータを入力してください。"
    Else
        lines = Split(Replace(csvData, vbCrLf, vbLf), vbLf)

        For j = 1 To UBound(lines)
            Dim line
            line = Trim(lines(j))

            If line <> "" Then
                cols = ParseCSVLine(line)
                lineErrors = ""

                If UBound(cols) >= 3 Then
                    Dim clientCode, projectCode, importance, folderAddress
                    requesterCode = Trim(cols(0))
                    assigneeCode = Trim(cols(1))
                    requestTitle = Trim(cols(2))
                    deadlineDate = Trim(cols(3))
                    requestContent = ""
                    clientCode = ""
                    projectCode = ""
                    importance = "B"
                    folderAddress = ""
                    If UBound(cols) >= 4 Then requestContent = Trim(cols(4))
                    If UBound(cols) >= 5 Then clientCode = Trim(cols(5))
                    If UBound(cols) >= 6 Then projectCode = Trim(cols(6))
                    If UBound(cols) >= 7 Then importance = Trim(cols(7))
                    If UBound(cols) >= 8 Then folderAddress = Trim(cols(8))
                    ' 重要度の不正値はB（中）に補正
                    If importance <> "A" And importance <> "B" And importance <> "C" Then importance = "B"

                    Dim requesterName, assigneeName, requesterEmail, assigneeEmail
                    requesterFound = GetEmployeeByCode(requesterCode, requesterName, requesterEmail)
                    assigneeFound  = GetEmployeeByCode(assigneeCode,  assigneeName,  assigneeEmail)

                    ' 取引先コードから名称を取得（Client DBへ接続）
                    Dim clientName, projectName, clientFound, projectFound
                    clientName   = ""
                    projectName  = ""
                    clientFound  = True  ' 空欄の場合はエラーにしない
                    projectFound = True
                    Dim connClientImp
                    Set connClientImp = GetClientDBConnection()
                    If clientCode <> "" Then
                        Dim rsClientLookup
                        Set rsClientLookup = connClientImp.Execute("SELECT client_name FROM IRAI.M_Client WHERE client_code = N'" & EscapeSQL(clientCode) & "' AND is_active = 1")
                        If Not rsClientLookup.EOF Then
                            clientName = rsClientLookup("client_name") & ""
                        Else
                            clientFound = False
                        End If
                        CloseRecordset rsClientLookup
                    End If

                    ' 案件コードから名称を取得（Client DB）
                    If projectCode <> "" Then
                        Dim rsProjectLookup
                        Set rsProjectLookup = connClientImp.Execute("SELECT project_name FROM IRAI.M_Project WHERE project_code = N'" & EscapeSQL(projectCode) & "' AND is_active = 1")
                        If Not rsProjectLookup.EOF Then
                            projectName = rsProjectLookup("project_name") & ""
                        Else
                            projectFound = False
                        End If
                        CloseRecordset rsProjectLookup
                    End If
                    CloseDBConnection connClientImp

                    ' エラーチェック
                    If Not requesterFound Then
                        lineErrors = lineErrors & "依頼元コード「" & requesterCode & "」が見つかりません。"
                    End If
                    If Not assigneeFound Then
                        If lineErrors <> "" Then lineErrors = lineErrors & " "
                        lineErrors = lineErrors & "依頼先コード「" & assigneeCode & "」が見つかりません。"
                    End If
                    If requestTitle = "" Then
                        If lineErrors <> "" Then lineErrors = lineErrors & " "
                        lineErrors = lineErrors & "依頼件名が空です。"
                    End If
                    If Not IsDate(deadlineDate) Then
                        If lineErrors <> "" Then lineErrors = lineErrors & " "
                        lineErrors = lineErrors & "期限日「" & deadlineDate & "」が無効です。"
                    End If
                    If Not clientFound Then
                        If lineErrors <> "" Then lineErrors = lineErrors & " "
                        lineErrors = lineErrors & "取引先コード「" & clientCode & "」が見つかりません。"
                    End If
                    If Not projectFound Then
                        If lineErrors <> "" Then lineErrors = lineErrors & " "
                        lineErrors = lineErrors & "案件コード「" & projectCode & "」が見つかりません。"
                    End If

                    If lineErrors <> "" Then
                        errorCount = errorCount + 1
                        errorDetails = errorDetails & "行" & (j + 1) & ": " & lineErrors & vbCrLf
                    Else
                        sql = "INSERT INTO IRAI.T_Request (requester_code, assignee_code, requester_name, assignee_name, requester_email, assignee_email, request_title, request_content, deadline_date, importance, folder_address, client_code, client_name, project_code, project_name, status_id) " & _
                              "VALUES (N'" & EscapeSQL(requesterCode) & "', N'" & EscapeSQL(assigneeCode) & "', N'" & EscapeSQL(requesterName) & "', N'" & EscapeSQL(assigneeName) & "', N'" & EscapeSQL(requesterEmail) & "', N'" & EscapeSQL(assigneeEmail) & "', N'" & EscapeSQL(requestTitle) & "', N'" & EscapeSQL(requestContent) & "', '" & FormatDateForSQL(deadlineDate) & "', '" & EscapeSQL(importance) & "', N'" & EscapeSQL(folderAddress) & "', N'" & EscapeSQL(clientCode) & "', N'" & EscapeSQL(clientName) & "', N'" & EscapeSQL(projectCode) & "', N'" & EscapeSQL(projectName) & "', " & STATUS_NOT_STARTED & ")"

                        On Error Resume Next
                        conn.Execute sql
                        If Err.Number <> 0 Then
                            errorCount = errorCount + 1
                            errorDetails = errorDetails & "行" & (j + 1) & ": データベース登録エラー" & vbCrLf
                            Err.Clear
                        Else
                            successCount = successCount + 1
                        End If
                        On Error GoTo 0
                    End If
                Else
                    errorCount = errorCount + 1
                    errorDetails = errorDetails & "行" & (j + 1) & ": 列数が不足しています（最低4列必要）" & vbCrLf
                End If
            End If
        Next

        If successCount > 0 And errorCount = 0 Then
            SetMessage "success", successCount & " 件の依頼を登録しました。"
            Response.Redirect "request_list.asp"
            Response.End
        ElseIf successCount > 0 And errorCount > 0 Then
            SetMessage "success", successCount & " 件の依頼を登録しました。（" & errorCount & " 件はエラー）"
            errorMsg = errorDetails
        ElseIf errorCount > 0 Then
            errorMsg = "インポートに失敗しました。" & vbCrLf & errorDetails
        End If
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
                    <li><a href="export.asp">エクスポート</a></li>
                    <li><a href="import.asp">インポート</a></li>
                </ul>
            </nav>
        </div>
    </header>
    <main class="main-content">
        <div class="container">
            <%= GetMessage() %>

<h2 class="page-title">インポート</h2>

<% If errorMsg <> "" Then %>
<div class="message message-error" style="white-space: pre-line;"><%= HtmlEncode(errorMsg) %></div>
<% End If %>

<div class="card">
    <h3 class="card-title">CSVインポート</h3>
    <form method="post" action="import.asp">
        <div class="form-group">
            <label>CSVファイル選択</label>
            <div style="display: flex; gap: 10px; align-items: center;">
                <input type="file" id="csv_file" class="form-control" accept=".csv,.txt" style="flex: 1;">
                <select id="csv_encoding" class="form-control" style="width: 120px;">
                    <option value="Shift_JIS">Shift-JIS</option>
                    <option value="UTF-8">UTF-8</option>
                </select>
            </div>
            <small style="color: #666;">※ファイルを選択すると下のテキストエリアに内容が読み込まれます（文字化けする場合は文字コードを変更してください）</small>
        </div>
        <div class="form-group">
            <label>CSVデータ<span class="required">*</span></label>
            <textarea name="csv_data" id="csv_data" class="form-control" rows="10" placeholder="依頼元コード,依頼先コード,依頼件名,期限日,依頼内容,取引先コード,案件コード,重要度,フォルダアドレス"><%= HtmlEncode(Request.Form("csv_data") & "") %></textarea>
        </div>
        <div class="btn-group">
            <button type="submit" class="btn btn-primary" onclick="return confirm('インポートを実行しますか？');">インポート実行</button>
            <button type="button" class="btn btn-secondary" onclick="clearCsvData();">クリア</button>
        </div>
    </form>
</div>

<%
CloseDBConnection conn
%>
        </div>
    </main>
    <script src="js/common.js"></script>
    <script>
    // CSVファイル読み込み処理
    function loadCsvFile() {
        var fileInput = document.getElementById('csv_file');
        var file = fileInput.files[0];
        if (!file) return;

        var encoding = document.getElementById('csv_encoding').value;
        var reader = new FileReader();
        reader.onload = function(event) {
            document.getElementById('csv_data').value = event.target.result;
        };
        reader.onerror = function() {
            alert('ファイルの読み込みに失敗しました。');
        };
        reader.readAsText(file, encoding);
    }

    // ファイル選択時
    document.getElementById('csv_file').addEventListener('change', loadCsvFile);

    // 文字コード変更時に再読み込み
    document.getElementById('csv_encoding').addEventListener('change', loadCsvFile);

    // クリアボタン処理
    function clearCsvData() {
        document.getElementById('csv_data').value = '';
        document.getElementById('csv_file').value = '';
    }
    </script>
</body>
</html>
