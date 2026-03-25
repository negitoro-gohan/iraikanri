<%@ Language="VBScript" CodePage=65001 %>
<%
Response.CodePage = 65001
Session.CodePage = 65001
Response.Charset = "UTF-8"
Response.ContentType = "application/json; charset=UTF-8"
%>
<!--#include file="include/config.asp"-->
<!--#include file="include/db_connection.asp"-->
<!--#include file="include/functions.asp"-->
<%
' ============================================
' 社員検索API（AJAX用）
' パラメータ:
'   keyword : 社員コード or 社員名で部分一致検索
' 戻り値: JSON配列
'   [{"employee_code":"EMP001", "employee_name":"山田太郎", "email":"..."}]
' ============================================

Dim keyword
keyword = Trim(Request.QueryString("keyword") & "")

Dim conn, rs, sql
Set conn = GetEmployeeDBConnection()

' 社員コードまたは社員名で部分一致検索（有効な社員のみ、最大50件）
' キーワードが空の場合は全件返す
If keyword = "" Then
    sql = "SELECT TOP 50 employee_code, employee_name, email " & _
          "FROM IRAI.M_Employee " & _
          "WHERE is_active = 1 " & _
          "ORDER BY employee_code"
Else
    sql = "SELECT TOP 50 employee_code, employee_name, email " & _
          "FROM IRAI.M_Employee " & _
          "WHERE is_active = 1 " & _
          "AND (employee_code LIKE N'%" & EscapeSQL(keyword) & "%' " & _
          "  OR employee_name LIKE N'%" & EscapeSQL(keyword) & "%') " & _
          "ORDER BY employee_code"
End If

Set rs = conn.Execute(sql)

' JSON配列を構築
Dim jsonResult, isFirst
jsonResult = "["
isFirst = True

Do While Not rs.EOF
    If Not isFirst Then jsonResult = jsonResult & ","
    isFirst = False

    ' 各フィールドのエスケープ処理
    Dim empCode, empName, empEmail
    empCode = rs("employee_code") & ""
    empCode = Replace(empCode, "\", "\\")
    empCode = Replace(empCode, """", "\""")

    empName = rs("employee_name") & ""
    empName = Replace(empName, "\", "\\")
    empName = Replace(empName, """", "\""")

    empEmail = rs("email") & ""
    empEmail = Replace(empEmail, "\", "\\")
    empEmail = Replace(empEmail, """", "\""")

    jsonResult = jsonResult & "{" & _
        """employee_code"":""" & empCode & """," & _
        """employee_name"":""" & empName & """," & _
        """email"":""" & empEmail & """" & _
        "}"

    rs.MoveNext
Loop

jsonResult = jsonResult & "]"

Response.Write jsonResult

CloseRecordset rs
CloseDBConnection conn
%>
