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
' 社員コードから社員情報を取得するAPI
' パラメータ:
'   code : 社員コード
' 戻り値: JSON
'   { "found": true/false, "name": "名前", "email": "..." }
' ============================================

Dim employeeCode
employeeCode = Trim(Request.QueryString("code") & "")

' 入力チェック
If employeeCode = "" Then
    Response.Write "{""found"":false,""name"":"""",""email"":""""}"
    Response.End
End If

Dim conn, rs, sql
Set conn = GetEmployeeDBConnection()

' 社員コードで検索（有効な社員のみ）
sql = "SELECT employee_name, email FROM IRAI.M_Employee " & _
      "WHERE employee_code = N'" & EscapeSQL(employeeCode) & "' AND is_active = 1"

Set rs = conn.Execute(sql)

If Not rs.EOF Then
    ' JSON文字列内のダブルクォートをエスケープ
    Dim empName, empEmail
    empName = rs("employee_name") & ""
    empName = Replace(empName, "\", "\\")
    empName = Replace(empName, """", "\""")
    empEmail = rs("email") & ""
    empEmail = Replace(empEmail, "\", "\\")
    empEmail = Replace(empEmail, """", "\""")

    Response.Write "{""found"":true,""name"":""" & empName & """,""email"":""" & empEmail & """}"
Else
    Response.Write "{""found"":false,""name"":"""",""email"":""""}"
End If

CloseRecordset rs
CloseDBConnection conn
%>
