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
'   { "found": true/false, "id": 数値, "name": "名前" }
' ============================================

Dim employeeCode
employeeCode = Trim(Request.QueryString("code") & "")

' 入力チェック
If employeeCode = "" Then
    Response.Write "{""found"":false,""id"":0,""name"":""""}"
    Response.End
End If

Dim conn, rs, sql
Set conn = GetDBConnection()

' 社員コードで検索（有効な社員のみ）
sql = "SELECT employee_id, employee_name FROM IRAI.M_Employee " & _
      "WHERE employee_code = N'" & EscapeSQL(employeeCode) & "' AND is_active = 1"

Set rs = conn.Execute(sql)

If Not rs.EOF Then
    ' JSON文字列内のダブルクォートをエスケープ
    Dim empName
    empName = rs("employee_name") & ""
    empName = Replace(empName, "\", "\\")
    empName = Replace(empName, """", "\""")

    Response.Write "{""found"":true,""id"":" & rs("employee_id") & ",""name"":""" & empName & """}"
Else
    Response.Write "{""found"":false,""id"":0,""name"":""""}"
End If

CloseRecordset rs
CloseDBConnection conn
%>
