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
' 依頼削除処理
' 論理削除（is_deleted = 1）
' ============================================

Dim conn, sql
Dim requestId

requestId = SafeInt(Request.QueryString("id"), 0)

If requestId = 0 Then
    SetMessage "error", "依頼IDが指定されていません。"
    Response.Redirect "request_list.asp"
    Response.End
End If

Set conn = GetDBConnection()

' 論理削除処理
sql = "UPDATE T_Request SET is_deleted = 1, updated_at = GETDATE() WHERE request_id = " & requestId

On Error Resume Next
conn.Execute sql
If Err.Number <> 0 Then
    SetMessage "error", "削除に失敗しました：" & Err.Description
    Err.Clear
Else
    SetMessage "success", "依頼を削除しました。"
End If
On Error GoTo 0

CloseDBConnection conn

Response.Redirect "request_list.asp"
%>
