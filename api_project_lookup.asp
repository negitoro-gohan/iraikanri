<%@ Language="VBScript" CodePage=65001 %>
<%
Response.CodePage = 65001
Response.Charset = "UTF-8"
Response.ContentType = "application/json; charset=UTF-8"
%>
<!--#include file="include/config.asp"-->
<!--#include file="include/db_connection.asp"-->
<!--#include file="include/functions.asp"-->
<%
' ============================================
' 取引先・案件検索API
' ============================================

Dim conn, rs, sql
Dim action, clientCode

action     = Trim(Request.QueryString("action") & "")
clientCode = Trim(Request.QueryString("client_code") & "")

Set conn = GetClientDBConnection()

If action = "clients" Then
    ' 取引先一覧を返す（code / name のみ）
    sql = "SELECT client_code, client_name FROM IRAI.M_Client WHERE is_active = 1 ORDER BY client_code"
    Set rs = conn.Execute(sql)

    Response.Write "["
    Dim first
    first = True
    Do While Not rs.EOF
        If Not first Then Response.Write ","
        first = False
        Response.Write "{""code"":""" & EscapeJSON(rs("client_code") & "") & ""","
        Response.Write """name"":""" & EscapeJSON(rs("client_name") & "") & """}"
        rs.MoveNext
    Loop
    Response.Write "]"
    CloseRecordset rs

ElseIf action = "projects" And clientCode <> "" Then
    ' 指定した取引先コードの案件一覧を返す（code / name のみ）
    sql = "SELECT project_code, project_name FROM IRAI.M_Project WHERE client_code = N'" & EscapeSQL(clientCode) & "' AND is_active = 1 ORDER BY project_code"
    Set rs = conn.Execute(sql)

    Response.Write "["
    first = True
    Do While Not rs.EOF
        If Not first Then Response.Write ","
        first = False
        Response.Write "{""code"":""" & EscapeJSON(rs("project_code") & "") & ""","
        Response.Write """name"":""" & EscapeJSON(rs("project_name") & "") & """}"
        rs.MoveNext
    Loop
    Response.Write "]"
    CloseRecordset rs

ElseIf action = "lookup_client" Then
    ' 取引先コードから名称を返す（コード入力のonblur用）
    Dim lookupClientCode
    lookupClientCode = Trim(Request.QueryString("code") & "")
    If lookupClientCode = "" Then
        Response.Write "{""found"":false}"
    Else
        sql = "SELECT client_name FROM IRAI.M_Client WHERE client_code = N'" & EscapeSQL(lookupClientCode) & "' AND is_active = 1"
        Set rs = conn.Execute(sql)
        If rs.EOF Then
            Response.Write "{""found"":false}"
        Else
            Response.Write "{""found"":true,""name"":""" & EscapeJSON(rs("client_name") & "") & """}"
        End If
        CloseRecordset rs
    End If

ElseIf action = "lookup_project" Then
    ' 案件コードから名称を返す（コード入力のonblur用）
    Dim lookupProjectCode
    lookupProjectCode = Trim(Request.QueryString("code") & "")
    If lookupProjectCode = "" Then
        Response.Write "{""found"":false}"
    Else
        sql = "SELECT project_name FROM IRAI.M_Project WHERE project_code = N'" & EscapeSQL(lookupProjectCode) & "' AND is_active = 1"
        Set rs = conn.Execute(sql)
        If rs.EOF Then
            Response.Write "{""found"":false}"
        Else
            Response.Write "{""found"":true,""name"":""" & EscapeJSON(rs("project_name") & "") & """}"
        End If
        CloseRecordset rs
    End If

Else
    Response.Write "{""error"":""Invalid action or parameters""}"
End If

CloseDBConnection conn

' JSONエスケープ関数
Function EscapeJSON(str)
    Dim result
    result = str
    result = Replace(result, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCr, "")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbTab, "\t")
    EscapeJSON = result
End Function
%>
