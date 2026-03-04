<%@ Language="VBScript" CodePage=65001 %>
<%
' ============================================
' リマインドメール送信バッチ
' タスクスケジューラで定期実行する
'
' 使用方法:
' cscript //nologo send_reminder.asp
' または
' curl http://localhost/iraikanri/batch/send_reminder.asp
' ============================================

Response.ContentType = "text/plain"
Response.Charset = "UTF-8"
%>
<!--#include file="../include/config.asp"-->
<!--#include file="../include/db_connection.asp"-->
<!--#include file="../include/functions.asp"-->
<!--#include file="../include/mail.asp"-->
<%
Dim conn, rs, sql
Dim sentCount, errorCount

sentCount = 0
errorCount = 0

Response.Write "========================================" & vbCrLf
Response.Write " リマインドメール送信バッチ" & vbCrLf
Response.Write " 実行日時: " & Now() & vbCrLf
Response.Write "========================================" & vbCrLf & vbCrLf

Set conn = GetDBConnection()

' --------------------------------------------
' 1. 期限間近の依頼（リマインドメール）
' --------------------------------------------
Response.Write "[期限間近の依頼]" & vbCrLf

sql = "SELECT r.request_id, r.request_title, r.deadline_date, r.request_content, " & _
      "req.employee_name AS requester_name, ass.employee_name AS assignee_name, ass.email AS assignee_email " & _
      "FROM IRAI.T_Request r " & _
      "INNER JOIN IRAI.M_Employee req ON r.requester_id = req.employee_id " & _
      "INNER JOIN IRAI.M_Employee ass ON r.assignee_id = ass.employee_id " & _
      "WHERE r.is_deleted = 0 " & _
      "AND r.status_id NOT IN (" & STATUS_COMPLETED & "," & STATUS_NOT_APPLICABLE & ") " & _
      "AND r.deadline_date = DATEADD(day, " & REMIND_DAYS_BEFORE & ", CONVERT(DATE, GETDATE())) " & _
      "AND ass.email IS NOT NULL AND ass.email <> ''"
Set rs = conn.Execute(sql)

If rs.EOF Then
    Response.Write "  対象なし" & vbCrLf
Else
    Do While Not rs.EOF
        Dim toEmail1, subject1, body1, result1

        toEmail1 = rs("assignee_email")
        subject1 = "【期限間近】" & rs("request_title")

        body1 = "以下の依頼の期限が" & REMIND_DAYS_BEFORE & "日後に迫っています。" & vbCrLf & vbCrLf
        body1 = body1 & "依頼ID: " & rs("request_id") & vbCrLf
        body1 = body1 & "依頼件名: " & rs("request_title") & vbCrLf
        body1 = body1 & "依頼元: " & rs("requester_name") & vbCrLf
        body1 = body1 & "期限日: " & FormatDateJP(rs("deadline_date")) & vbCrLf
        body1 = body1 & vbCrLf
        body1 = body1 & "対応をお願いいたします。" & vbCrLf
        body1 = body1 & vbCrLf
        body1 = body1 & "----" & vbCrLf
        body1 = body1 & APP_TITLE

        result1 = SendMail(toEmail1, subject1, body1)
        LogMailSend conn, rs("request_id"), MAIL_TYPE_REMIND, toEmail1, subject1, body1, result1, ""

        If result1 Then
            sentCount = sentCount + 1
            Response.Write "  送信成功: ID=" & rs("request_id") & " -> " & toEmail1 & vbCrLf
        Else
            errorCount = errorCount + 1
            Response.Write "  送信失敗: ID=" & rs("request_id") & " -> " & toEmail1 & vbCrLf
        End If

        rs.MoveNext
    Loop
End If
CloseRecordset rs

Response.Write vbCrLf

' --------------------------------------------
' 2. 期限切れの依頼（期限切れ通知メール）
' --------------------------------------------
Response.Write "[期限切れの依頼]" & vbCrLf

sql = "SELECT r.request_id, r.request_title, r.deadline_date, r.request_content, " & _
      "req.employee_name AS requester_name, ass.employee_name AS assignee_name, ass.email AS assignee_email " & _
      "FROM IRAI.T_Request r " & _
      "INNER JOIN IRAI.M_Employee req ON r.requester_id = req.employee_id " & _
      "INNER JOIN IRAI.M_Employee ass ON r.assignee_id = ass.employee_id " & _
      "WHERE r.is_deleted = 0 " & _
      "AND r.status_id NOT IN (" & STATUS_COMPLETED & "," & STATUS_NOT_APPLICABLE & ") " & _
      "AND r.deadline_date = DATEADD(day, -1, CONVERT(DATE, GETDATE())) " & _
      "AND ass.email IS NOT NULL AND ass.email <> ''"
Set rs = conn.Execute(sql)

If rs.EOF Then
    Response.Write "  対象なし" & vbCrLf
Else
    Do While Not rs.EOF
        Dim toEmail2, subject2, body2, result2

        toEmail2 = rs("assignee_email")
        subject2 = "【期限超過】" & rs("request_title")

        body2 = "以下の依頼の期限が過ぎています。" & vbCrLf & vbCrLf
        body2 = body2 & "依頼ID: " & rs("request_id") & vbCrLf
        body2 = body2 & "依頼件名: " & rs("request_title") & vbCrLf
        body2 = body2 & "依頼元: " & rs("requester_name") & vbCrLf
        body2 = body2 & "期限日: " & FormatDateJP(rs("deadline_date")) & vbCrLf
        body2 = body2 & vbCrLf
        body2 = body2 & "至急対応をお願いいたします。" & vbCrLf
        body2 = body2 & vbCrLf
        body2 = body2 & "----" & vbCrLf
        body2 = body2 & APP_TITLE

        result2 = SendMail(toEmail2, subject2, body2)
        LogMailSend conn, rs("request_id"), MAIL_TYPE_OVERDUE, toEmail2, subject2, body2, result2, ""

        If result2 Then
            sentCount = sentCount + 1
            Response.Write "  送信成功: ID=" & rs("request_id") & " -> " & toEmail2 & vbCrLf
        Else
            errorCount = errorCount + 1
            Response.Write "  送信失敗: ID=" & rs("request_id") & " -> " & toEmail2 & vbCrLf
        End If

        rs.MoveNext
    Loop
End If
CloseRecordset rs

Response.Write vbCrLf
Response.Write "========================================" & vbCrLf
Response.Write " 処理完了" & vbCrLf
Response.Write " 送信成功: " & sentCount & " 件" & vbCrLf
Response.Write " 送信失敗: " & errorCount & " 件" & vbCrLf
Response.Write "========================================" & vbCrLf

CloseDBConnection conn
%>
