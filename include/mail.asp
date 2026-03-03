<%
' ============================================
' メール送信モジュール
' CDO (Collaboration Data Objects) を使用
' ============================================

' メール送信関数
' 戻り値: True=成功, False=失敗
Function SendMail(toAddress, subject, body)
    Dim objMail
    Dim result

    result = False

    On Error Resume Next

    Set objMail = Server.CreateObject("CDO.Message")

    ' SMTP設定
    With objMail.Configuration.Fields
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2  ' cdoSendUsingPort
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = MAIL_SMTP_SERVER
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = MAIL_SMTP_PORT
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
        ' 認証が必要な場合は以下をアンコメント
        '.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        '.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "username"
        '.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "password"
        .Update
    End With

    ' メール設定
    With objMail
        .From = MAIL_FROM_NAME & " <" & MAIL_FROM_ADDRESS & ">"
        .To = toAddress
        .Subject = subject
        .TextBody = body
        .Send
    End With

    If Err.Number = 0 Then
        result = True
    End If

    Set objMail = Nothing

    On Error GoTo 0

    SendMail = result
End Function

' メール送信ログを記録
Sub LogMailSend(conn, requestId, mailType, toEmail, subject, body, isSuccess, errorMessage)
    Dim sql
    sql = "INSERT INTO T_MailLog (request_id, mail_type, to_email, subject, body, is_success, error_message) " & _
          "VALUES (" & requestId & ", " & mailType & ", N'" & EscapeSQL(toEmail) & "', N'" & EscapeSQL(subject) & "', N'" & EscapeSQL(body) & "', " & IIf(isSuccess, 1, 0) & ", N'" & EscapeSQL(errorMessage) & "')"

    On Error Resume Next
    conn.Execute sql
    On Error GoTo 0
End Sub

' 新規依頼通知メール送信
Sub SendNewRequestMail(conn, requestId)
    Dim rs, sql
    Dim toEmail, subject, body

    sql = "SELECT r.request_id, r.request_title, r.deadline_date, r.request_content, " & _
          "req.employee_name AS requester_name, ass.employee_name AS assignee_name, ass.email AS assignee_email " & _
          "FROM T_Request r " & _
          "INNER JOIN M_Employee req ON r.requester_id = req.employee_id " & _
          "INNER JOIN M_Employee ass ON r.assignee_id = ass.employee_id " & _
          "WHERE r.request_id = " & requestId
    Set rs = conn.Execute(sql)

    If Not rs.EOF Then
        toEmail = rs("assignee_email")

        If Not IsNull(toEmail) And toEmail <> "" Then
            subject = "【新規依頼】" & rs("request_title")

            body = "新しい依頼が登録されました。" & vbCrLf & vbCrLf
            body = body & "依頼ID: " & rs("request_id") & vbCrLf
            body = body & "依頼件名: " & rs("request_title") & vbCrLf
            body = body & "依頼元: " & rs("requester_name") & vbCrLf
            body = body & "依頼先: " & rs("assignee_name") & vbCrLf
            body = body & "期限日: " & FormatDateJP(rs("deadline_date")) & vbCrLf
            body = body & vbCrLf
            body = body & "【依頼内容】" & vbCrLf
            body = body & SafeValue(rs("request_content"), "（なし）") & vbCrLf
            body = body & vbCrLf
            body = body & "----" & vbCrLf
            body = body & APP_TITLE

            Dim result
            result = SendMail(toEmail, subject, body)
            LogMailSend conn, requestId, MAIL_TYPE_NEW, toEmail, subject, body, result, ""
        End If
    End If

    CloseRecordset rs
End Sub

' IIf関数（ここでも定義）
Function IIf(condition, trueValue, falseValue)
    If condition Then
        IIf = trueValue
    Else
        IIf = falseValue
    End If
End Function
%>
