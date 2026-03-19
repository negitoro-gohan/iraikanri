<%
' ============================================
' 共通関数モジュール
' ============================================

' HTMLエスケープ（XSS対策）
Function HtmlEncode(str)
    If IsNull(str) Or str = "" Then
        HtmlEncode = ""
    Else
        HtmlEncode = Server.HTMLEncode(str)
    End If
End Function

' 日付フォーマット（YYYY/MM/DD形式）
Function FormatDateJP(dt)
    If IsNull(dt) Or dt = "" Then
        FormatDateJP = ""
    Else
        FormatDateJP = Year(dt) & "/" & Right("0" & Month(dt), 2) & "/" & Right("0" & Day(dt), 2)
    End If
End Function

' 日付フォーマット（YYYY-MM-DD形式、HTML5 input用）
Function FormatDateISO(dt)
    If IsNull(dt) Or dt = "" Then
        FormatDateISO = ""
    Else
        FormatDateISO = Year(dt) & "-" & Right("0" & Month(dt), 2) & "-" & Right("0" & Day(dt), 2)
    End If
End Function

' 期限状態を判定（-1:期限切れ、0:期限間近、1:通常）
Function GetDeadlineStatus(deadlineDate)
    Dim today, daysRemaining
    today = Date()

    If IsNull(deadlineDate) Or deadlineDate = "" Then
        GetDeadlineStatus = 1
        Exit Function
    End If

    daysRemaining = DateDiff("d", today, deadlineDate)

    If daysRemaining < 0 Then
        GetDeadlineStatus = -1  ' 期限切れ
    ElseIf daysRemaining <= REMIND_DAYS_BEFORE Then
        GetDeadlineStatus = 0   ' 期限間近
    Else
        GetDeadlineStatus = 1   ' 通常
    End If
End Function

' 期限状態に応じたCSSクラスを返す
Function GetDeadlineClass(deadlineDate, statusId)
    ' 完了または対象外の場合は通常表示
    If statusId = STATUS_COMPLETED Or statusId = STATUS_NOT_APPLICABLE Then
        GetDeadlineClass = ""
        Exit Function
    End If

    Dim status
    status = GetDeadlineStatus(deadlineDate)

    Select Case status
        Case -1
            GetDeadlineClass = "deadline-overdue"
        Case 0
            GetDeadlineClass = "deadline-soon"
        Case Else
            GetDeadlineClass = ""
    End Select
End Function

' ステータス名を取得
Function GetStatusName(statusId)
    Select Case statusId
        Case STATUS_NOT_STARTED
            GetStatusName = "未着手"
        Case STATUS_IN_PROGRESS
            GetStatusName = "着手済"
        Case STATUS_COMPLETED
            GetStatusName = "対応終了"
        Case STATUS_NOT_APPLICABLE
            GetStatusName = "対象外"
        Case Else
            GetStatusName = "不明"
    End Select
End Function

' ステータスに応じたCSSクラスを返す
Function GetStatusClass(statusId)
    Select Case statusId
        Case STATUS_NOT_STARTED
            GetStatusClass = "status-not-started"
        Case STATUS_IN_PROGRESS
            GetStatusClass = "status-in-progress"
        Case STATUS_COMPLETED
            GetStatusClass = "status-completed"
        Case STATUS_NOT_APPLICABLE
            GetStatusClass = "status-not-applicable"
        Case Else
            GetStatusClass = ""
    End Select
End Function

' 改行をBRタグに変換
Function NL2BR(str)
    If IsNull(str) Or str = "" Then
        NL2BR = ""
    Else
        NL2BR = Replace(Replace(str, vbCrLf, "<br>"), vbLf, "<br>")
    End If
End Function

' 現在のログインユーザー名を取得
' DOMAIN\username 形式の場合、DOMAIN\ 部分を除いたユーザー名のみ返す
' DEBUG_USER が設定されている場合はそちらを返す（テスト用なりすまし）
Function GetCurrentUser()
    Dim u
    If DEBUG_USER <> "" Then
        u = DEBUG_USER
    Else
        u = Request.ServerVariables("LOGON_USER")
        If u = "" Then u = Request.ServerVariables("AUTH_USER")
    End If
    ' DOMAIN\ プレフィックスを除去
    Dim pos
    pos = InStr(u, "\")
    If pos > 0 Then u = Mid(u, pos + 1)
    GetCurrentUser = u
End Function

' メッセージ表示用のセッションに保存
Sub SetMessage(msgType, msgText)
    Session("MSG_TYPE") = msgType
    Session("MSG_TEXT") = msgText
End Sub

' メッセージを取得して削除
Function GetMessage()
    Dim msg
    msg = ""

    If Session("MSG_TEXT") <> "" Then
        Dim msgType, msgClass
        msgType = Session("MSG_TYPE")

        Select Case msgType
            Case "success"
                msgClass = "message-success"
            Case "error"
                msgClass = "message-error"
            Case "warning"
                msgClass = "message-warning"
            Case Else
                msgClass = "message-info"
        End Select

        msg = "<div class=""message " & msgClass & """>" & HtmlEncode(Session("MSG_TEXT")) & "</div>"

        Session("MSG_TYPE") = ""
        Session("MSG_TEXT") = ""
    End If

    GetMessage = msg
End Function
%>
