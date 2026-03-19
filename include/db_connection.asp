<%
' ============================================
' データベース接続モジュール
' ============================================

' データベース接続オブジェクトを取得
Function GetDBConnection()
    Dim conn
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open DB_CONNECTION_STRING
    Set GetDBConnection = conn
End Function

' 取引先・案件DBへの接続オブジェクトを取得（Client DB）
Function GetClientDBConnection()
    Dim conn
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open CLIENT_DB_CONNECTION_STRING
    Set GetClientDBConnection = conn
End Function

' 社員DBへの接続オブジェクトを取得（Employee DB）
Function GetEmployeeDBConnection()
    Dim conn
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open EMPLOYEE_DB_CONNECTION_STRING
    Set GetEmployeeDBConnection = conn
End Function

' データベース接続を閉じる
Sub CloseDBConnection(conn)
    If Not conn Is Nothing Then
        If conn.State = 1 Then
            conn.Close
        End If
        Set conn = Nothing
    End If
End Sub

' レコードセットを閉じる
Sub CloseRecordset(rs)
    If Not rs Is Nothing Then
        If rs.State = 1 Then
            rs.Close
        End If
        Set rs = Nothing
    End If
End Sub

' SQLインジェクション対策：文字列エスケープ
Function EscapeSQL(str)
    If IsNull(str) Or str = "" Then
        EscapeSQL = ""
    Else
        EscapeSQL = Replace(str, "'", "''")
    End If
End Function

' NULL安全な値取得
Function SafeValue(val, defaultVal)
    If IsNull(val) Or val = "" Then
        SafeValue = defaultVal
    Else
        SafeValue = val
    End If
End Function

' 数値安全変換
Function SafeInt(val, defaultVal)
    If IsNumeric(val) Then
        SafeInt = CInt(val)
    Else
        SafeInt = defaultVal
    End If
End Function

' 日付安全変換
Function SafeDate(val, defaultVal)
    If IsDate(val) Then
        SafeDate = CDate(val)
    Else
        SafeDate = defaultVal
    End If
End Function
%>
