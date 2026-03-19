<%@ Language="VBScript" CodePage=65001 %>
<%
Response.CodePage = 65001
Response.Charset = "UTF-8"
Response.ContentType = "text/html; charset=UTF-8"

' ASPError オブジェクトからエラー情報を取得
Dim aspErr
Set aspErr = Server.GetLastError()
%>
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>エラー - 依頼事項 期限管理システム</title>
    <link rel="stylesheet" href="/iraikanri/css/style.css">
    <style>
        .error-detail { background:#fff3f3; border:1px solid #f5c6cb; border-radius:4px; padding:20px; margin-top:15px; }
        .error-detail table { width:100%; border-collapse:collapse; }
        .error-detail th { text-align:left; padding:6px 12px; width:120px; color:#666; font-weight:normal; }
        .error-detail td { padding:6px 12px; font-family:monospace; word-break:break-all; }
        .error-detail tr + tr { border-top:1px solid #f0d0d0; }
    </style>
</head>
<body>
    <header class="site-header">
        <div class="header-container">
            <h1 class="site-title"><a href="/iraikanri/index.asp">依頼事項 期限管理システム</a></h1>
        </div>
    </header>
    <main class="main-content">
        <div class="container">
            <h2 class="page-title">エラーが発生しました</h2>
            <div class="message message-error">
                サーバーでエラーが発生しました。操作を中断してください。
            </div>

            <% If Not (aspErr Is Nothing) Then %>
            <div class="card">
                <h3 class="card-title">エラー詳細</h3>
                <div class="error-detail">
                    <table>
                        <tr>
                            <th>ファイル</th>
                            <td><%= Server.HTMLEncode(aspErr.File & "") %></td>
                        </tr>
                        <tr>
                            <th>行番号</th>
                            <td><%= aspErr.Line %></td>
                        </tr>
                        <tr>
                            <th>エラー番号</th>
                            <td><%= aspErr.Number %> (<%= aspErr.ASPCode & "" %>)</td>
                        </tr>
                        <tr>
                            <th>説明</th>
                            <td><%= Server.HTMLEncode(aspErr.Description & "") %></td>
                        </tr>
                        <% If Trim(aspErr.Source & "") <> "" Then %>
                        <tr>
                            <th>ソース</th>
                            <td><pre style="margin:0;"><%= Server.HTMLEncode(aspErr.Source & "") %></pre></td>
                        </tr>
                        <% End If %>
                    </table>
                </div>
            </div>
            <% End If %>

            <div style="margin-top:20px;">
                <a href="/iraikanri/index.asp" class="btn btn-primary">ホームに戻る</a>
            </div>
        </div>
    </main>
</body>
</html>
