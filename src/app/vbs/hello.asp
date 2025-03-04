<%@ Language="VBScript" %>
<%
Response.CodePage = 65001  ' UTF-8を使用
Response.Charset = "UTF-8" ' クライアントにUTF-8を指示
%>
<!DOCTYPE html>
<html>
<head>
    <title>Hello World in ASP</title>
</head>
<body>
    <h1>Hello, World!</h1>
    <p>This is a simple ASP page.</p>
    <p>日本語テスト</p>
    <%
        ' VBScriptでメッセージを表示
        Response.Write("Hello from ASP!")
    %>
</body>
</html>
