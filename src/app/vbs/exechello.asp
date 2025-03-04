<%@ Language="VBScript" %>

<%
Dim aaa
aaa = "abc"
Response.Write "Hello 1 " & aaa & vbCrLf

On Error Resume Next

Dim objShell, strCommand, objExec
Set objShell = Server.CreateObject("WScript.Shell")

' 実行するコマンド
Response.Write("exe 1 ")
strCommand = "D:\tmp\test\hello.exe >> D:\tmp\test\log-hello.txt"

' 実行
objShell.Run strCommand, 0, True


Response.Write " Err.Number:" & Err.Number

' 結果の出力'
If Err.Number = 0 Then
    Response.Write "hello.exe を実行しました。"
Else
    Response.Write "エラー: " & Err.Description
End If

Set objShell = Nothing


Response.Write("[Hello]last")
%>
