<%@ Language="VBScript" %>

<%
Response.Write "[START]"

On Error Resume Next

Dim objShell, strCommand, objExec
Set objShell = Server.CreateObject("WScript.Shell")

' 実行するコマンド
strCommand = "cmd.exe /c whoami >> D:\tmp\test\whoami.txt && whoami /groups >> D:\tmp\test\whoami.txt"

' 実行
objShell.Run strCommand, 0, True

Response.Write "[Err.Number:" & Err.Number & "]"

' 結果の出力'
If Err.Number = 0 Then
    Response.Write "whoami コマンドを実行しました。"
Else
    Response.Write "エラー: " & Err.Description
End If


Set objShell = Nothing

Response.Write "[End]"
%>
