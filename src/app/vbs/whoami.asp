<%@ Language="VBScript" %>

<%
Response.Write "[START]"

On Error Resume Next

Dim objShell, strCommand, objExec
Set objShell = Server.CreateObject("WScript.Shell")

' ���s����R�}���h
strCommand = "cmd.exe /c whoami >> D:\tmp\test\whoami.txt && whoami /groups >> D:\tmp\test\whoami.txt"

' ���s
objShell.Run strCommand, 0, True

Response.Write "[Err.Number:" & Err.Number & "]"

' ���ʂ̏o��'
If Err.Number = 0 Then
    Response.Write "whoami �R�}���h�����s���܂����B"
Else
    Response.Write "�G���[: " & Err.Description
End If


Set objShell = Nothing

Response.Write "[End]"
%>
