<%@ Language="VBScript" %>

<%
Dim aaa
aaa = "abc"
Response.Write "Hello 1 " & aaa & vbCrLf

On Error Resume Next

Dim objShell, strCommand, objExec
Set objShell = Server.CreateObject("WScript.Shell")

' ���s����R�}���h
Response.Write("exe 1 ")
strCommand = "D:\tmp\test\hello.exe >> D:\tmp\test\log-hello.txt"

' ���s
objShell.Run strCommand, 0, True


Response.Write " Err.Number:" & Err.Number

' ���ʂ̏o��'
If Err.Number = 0 Then
    Response.Write "hello.exe �����s���܂����B"
Else
    Response.Write "�G���[: " & Err.Description
End If

Set objShell = Nothing


Response.Write("[Hello]last")
%>
