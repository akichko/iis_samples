<%@ Language="VBScript" %>
<%
Dim fso, inputFile, outputFile, filePathInput, filePathOutput, fileContent, deleteFilePath

' �t�@�C���p�X�̐ݒ�
filePathInput = "D:\tmp\test\IUSR\input.txt"
filePathOutput = "D:\tmp\test\IUSR\output.txt"
deleteFilePath = "D:\tmp\test\IUSR\delete.txt"

' FileSystemObject �̍쐬
Set fso = CreateObject("Scripting.FileSystemObject")

' input.txt �����݂��邩�m�F
If fso.FileExists(filePathInput) Then
    ' input.txt ��ǂݍ��݃��[�h�ŊJ��
    Set inputFile = fso.OpenTextFile(filePathInput, 1, False) ' 1 = �ǂݎ���p
    fileContent = inputFile.ReadAll() ' �t�@�C���̓��e��ǂݍ���
    inputFile.Close
    Set inputFile = Nothing
    
    ' output.txt ��ǋL���[�h�ŊJ���i8 = �ǋL���[�h�j
    Set outputFile = fso.OpenTextFile(filePathOutput, 8, True)
    outputFile.WriteLine fileContent ' �ǂݍ��񂾓��e����������
    outputFile.Close
    Set outputFile = Nothing

    Response.Write "�t�@�C���̓��e��ǋL���܂����B<br>"
Else
    Response.Write "���̓t�@�C����������܂���: " & filePathInput & "<br>"
End If

' delete.txt �����݂��邩�m�F���A���݂���ꍇ�͍폜
If fso.FileExists(deleteFilePath) Then
    fso.DeleteFile(deleteFilePath)
    Response.Write "�t�@�C�����폜���܂���: " & deleteFilePath & "<br>"
Else
    Response.Write "�폜����t�@�C����������܂���: " & deleteFilePath & "<br>"
End If

' �I�u�W�F�N�g�̉��
Set fso = Nothing
%>
