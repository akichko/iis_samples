<%@ Language="VBScript" %>
<%
Dim fso, inputFile, outputFile, filePathInput, filePathOutput, fileContent, deleteFilePath

' ファイルパスの設定
filePathInput = "D:\tmp\test\IUSR\input.txt"
filePathOutput = "D:\tmp\test\IUSR\output.txt"
deleteFilePath = "D:\tmp\test\IUSR\delete.txt"

' FileSystemObject の作成
Set fso = CreateObject("Scripting.FileSystemObject")

' input.txt が存在するか確認
If fso.FileExists(filePathInput) Then
    ' input.txt を読み込みモードで開く
    Set inputFile = fso.OpenTextFile(filePathInput, 1, False) ' 1 = 読み取り専用
    fileContent = inputFile.ReadAll() ' ファイルの内容を読み込む
    inputFile.Close
    Set inputFile = Nothing
    
    ' output.txt を追記モードで開く（8 = 追記モード）
    Set outputFile = fso.OpenTextFile(filePathOutput, 8, True)
    outputFile.WriteLine fileContent ' 読み込んだ内容を書き込む
    outputFile.Close
    Set outputFile = Nothing

    Response.Write "ファイルの内容を追記しました。<br>"
Else
    Response.Write "入力ファイルが見つかりません: " & filePathInput & "<br>"
End If

' delete.txt が存在するか確認し、存在する場合は削除
If fso.FileExists(deleteFilePath) Then
    fso.DeleteFile(deleteFilePath)
    Response.Write "ファイルを削除しました: " & deleteFilePath & "<br>"
Else
    Response.Write "削除するファイルが見つかりません: " & deleteFilePath & "<br>"
End If

' オブジェクトの解放
Set fso = Nothing
%>
