Sub readFileContents(ByVal fullFilename As String, ByRef Return_str() As String)

    Dim objFSO As Object
    Dim objTF As Object
    Dim strIn As String

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objTF = objFSO.OpenTextFile(fullFilename, 1)
    strIn = objTF.readall
    objTF.Close
    Return_str = Split(strIn, vbCrLf)
        
End Sub
