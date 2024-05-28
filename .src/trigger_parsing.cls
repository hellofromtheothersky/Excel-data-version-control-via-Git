Dim parentPath As String
Function GetPythonPath() As String
    Dim oShell As Object
    Dim PythonPath As String
    Dim PythonPathLines As Variant
    
    Set oShell = CreateObject("WScript.Shell")
    
    ' Use the 'where' command to find the Python executable
    PythonPath = Trim(oShell.Exec("where python").StdOut.ReadAll())
    
    If PythonPath = "" Then
        ' Python not found
        GetPythonPath = ""
    Else
        ' Check if there are multiple paths returned
        PythonPathLines = Split(PythonPath, vbNewLine)
        If UBound(PythonPathLines) > 0 Then
            ' Take the first path if there are multiple
            GetPythonPath = Trim(PythonPathLines(0))
        Else
            ' Return the single path
            GetPythonPath = PythonPath
        End If
    End If
End Function
Sub FindGitParentFolder()
  Dim currentPath As String

  ' Get the current workbook path
  currentPath = ThisWorkbook.Path

  ' Loop until a folder with .git is found or the root directory is reached
  Do While Not Dir(currentPath & "\excel_as_text", vbDirectory) <> ""
    ' Get the parent folder path
    parentPath = Left(currentPath, InStrRev(currentPath, "\") - 1)
    ' Update current path to the parent folder
    currentPath = parentPath
  Loop

  ' Check if a folder with .git was found
  If Dir(currentPath & "\excel_as_text", vbDirectory) <> "" Then
    ' Replace ThisWorkbook.Path with the parent folder containing the .git directory
    parentPath = currentPath
  Else
    MsgBox "No excel_as_text folder found in parent directories", vbExclamation
  End If

End Sub
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    FindGitParentFolder
    Dim objShell As Object
    Dim pythonExePath, PythonScriptPath As String
    Set objShell = VBA.CreateObject("Wscript.Shell")
    
    ' Specify the file path
    filePath = parentPath & "\PYTHON_PATH.txt"
    
    ' Open the file for reading
    pythonExePath = GetPythonPath()
    
    PythonScriptPath = parentPath & "\.src\main.py"
    
    arugument = " --action to_text --wait 1 --excelpath " & """" & ThisWorkbook.FullName & """"
    objShell.Run pythonExePath & " " & PythonScriptPath & " " & arugument
End Sub

