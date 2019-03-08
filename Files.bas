Attribute VB_Name = "Files"
Public Function CreateCopyNK2() As Boolean
    Dim FSO As FileSystemObject, outApp As Outlook.Application, outFile As File, PID1, PID2, PID3 As Variant
    
    On Error Resume Next
    Set FSO = New FileSystemObject
    Set outApp = GetObject(, "Outlook.Application")
    
    If outApp.ActiveWindow Is Nothing Then
        NK2Dir = "C:\Users\" & Environ("Username") & "\AppData\Roaming\Microsoft\Outlook\Outlook.NK2"
        tempDir = "P:\Outlook.txt"
        
        FileCopy NK2Dir, tempDir
        
        PID1 = Shell("notepad " & tempDir, vbNormalFocus)
        SendKeys "^s"
        PID3 = Shell("TaskKill /F /PID " & CStr(PID1), vbNormalFocus)
        
        CreateCopyNK2 = True
    Else
        MsgBox "Please close Microsoft Outlook and try again."
    End If
    
    
End Function
Public Sub DeleteCopyNK2()
    Kill tempDir
End Sub
