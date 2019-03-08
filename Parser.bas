Attribute VB_Name = "Parser"
Option Explicit
Public Sub Execute()

    Dim filePath As String, FSO As FileSystemObject, ts As TextStream, test As Boolean
    Dim i As Long, char As String, delimiter As String, word As String, words As Collection, dLength As Byte, lenStr As Long, index As Long
    
    DoEvents
    
    test = False
    filePath = tempDir
    If test Then: filePath = "" ''Filepath omitted
    delimiter = "SMTP:"
    dLength = Len(delimiter)
    Set FSO = New FileSystemObject
    Set words = New Collection
    Set ts = FSO.OpenTextFile(filePath)
    gString = ts.ReadAll
    lenStr = Len(gString)
    index = 1
    
    Do
        index = InStr(index, gString, delimiter)
        If Not index = 0 Then
            index = index + dLength
            word = "": char = ""
            Do
                word = word & char
                char = Mid(gString, index, 1)
                index = index + 1
            Loop Until char = " "
            If word <> "" Then
                words.Add (word)
            End If
            Call Iterated(index, lenStr)
            DoEvents
        End If
    Loop Until index = 0
    
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    
    Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
    For i = 1 To words.Count
        ws.Cells(i, "A") = words(i)
    Next i
    Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic
    UserForm1.Hide

End Sub
Private Sub Iterated(i As Long, j As Long)
    If i Mod 200 = 0 Then
        UserForm1.Label1 = i & "/" & j & vbCrLf & CInt((i / j) * 100) & " % Complete"
        DoEvents
    End If
    If UserForm1.Visible = False Then: UserForm1.Show
End Sub
