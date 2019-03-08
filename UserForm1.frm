VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Message"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   4995
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub UserForm_Activate()
    Dim bool As Boolean
    
    DoEvents
    UserForm1.Label1 = "Fetching your autocomplete addresses..."
    If Files.CreateCopyNK2 Then
        Parser.Execute
        Files.DeleteCopyNK2
    Else
        UserForm1.Hide
    End If
    
End Sub


