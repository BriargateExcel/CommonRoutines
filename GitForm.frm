VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GitForm 
   Caption         =   "Select the VBA Project to Export/Import"
   ClientHeight    =   2535
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5205
   OleObjectBlob   =   "GitForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GitForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelButton_Click()
    LetGitFormCanceled True
    GitForm.Hide
End Sub

Private Sub DeleteBox_Click()
    If DeleteBox Then LetGitFormDelete True
End Sub

Private Sub SelectButton_Click()
    LetGitFormCanceled False
    GitForm.Hide
End Sub

Private Sub UserForm_Activate()
    ' Version 1.0.2
    ' Added positioning to the middle of the application
    
    LetGitFormCanceled False
    LetGitFormDelete False
    With GitForm
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    LetGitFormCanceled True
End Sub

Private Sub UserForm_Terminate()
    LetGitFormCanceled True
End Sub

