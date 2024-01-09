VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Report Checker"
   ClientHeight    =   7660
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   8200
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////
'///////////////////////////////////////////
'
'Written by Jean-Pierre Crozemarie
'January 2024
'Version : Demo
'Description: This tools help me to quickly check if files are available in different folders.
'
'
'
'///////////////////////////////////////////
'///////////////////////////////////////////


Private Sub Btn_FT_Click()
FileType = "FT"
Module1.IsFileAvailable (FileType)

End Sub

Private Sub Btn_PT_Click()
FileType = "PT"
Module1.IsFileAvailable (FileType)

End Sub

Private Sub Btn_SPD_Click()
FileType = "SPD"
Module1.IsFileAvailable (FileType)

End Sub


Private Sub UserForm_Click()

End Sub
