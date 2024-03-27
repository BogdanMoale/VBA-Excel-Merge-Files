VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Method"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2685
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    If Worksheets("Input").Cells(1, 1) = "TRUE" Then
        ModuleOldFormat.RDB_Merge_Data_BrowseNume
    Else
        ModuleNewFormat.RDB_Merge_Data_BrowseNume
    End If
    
End Sub

Private Sub CommandButton2_Click()
    If Worksheets("Input").Cells(1, 1) = "TRUE" Then
        ModuleOldFormat.RDB_Merge_Data_BrowseIndex
    Else
        ModuleNewFormat.RDB_Merge_Data_BrowseIndex
    End If
    
End Sub

