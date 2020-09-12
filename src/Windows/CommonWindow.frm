VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CommonWindow 
   ClientHeight    =   2508
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3828
   OleObjectBlob   =   "CommonWindow.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "CommonWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function WriteStatus(ByVal startStr As String, ByVal endStr As String, ByVal contentStr As String, ByVal objectStr As String)
    CommonWindow.StartPosLabel.Caption = startStr
    CommonWindow.EndPosLabel.Caption = endStr
    CommonWindow.OperationNameLabel.Caption = contentStr
    CommonWindow.OperationObjectLabel.Caption = objectStr
    DoEvents
    
End Function
