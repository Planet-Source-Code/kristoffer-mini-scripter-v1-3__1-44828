VERSION 5.00
Begin VB.UserControl Line3D 
   AutoRedraw      =   -1  'True
   ClientHeight    =   45
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   2895
   ScaleHeight     =   45
   ScaleWidth      =   2895
End
Attribute VB_Name = "Line3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' DM 3DLine ActiveX Control
' Writen and designed by Ben Jones

Private Sub UserControl_Initialize()
    UserControl.Line (UserControl.Width, 8)-(0, 8), &HFFFFFF
    UserControl.Line (UserControl.Width, 0)-(0, 0), &H808080
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    UserControl.Height = 50
    UserControl.Line (UserControl.Width, 8)-(0, 8), &HFFFFFF
    UserControl.Line (UserControl.Width, 0)-(0, 0), &H808080
End Sub
