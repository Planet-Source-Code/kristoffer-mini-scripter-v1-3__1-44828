VERSION 5.00
Begin VB.Form frmPreView 
   ClientHeight    =   2385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5625
   LinkTopic       =   "Form2"
   ScaleHeight     =   2385
   ScaleWidth      =   5625
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Unload this form"
      Height          =   420
      Left            =   4140
      TabIndex        =   0
      Top             =   1920
      Width           =   1440
   End
End
Attribute VB_Name = "frmPreView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    Unload frmPreView
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPreView = Nothing
End Sub
