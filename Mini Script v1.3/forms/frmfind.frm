VERSION 5.00
Begin VB.Form frmfind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find Text"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkmatch 
      Caption         =   "Match Case"
      Height          =   195
      Left            =   150
      TabIndex        =   4
      Top             =   975
      Width           =   1320
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4125
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   660
      Width           =   885
   End
   Begin VB.CommandButton cmdfind 
      Caption         =   "&Find"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4125
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   135
      Width           =   885
   End
   Begin VB.TextBox txtfind 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   915
      TabIndex        =   1
      Top             =   172
      Width           =   3060
   End
   Begin VB.Label lblfind 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Find Text:"
      Height          =   195
      Left            =   165
      TabIndex        =   0
      Top             =   225
      Width           =   705
   End
End
Attribute VB_Name = "frmfind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Find text dilaog added by ben jones

Private Sub cmdcancel_Click()
    Unload frmfind  ' unload the form
End Sub

Private Sub cmdfind_Click()
Dim nPos As Long
Static Found As Long
Dim mComare As VbCompareMethod
On Error Resume Next

        If chkmatch Then mComare = vbBinaryCompare Else mComare = vbTextCompare
        nPos = InStr(Found + 1, Form1.Txt.Text, txtfind.Text, mComare)
        If nPos = 0 Then
            MsgBox "The string " & Chr(34) & txtfind.Text & Chr(34) & " was not found", vbInformation, frmfind.Caption
            Found = 0
            cmdfind.Caption = "&Find"
            Exit Sub
        Else
            Found = Found + nPos
            Form1.Txt.SelStart = nPos - 1
            Form1.Txt.SelLength = Len(txtfind.Text)
            Form1.Txt.SetFocus
            cmdfind.Caption = "Find &Next"
        End If
End Sub

Private Sub Form_Load()
    frmfind.Icon = Nothing  ' Remove the forms icon
End Sub

Private Sub txtfind_Change()
    If Len(txtfind.Text) > 0 Then
        cmdfind.Enabled = True
        Exit Sub
    Else
        cmdfind.Enabled = False
    End If
    
End Sub
