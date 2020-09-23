VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Mini Script Example Project"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   6780
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   195
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   1470
      Width           =   1365
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   195
      Left            =   4020
      TabIndex        =   7
      Top             =   2610
      Width           =   2445
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   195
      Left            =   4050
      TabIndex        =   6
      Top             =   2190
      Width           =   2460
   End
   Begin VB.DirListBox Dir1 
      Height          =   765
      Left            =   2040
      TabIndex        =   5
      Top             =   1845
      Width           =   1485
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   2055
      TabIndex        =   4
      Top             =   1455
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1065
      Left            =   3990
      TabIndex        =   3
      Top             =   870
      Width           =   2160
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   2055
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   945
      Width           =   1680
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   180
      TabIndex        =   1
      Top             =   960
      Width           =   1560
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   5  'Dash-Dot-Dot
      FillStyle       =   4  'Upward Diagonal
      Height          =   795
      Left            =   150
      Shape           =   3  'Circle
      Top             =   1935
      Width           =   990
   End
   Begin VB.Line Line1 
      X1              =   105
      X2              =   5940
      Y1              =   690
      Y2              =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Welcome to Mini Script a Visual Basic's Buddy"
      Height          =   195
      Left            =   30
      TabIndex        =   0
      Top             =   105
      Width           =   3270
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    MsgBox "Hello world"
End Sub

Private Sub Form_Load()
    ' This forms loads here
End Sub
