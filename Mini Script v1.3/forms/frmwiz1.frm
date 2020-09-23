VERSION 5.00
Begin VB.Form frmwiz1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Visual Basic If then wizard"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MiniScripter.Line3D Line3D1 
      Height          =   45
      Left            =   240
      TabIndex        =   11
      Top             =   1785
      Width           =   3780
      _ExtentX        =   6668
      _ExtentY        =   79
   End
   Begin VB.TextBox txtElse 
      Height          =   300
      Left            =   615
      TabIndex        =   4
      Top             =   1275
      Width           =   3330
   End
   Begin VB.CommandButton Command3 
      Caption         =   "C&lear"
      Height          =   375
      Left            =   915
      TabIndex        =   5
      Top             =   1980
      Width           =   1035
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Insert"
      Height          =   375
      Left            =   2010
      TabIndex        =   6
      Top             =   1980
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   1980
      Width           =   1035
   End
   Begin VB.TextBox txtThen 
      Height          =   300
      Left            =   615
      TabIndex        =   3
      Top             =   862
      Width           =   3315
   End
   Begin VB.TextBox txtval1 
      Height          =   300
      Left            =   615
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtval2 
      Height          =   300
      Left            =   2700
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.ComboBox cboOp 
      Height          =   315
      Left            =   1890
      TabIndex        =   8
      Top             =   360
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Else"
      Height          =   195
      Left            =   180
      TabIndex        =   10
      Top             =   1335
      Width           =   300
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Then"
      Height          =   195
      Left            =   150
      TabIndex        =   9
      Top             =   915
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "IF"
      Height          =   195
      Left            =   210
      TabIndex        =   0
      Top             =   413
      Width           =   135
   End
End
Attribute VB_Name = "frmwiz1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' IF then else wizard added by ben jones
Private Function WriteIfStat(Str1 As String, StrOp As String, Str2 As String, Str3 As String, Str4 As String) As String
Dim StrBuf As String
'
    StrBuf = StrBuf & "If " & Str1 & " " & StrOp & " " & Str2 & " Then" & vbNewLine
    StrBuf = StrBuf & Chr(9) & Str3 & vbNewLine
    StrBuf = StrBuf & "Else" & vbNewLine
    StrBuf = StrBuf & Chr(9) & Str4 & vbNewLine
    StrBuf = StrBuf & "End If"
    WriteIfStat = StrBuf
    StrBuff = ""
End Function

Private Sub cboOp_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Command1_Click()
    Unload frmwiz1
    
End Sub

Private Sub Command2_Click()
    Form1.Txt.SelText = vbCrLf & WriteIfStat(txtval1, cboOp.Text, txtval2, txtThen, txtElse) & vbCrLf
    Unload frmwiz1
    
End Sub

Private Sub Command3_Click()
    txtval1.Text = ""
    txtval2.Text = ""
    txtThen.Text = ""
    txtElse.Text = ""
    cboOp.ListIndex = 0
End Sub

Private Sub Form_Load()
    ' Visual Basic Operators
    cboOp.AddItem ">"
    cboOp.AddItem ">="
    cboOp.AddItem "<"
    cboOp.AddItem "<="
    cboOp.AddItem "="
    cboOp.AddItem "<>"
    cboOp.AddItem "AND"
    cboOp.AddItem "OR"
    cboOp.AddItem "NOT"
    cboOp.ListIndex = 0
    Me.Icon = Nothing
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmwiz1 = Nothing
    
End Sub
