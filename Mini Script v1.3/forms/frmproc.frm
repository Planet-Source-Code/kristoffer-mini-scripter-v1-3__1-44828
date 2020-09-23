VERSION 5.00
Begin VB.Form frmproc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add procedure"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Procedure Scope"
      Height          =   810
      Left            =   105
      TabIndex        =   11
      Top             =   1920
      Width           =   2655
      Begin VB.OptionButton Option1 
         Caption         =   "Private"
         Height          =   210
         Left            =   1410
         TabIndex        =   9
         Top             =   330
         Width           =   1050
      End
      Begin VB.OptionButton optscrope 
         Caption         =   "Public"
         Height          =   210
         Left            =   180
         TabIndex        =   8
         Top             =   330
         Value           =   -1  'True
         Width           =   780
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Procedure Type"
      Height          =   1050
      Left            =   120
      TabIndex        =   10
      Top             =   705
      Width           =   2595
      Begin VB.OptionButton optproctype 
         Caption         =   "Event"
         Height          =   225
         Index           =   3
         Left            =   1425
         TabIndex        =   7
         Top             =   645
         Width           =   960
      End
      Begin VB.OptionButton optproctype 
         Caption         =   "Sub"
         Height          =   225
         Index           =   0
         Left            =   165
         TabIndex        =   4
         Top             =   285
         Value           =   -1  'True
         Width           =   585
      End
      Begin VB.OptionButton optproctype 
         Caption         =   "Function"
         Height          =   225
         Index           =   1
         Left            =   165
         TabIndex        =   6
         Top             =   645
         Width           =   1035
      End
      Begin VB.OptionButton optproctype 
         Caption         =   "Property"
         Height          =   225
         Index           =   2
         Left            =   1425
         TabIndex        =   5
         Top             =   285
         Width           =   960
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   2850
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Height          =   345
      Left            =   2850
      TabIndex        =   2
      Top             =   180
      Width           =   975
   End
   Begin VB.TextBox txtprocname 
      Height          =   315
      Left            =   690
      TabIndex        =   1
      Top             =   195
      Width           =   1875
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   255
      Width           =   465
   End
End
Attribute VB_Name = "frmproc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'added Add procedure dialog by ben jones
Private ScopeType As String
Private PocType As Integer
Private ProcData As String

Private Sub cmdok_Click()
    If Trim(Len(txtprocname.Text)) <= 0 Then
        MsgBox "You need to enter in a procedure eg myFunction", vbInformation
        Exit Sub
    Else
        Select Case PocType
            Case 0
                ProcData = vbCrLf & ScopeType & " " & "Sub" & " " & txtprocname.Text & "()" _
                & vbCrLf & vbCrLf & "End Sub"
            Case 1
                ProcData = vbCrLf & ScopeType & " " & "Function" & " " & txtprocname.Text & "()" _
                & vbCrLf & vbCrLf & "End Function"
            Case 2
                ProcData = vbCrLf & ScopeType & " " & "Property" & " Get " & txtprocname.Text & "() As Variant" _
                & vbCrLf & vbCrLf & "End Property" & vbCrLf & vbCrLf & ScopeType & " Let " & txtprocname.Text _
                & "(ByVal vNewValue As Variant) " & vbCrLf & vbCrLf & "End Property"
            Case 3
                ProcData = vbCrLf & "Public Event " & txtprocname.Text & "()"
        End Select
    End If
    
    Form1.Txt.SelStart = Len(Form1.Txt)
    Form1.Txt.SelText = ProcData
    ProcData = ""
    ScopeType = ""
    PocType = 0
    Form1.RefreshText
    Unload frmproc
End Sub

Private Sub Command2_Click()
    Unload frmproc
End Sub

Private Sub Form_Load()
    Option1_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmproc = Nothing
End Sub


Private Sub Option1_Click()
    ScopeType = "Private"
End Sub

Private Sub optproctype_Click(Index As Integer)
    PocType = Index
    If Index = 3 Then
        optscrope.Enabled = False
        Option1.Enabled = False
    Else
        optscrope.Enabled = True
        Option1.Enabled = True
    End If
    
End Sub

Private Sub optscrope_Click()
    ScopeType = "Public"
End Sub




