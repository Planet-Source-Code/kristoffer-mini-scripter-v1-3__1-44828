VERSION 5.00
Begin VB.Form frmreplace 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Replace Text"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmrepall 
      Caption         =   "Replace &All"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   660
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1185
      Width           =   1155
   End
   Begin VB.TextBox txtreplace 
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
      Left            =   1230
      TabIndex        =   2
      Top             =   645
      Width           =   2715
   End
   Begin VB.CheckBox chkmatch 
      Caption         =   "Match Case"
      Height          =   195
      Left            =   135
      TabIndex        =   3
      Top             =   1305
      Width           =   1320
   End
   Begin VB.CommandButton cmdreplace 
      Caption         =   "&Replace"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   1155
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
      Left            =   1230
      TabIndex        =   1
      Top             =   172
      Width           =   2715
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Replace With"
      Height          =   195
      Left            =   135
      TabIndex        =   7
      Top             =   690
      Width           =   975
   End
   Begin VB.Label lblfind 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Find Text:"
      Height          =   195
      Left            =   135
      TabIndex        =   0
      Top             =   225
      Width           =   705
   End
End
Attribute VB_Name = "frmreplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Replace text dialog by Ben Jones

Private Sub cmdreplace_Click()
Dim npos As Long
Static Found As Long
Dim mComare As VbCompareMethod
On Error Resume Next

        If chkmatch Then mComare = vbBinaryCompare Else mComare = vbTextCompare
        ' The above line is used for text compare mehods
        npos = InStr(Found + 1, Form1.Txt.Text, txtfind.Text, mComare)
        If npos = 0 Then ' have we found anything
            ' Display message string was not found
            MsgBox "The string " & Chr(34) & txtfind.Text & Chr(34) & " was not found", vbInformation, frmfind.Caption
            Found = 0 ' retset our counter
            cmdfind.Caption = "&Find" ' update the caption
            Exit Sub ' we stop here
        Else
            Found = Found + npos ' Add the text length to our pos
            Form1.Txt.SelStart = npos - 1 ' set the position were to start form
            Form1.Txt.SelLength = Len(txtfind.Text) ' set the length we like to have
            Form1.Txt.SelText = txtreplace.Text
            Form1.Txt.SetFocus ' set focus on the text let user know it was found
            cmdfind.Caption = "Find &Next" ' update the caption
            cmrepall.Enabled = True
            
        End If

End Sub

Private Sub cmrepall_Click()
Dim npos As Long
Static Found As Long
Dim mComare As VbCompareMethod
On Error Resume Next
    
    Do While n < Len(Form1.Txt.Text)
        n = n + 1
        If chkmatch Then mComare = vbBinaryCompare Else mComare = vbTextCompare
        ' The above line is used for text compare mehods
        npos = InStr(Found + 1, Form1.Txt.Text, txtfind.Text, mComare)
        If npos = 0 Then ' have we found anything
            ' Display message string was not found
            cmdreplace.Enabled = False
            cmrepall.Enabled = False
            Found = 0 ' retset our counter

            'Exit Sub ' we stop here
        Else
            Found = Found + npos ' Add the text length to our pos
            Form1.Txt.SelStart = npos - 1 ' set the position were to start form
            Form1.Txt.SelLength = Len(txtfind.Text) ' set the length we like to have
            Form1.Txt.SelText = txtreplace.Text
            Form1.Txt.SetFocus ' set focus on the text let user know it was found
        End If
        DoEvents
    Loop
    
End Sub

Private Sub Command1_Click()
    Unload frmreplace   ' unload the form
End Sub

Private Sub Form_Load()
    frmreplace.Icon = Nothing  ' Remove the forms icon
    cmdreplace.Enabled = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmreplace = Nothing
End Sub

Private Sub txtfind_Change()
    If Len(Trim(txtfind.Text)) > 0 Then
        cmdreplace.Enabled = True
    Else
        cmdreplace.Enabled = False
    End If
    
End Sub
