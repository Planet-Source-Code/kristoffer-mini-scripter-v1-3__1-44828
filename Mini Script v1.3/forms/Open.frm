VERSION 5.00
Begin VB.Form Opena 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "[Open file]"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5655
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   377
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   2655
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   2655
   End
   Begin VB.PictureBox Picture1 
      Height          =   435
      Left            =   4200
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   2
      Top             =   2640
      Width           =   1275
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "D:\[visual basic]\Sources\Ball\Form1.frm"
      Top             =   360
      Width           =   5415
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1650
      Left            =   2880
      TabIndex        =   6
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Open source file path:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "Opena"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim CodeStart As Integer
Dim Item As Long, ListItem As String
Dim nFile As Long
On Error GoTo ENDA:


    nFile = FreeFile
    Open Text1 For Input As #nFile
        Form1.Txt.Text = ""
        Do While Not EOF(1)
            Input #nFile, ListItem
            Item = Item + 1
            
            Form1.Txt.Text = Form1.Txt.Text & ListItem & vbCrLf
        Loop
    Close #nFile
    
    If LCase(Right(Text1, 4)) = ".frm" Then
        If Mid(Form1.Txt.Text, 1, 8) = "VERSION " Then Form1.Status.Panels(3).Text = Mid(Form1.Txt.Text, 8, 4)
        CodeStart = InStr(1, LCase(Form1.Txt.Text), "private")
        If CodeStart <> 0 Then Form1.Txt.Text = Mid(Form1.Txt.Text, CodeStart, Len(Form1.Txt.Text))
    End If

    If LCase(Right(Text1, 4)) = ".bas" Then
        Form1.Status.Panels(3).Text = "0.0"
        CodeStart = InStr(1, LCase(Form1.Txt.Text), vbCrLf)
        If CodeStart <> 0 Then Form1.Txt.Text = Mid(Form1.Txt.Text, CodeStart, Len(Form1.Txt.Text))
    End If

    Form1.Status.Panels(0).Text = Text1.Text
    

    Call Form1.RefreshText
    Unload Opena
    Exit Sub

ENDA:
    MsgBox "Couldn't load file" & vbCrLf & "Error: " & Error
    Unload Opena
    MsgBox Form1.Txt.Text
    
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub
Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub
Private Sub File1_Click()
Text1 = File1.Path & "\" & File1.FileName
End Sub
Private Sub File1_DblClick()
Command1_Click
End Sub

