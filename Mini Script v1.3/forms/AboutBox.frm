VERSION 5.00
Begin VB.Form AboutBox 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Mini Scripter"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5055
   Icon            =   "AboutBox.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   248
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   337
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "AboutBox.frx":0442
      Top             =   3240
      Width           =   3375
   End
   Begin VB.PictureBox Picture1 
      Height          =   435
      Left            =   3600
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   4
      Top             =   3240
      Width           =   1275
      Begin VB.CommandButton Command1 
         Caption         =   "Okey!"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   480
      Width           =   4815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   4815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "About Mini Scripter"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "About Mini Scripter"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   180
      TabIndex        =   7
      Top             =   75
      Width           =   3135
   End
End
Attribute VB_Name = "AboutBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Text1 = Text1 & "The idéa was simple. I wanted to hack vb code on computers that didn't have Visual Basic. Just write the code, but not nessesery compile it." & vbCrLf & vbCrLf & _
"Use: I have this program on my USB memory, and the times when I'am bored in school I just plug in my USB device and start to code some VB. But.... I have not enough ennergy to finish this project by my self. I need help by you! Give me ideas and code to this project." & vbCrLf & vbCrLf & _
"Wouldn't it be fun to get this program working? ;)" & vbCrLf & _
"Mail me @ zytric@msn.com" & vbCrLf & vbCrLf & _
"Created by Kristoffer Sörquist" & vbCrLf & vbCrLf & _
"Big thanks to:" & vbCrLf & _
"Ben Jones - dreamvb@yahoo.com" & vbCrLf & _
"" & vbCrLf & _
""
End Sub
