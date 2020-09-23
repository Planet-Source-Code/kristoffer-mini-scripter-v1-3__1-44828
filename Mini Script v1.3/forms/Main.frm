VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Mini Scripter v1.3"
   ClientHeight    =   5445
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8445
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   363
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   563
   StartUpPosition =   2  'CenterScreen
   Begin MiniScripter.Line3D Line3D1 
      Height          =   45
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   79
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   600
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":08A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0BF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0F4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":129C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":15EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1940
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Left            =   165
      TabIndex        =   7
      Top             =   75
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NEW"
            Object.ToolTipText     =   "New Visual Basic File"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "OPEN"
            Object.ToolTipText     =   "Open Visual Basic Project"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SAVE"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "CUT"
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "COPY"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PASTE"
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "FIND"
            Object.ToolTipText     =   "Find"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   5130
      Width           =   8445
      _ExtentX        =   14896
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "[file]"
            TextSave        =   "[file]"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Visual Basic"
            TextSave        =   "Visual Basic"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   58208
            MinWidth        =   58208
            Text            =   "By Kristoffer Sörquist, Help by Ben Jones"
            TextSave        =   "By Kristoffer Sörquist, Help by Ben Jones"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Toppen 
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   45
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   494
      TabIndex        =   4
      Top             =   600
      Width           =   7410
      Begin VB.ComboBox Funcs 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "Main.frx":1C92
         Left            =   420
         List            =   "Main.frx":1C94
         TabIndex        =   6
         Top             =   30
         Width           =   2880
      End
      Begin VB.ComboBox Proj 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "Main.frx":1C96
         Left            =   3345
         List            =   "Main.frx":1C98
         TabIndex        =   5
         Text            =   "[Work Space]"
         Top             =   30
         Width           =   3255
      End
      Begin MiniScripter.Panel Panel1 
         Height          =   435
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   767
         PanelStyle      =   1
         Begin VB.Image imgfunc 
            Height          =   240
            Left            =   60
            Picture         =   "Main.frx":1C9A
            ToolTipText     =   "Refresh Functions List"
            Top             =   60
            Width           =   240
         End
      End
   End
   Begin VB.PictureBox MainBottom 
      Height          =   3735
      Left            =   0
      ScaleHeight     =   245
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   510
      TabIndex        =   2
      Top             =   555
      Width           =   7710
      Begin MSComDlg.CommonDialog CDialog 
         Left            =   1200
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin RichTextLib.RichTextBox Txt 
         Height          =   3060
         Left            =   600
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   435
         Width           =   6075
         _ExtentX        =   10716
         _ExtentY        =   5398
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         ScrollBars      =   3
         Appearance      =   0
         RightMargin     =   99999
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"Main.frx":2224
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.PictureBox picmargin 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2955
         Left            =   465
         ScaleHeight     =   2955
         ScaleWidth      =   135
         TabIndex        =   3
         Top             =   450
         Width           =   135
      End
      Begin VB.Line lnmargin 
         BorderColor     =   &H00808080&
         X1              =   30
         X2              =   30
         Y1              =   30
         Y2              =   120
      End
   End
   Begin MiniScripter.Line3D Line3D2 
      Height          =   45
      Left            =   0
      TabIndex        =   9
      Top             =   465
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   79
   End
   Begin VB.Image Image1 
      Height          =   330
      Left            =   30
      Picture         =   "Main.frx":22A4
      Top             =   75
      Width           =   90
   End
   Begin VB.Menu Menu 
      Caption         =   "&Menu"
      Begin VB.Menu New 
         Caption         =   "New"
         Shortcut        =   {F1}
      End
      Begin VB.Menu Open 
         Caption         =   "Open Project"
         Shortcut        =   ^O
      End
      Begin VB.Menu Save 
         Caption         =   "Save"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu Line01 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "Edit"
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnupaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
         Enabled         =   0   'False
      End
      Begin VB.Menu Line02 
         Caption         =   "-"
      End
      Begin VB.Menu Select 
         Caption         =   "Select all"
      End
      Begin VB.Menu mnufind 
         Caption         =   "&Find Text"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuReplace 
         Caption         =   "&Replace Text"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuindent 
         Caption         =   "&Indent"
      End
      Begin VB.Menu mnuoutdent 
         Caption         =   "&Outdent"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuwrsp 
         Caption         =   "Project Workspace"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnutools 
      Caption         =   "&Tools"
      Begin VB.Menu addproc 
         Caption         =   "&Add procedure"
      End
      Begin VB.Menu mnuwiz 
         Caption         =   "&Wizards"
         Begin VB.Menu mnuifwiz 
            Caption         =   "&If then else"
         End
      End
   End
   Begin VB.Menu About 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public SomethingChanged As Boolean
Dim m_indent As Integer

Private Const IndentVal = 10
Function Checked(TChecked As Boolean) As Boolean
    If TChecked Then
        Checked = False
    ElseIf TChecked = False Then
        Checked = True
    End If
End Function

Public Function LoadProjectFile(Op As VB_FILETYPE, lzFile As String)
Dim StrFileBuff As String
    Select Case Op
        Case VB_FORM
            ' Here you may want to add something that ask the user
            ' if they like to save there changes the project
            Txt.Text = ""
            StrFileBuff = OpenFile(lzFile)
            Status.Panels(1).Text = "Form Name: " + GetFilename(StrFileBuff) ' Update statsbar with the forms name
            Status.Panels(3).Text = GetVBVer(StrFileBuff) ' update statusbar with the VB ver number
            Txt.Text = GetMainCode(StrFileBuff, VB_FORM)
            lstVBFunctions Txt.Text, Funcs ' Update the list of the function name in the function combo box
            RefreshText ' Call the colour code function
            StrFileBuff = "" ' Clear the buffer
        Case VB_BAS
            StrFileBuff = OpenFile(lzFile)
            Status.Panels(1).Text = "Basic Name: " + GetFileTitle(lzFile)
            Status.Panels(3).Text = "0.0" ' there is no number in a bas file so we just make one up for now
            Txt.Text = GetMainCode(StrFileBuff, VB_BAS)
            lstVBFunctions Txt.Text, Funcs ' Update the list of the function name in the function combo box
            RefreshText ' Call the colour code function
            StrFileBuff = "" ' Clear the buffer
    End Select
    
End Function


Private Sub EnableMenuItems()
    If Len(Txt.SelText) > 0 Then
        mnuCut.Enabled = True
        mnuCopy.Enabled = True
        Toolbar1.Buttons(6).Enabled = True
        Toolbar1.Buttons(7).Enabled = True
    Else
        mnuCut.Enabled = False
        mnuCopy.Enabled = False
        Toolbar1.Buttons(6).Enabled = False
        Toolbar1.Buttons(7).Enabled = False
    End If
    
End Sub
Private Sub About_Click()
AboutBox.Show vbModal
End Sub

Private Sub addproc_Click()
    frmproc.Show vbModal, Form1
End Sub

Private Sub Exit_Click()
Dim ans
    If SomethingChanged Then
        ans = MsgBox("You have unsaved work do you want to quit now", vbQuestion Or vbYesNo, Form1.Caption)
        If ans = vbNo Then Unload Form1
        MsgBox " we need to add some saveing code here"
    Else
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
PreLoad
Form_Resize
mnuwrsp.Checked = False
End Sub

Private Sub Form_Paint()
    Form_Resize
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Line3D1.Width = Form1.ScaleWidth - 1
    Line3D2.Width = Form1.ScaleWidth - 1
    MainBottom.Width = Form1.ScaleWidth - MainBottom.Left
    Panel1.Width = MainBottom.ScaleWidth - 1
    lnmargin.Y2 = MainBottom.ScaleHeight
    Toppen.Width = MainBottom.ScaleWidth
    MainBottom.Height = Status.Top - Status.Height - 16
    picmargin.Height = MainBottom.ScaleHeight
    Txt.Width = MainBottom.Width - 43.6
    Txt.Height = MainBottom.Height - 33
    Proj.Width = (Proj.Width + Funcs.Width) / 3
    Proj.Left = (MainBottom.Width - Proj.Width) - 10
    Funcs.Width = Proj.Left - Funcs.Left - 10
If Err Then Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Form1 = Nothing
    Set frmfind = Nothing
    Set frmPreView = Nothing
    Set frmworkspace = Nothing
    Set AboutBox = Nothing
End Sub

Private Sub Funcs_Click()
' Code added by ben jones
Dim lpos As Long, ypos As Long
    lpos = InStr(1, Txt.Text, Funcs.Text, vbTextCompare)
    If lpos <= 0 Then
        Exit Sub
    Else
        ypos = InStr(lpos + 1, Txt.Text, vbCrLf, vbTextCompare)
        Txt.SelStart = ypos + 2
        Txt.SetFocus
    End If
    
End Sub

Private Sub imgfunc_Click()
    lstVBFunctions Txt.Text, Funcs
End Sub

Private Sub mnuCopy_Click()
    EditMenu M_COPY, Txt
End Sub

Private Sub mnuCut_Click()
    EditMenu M_CUT, Txt
End Sub

Private Sub mnufind_Click()
    frmfind.Show , Form1
End Sub

Private Sub mnuifwiz_Click()
    frmwiz1.Show , Form1
End Sub

Private Sub mnuindent_Click()
    m_indent = m_indent + IndentVal
    Txt.SelIndent = m_indent
End Sub

Private Sub mnuoutdent_Click()
    If m_indent = 0 Then m_indent = IndentVal
    m_indent = m_indent - IndentVal
    Txt.SelIndent = m_indent
End Sub

Private Sub mnupaste_Click()
    EditMenu M_PASTE, Txt ' Edit menu paste command
End Sub

Private Sub mnuReplace_Click()
    frmreplace.Show , Form1 ' show the replace text dialog
End Sub

Private Sub mnuwrsp_Click()
    mnuwrsp.Checked = Checked(mnuwrsp.Checked)
    If mnuwrsp.Checked = False Then
        frmworkspace.Hide
    Else
        frmworkspace.Show
        frmworkspace.Top = Form1.Top
        frmworkspace.Left = Form1.Left - frmworkspace.Width
    End If
    
End Sub

Private Sub New_Click()
'Not ready
Txt.Text = "Private Sub Form_Load()" & vbCrLf & vbCrLf & "End Sub"
Call RefreshText

End Sub

Private Sub Open_Click()
Dim StrFileBuff As String
' Code added by ben jones
' This code allows you to open files useing the dialog control
    With CDialog
        .DialogTitle = "Open Visual Basic Project"
        .Filter = "Visual Basic Project(*.vbp)|*.vbp|"
        .ShowOpen
        If Len(.FileName) = 0 Then Exit Sub
        If Not UCase(GetExtension(.FileName)) = "VBP" Then
            MsgBox "You need to select a vaid Visual Basic project", vbInformation, "Invaild Project file"
            Exit Sub
        Else
            mnuwrsp_Click
            frmworkspace.OpenVBProject (.FileName) ' Load in the workspace for the project
        End If
    End With
    
End Sub

Private Sub Select_Click()
    EditMenu M_SELECTALL, Txt ' Edit menu selectall command
    EnableMenuItems
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "OPEN"
            Open_Click ' Call open menu open sub
        Case "CUT"
            mnuCut_Click ' Call the menu cut sub
        Case "COPY"
            mnuCopy_Click ' Call the menu copy sub
        Case "PASTE"
            mnupaste_Click ' Call the menu paste sub
        Case "FIND"
            mnufind_Click
            
    End Select
End Sub

Private Sub Txt_Change()
    SomethingChanged = True
    If Len(Txt.Text) <= 0 Then
        mnufind.Enabled = False
        Toolbar1.Buttons(9).Enabled = False
    Else
        mnufind.Enabled = True
        mnuReplace.Enabled = True
        Toolbar1.Buttons(9).Enabled = True
    End If
    
End Sub

Private Sub Txt_KeyUp(KeyCode As Integer, Shift As Integer)
If SomethingChanged = False Then Exit Sub
If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then RefreshText
End Sub

Private Sub Txt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If SomethingChanged = True Then
        RefreshText
    End If
End Sub

Private Sub Txt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    EnableMenuItems
End Sub
