VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmworkspace 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Workspace"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2805
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   2805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdview 
      Caption         =   "&View Form"
      Enabled         =   0   'False
      Height          =   315
      Left            =   75
      TabIndex        =   1
      Top             =   30
      Width           =   960
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   2370
      Left            =   30
      TabIndex        =   0
      Top             =   420
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   4180
      _Version        =   393217
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   840
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   14
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmworkspace.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmworkspace.frx":02F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmworkspace.frx":05E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmworkspace.frx":08D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmworkspace.frx":0BC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmworkspace.frx":0EBA
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmworkspace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' By ben jones
' email dreamvb@yahoo.com
' if you find this code usfull then please let me know
' if You want to use this code in your programs then please do so
' but please remmber were it came from.

' Last upadted 22-10-02
' Added support for viewing forms and thier controls not all finished yet
' still some bugs loading forms that controls have font properties set on

Private Type TForm
    TFormCaption As String
    TFormTop As Long
    TFormLeft As Long
    TFormHeight As Long
    TFormWidth As Long
End Type

Private Type vbControl
    mType As String
    mName As String
    mHeight As Long
    mWidth As Long
    mTop As Long
    mLeft As Long
    mCaption As String
    bkColor As String
    fColor As String
    mText As String
    mBorderWidth As String
    mBorderColor As String
    mFillStyle As String
    mShape As String
    mEnabled As Integer
End Type


Public ResFile As String
Dim ControlsCNT As Long
Dim TVBForm As TForm
Dim vbControl(100) As vbControl

Sub PhaseForm(FrmData As String)
On Error Resume Next

Dim kFile As Long
Dim StrB As String, strln As String, ControlItem As String, FrmItem As String
Dim xpos, ypos As Long, I As Long

Dim CtrName As String, vbName As String

    controlcnt = 0
    kFile = FreeFile
        
    Open FrmData For Input As #kFile
        Do While Not EOF(kFile)
            Input #kFile, StrB
            StrB = Trim(StrB)
            If UCase$(Left$(StrB, 5)) = "BEGIN" Then
                strln = Trim(Right(StrB, Len(StrB) - 5))
                xpos = InStr(strln, " ")
                
                vbName = UCase(Trim(Mid(strln, 1, xpos)))
                CtrName = Trim(Mid(strln, xpos + 1, Len(strln) - 1))
                
                vbControl(controlcnt).mType = vbName
                vbControl(controlcnt).mName = CtrName
                vbControl(controlcnt).mEnabled = 1
                controlcnt = controlcnt + 1
                ControlsCNT = controlcnt
            End If
            
            
            If vbName = "VB.FORM" Then
                ypos = InStr(StrB, " ")
                FrmItem = Trim(UCase(Mid(StrB, 1, ypos)))
                
                Select Case FrmItem
                    Case "CAPTION"
                        TVBForm.TFormCaption = GetValue(StrB)
                    Case "CLIENTHEIGHT"
                        TVBForm.TFormHeight = Val(GetValue(StrB))
                    Case "CLIENTWIDTH"
                        TVBForm.TFormWidth = Val(GetValue(StrB))
                    Case "CLIENTLEFT"
                        TVBForm.TFormLeft = Val(GetValue(StrB))
                    Case "CLIENTTOP"
                        TVBForm.TFormTop = Val(GetValue(StrB))
                End Select
            End If
            
            xpos = InStr(StrB, " ")
            ControlItem = Trim(UCase(Mid(StrB, 1, xpos)))
            

            Select Case ControlItem
                Case "CAPTION"
                    vbControl(controlcnt - 1).mCaption = GetValue(StrB)
                Case "HEIGHT"
                    vbControl(controlcnt - 1).mHeight = Val(GetValue(StrB))
                Case "WIDTH"
                    vbControl(controlcnt - 1).mWidth = Val(GetValue(StrB))
                Case "LEFT"
                    vbControl(controlcnt - 1).mLeft = Val(GetValue(StrB))
                Case "TOP"
                    vbControl(controlcnt - 1).mTop = Val(GetValue(StrB))
                Case "BACKCOLOR"
                    vbControl(controlcnt - 1).bkColor = Val(GetValue(StrB))
                Case "FORECOLOR"
                    vbControl(controlcnt - 1).fColor = Val(GetValue(StrB))
                Case "TEXT"
                    vbControl(controlcnt - 1).mText = GetValue(StrB)
                Case "BORDERWIDTH"
                    vbControl(controlcnt - 1).mBorderWidth = Val(GetValue(StrB))
                Case "BORDERCOLOR"
                    vbControl(controlcnt - 1).mBorderColor = Val(GetValue(StrB))
                Case "FILLSTYLE"
                    vbControl(controlcnt - 1).mFillStyle = Val(GetValue(StrB))
                Case "SHAPE"
                    vbControl(controlcnt - 1).mShape = Val(GetValue(StrB))
                Case "ENABLED"
                    vbControl(controlcnt - 1).mEnabled = Val(GetValue(StrB))
            End Select
            DoEvents
        Loop
        Close #kFile
        
        StrB = "": strln = "": ControlItem = ""
        FrmItem = "": xpos = 0: ypos = 0: vbName = "": CtrName = ""
        
        frmPreView.Caption = TVBForm.TFormCaption
        frmPreView.Height = TVBForm.TFormHeight + TVBForm.TFormTop + 40
        frmPreView.Width = TVBForm.TFormWidth + TVBForm.TFormLeft + 40
        frmPreView.Left = TVBForm.TFormLeft
        frmPreView.Top = TVBForm.TFormTop
        
        For I = 1 To ControlsCNT - 1
            Set tcontrol = frmPreView.Controls.Add(vbControl(I).mType, vbControl(I).mName, frmPreView)
            tcontrol.Visible = True
            tcontrol.Height = vbControl(I).mHeight
            tcontrol.Width = vbControl(I).mWidth
            tcontrol.Top = vbControl(I).mTop
            tcontrol.Left = vbControl(I).mLeft
            tcontrol.Caption = vbControl(I).mCaption
            tcontrol.BackColor = vbControl(I).bkColor
            tcontrol.ForeColor = vbControl(I).fColor
            tcontrol.Text = vbControl(I).mText
            tcontrol.BorderWidth = vbControl(I).mBorderWidth
            tcontrol.BorderColor = vbControl(I).mBorderColor
            tcontrol.FillStyle = vbControl(I).mFillStyle
            tcontrol.Shape = vbControl(I).mShape
            tcontrol.Enabled = vbControl(I).mEnabled
        Next
        I = 0
        frmPreView.Show
        
End Sub

Sub OpenVBProject(VbProJFilename As String)
Dim hFile As Long, Pos As Long
Dim Item As String, ItemName As String, StrB As String, ProjectFile As String, ItemFileName As String, _
ProjectPath As String, ProjectName As String

    TreeView1.Nodes.Add , , , "Forms", 2, 1
    TreeView1.Nodes.Add , , , "Modules", 2, 1
    TreeView1.Nodes.Add , , , "Class Modules", 2, 1
    TreeView1.Nodes.Add , , , "User Controls", 2, 1
    hFile = FreeFile
    ProjectFile = VbProJFilename
    ProjectPath = GetPath(ProjectFile)
    
    frmworkspace.MousePointer = vbHourglass
    Open ProjectFile For Input As #hFile
        Do While Not EOF(hFile)
            Input #hFile, StrB
                Pos = InStr(StrB, "=")
                Item = Mid(StrB, 1, Pos)
                l = InStr(StrB, ";")
                Select Case Item
                    Case "Form="
                        ItemFileName = GetItemFileName(StrB)
                        ItemName = GetItemName(ProjectPath & Mid(StrB, Pos + 1, Len(StrB)))
                        TreeView1.Nodes.Add 1, tvwChild, ProjectPath & Mid(StrB, Pos + 1, Len(StrB)), ItemName & " (" & ItemFileName & ")", 3, 3
                    
                    Case "Module="
                        ItemName = GetItemName(ProjectPath & Trim(Mid(StrB, l + 1, Len(StrB))))
                        ItemFileName = GetItemFileName(StrB)
                        TreeView1.Nodes.Add 2, tvwChild, ProjectPath & Trim(Mid(StrB, l + 1, Len(StrB))), ItemName & " (" & ItemFileName & ")", 4, 4
                    
                    Case "Class="
                        ItemName = GetItemName(ProjectPath & Trim(Mid(StrB, l + 1, Len(StrB))))
                        ItemFileName = GetItemFileName(StrB)
                        TreeView1.Nodes.Add 3, tvwChild, ProjectPath & Trim(Mid(StrB, l + 1, Len(StrB))), ItemName & " (" & ItemFileName & ")", 5, 5

                    Case "UserControl="
                        ItemName = GetItemName(ProjectPath & Mid(StrB, InStr(StrB, "=") + 1, Len(StrB)))
                        ItemFileName = GetItemFileName(StrB)
                        TreeView1.Nodes.Add 4, tvwChild, ProjectPath & Mid(StrB, InStr(StrB, "=") + 1, Len(StrB)), ItemName & " (" & ItemFileName & ")", 6, 6
                        Case "Name="
                        'ProjectName = "DM VB Project Explorer - " & Mid(StrB, Pos + 2, Len(StrB) - Pos - 2)
                End Select
            DoEvents
        Loop
    Close #hFile
    Pos = 0
    StrB = ""
    ProjectName = ""
    ProjectFile = ""
    ProjectPath = ""
    ItemFileName = ""
    ItemName = ""
    frmworkspace.MousePointer = vbDefault
    
End Sub

Private Sub cmdview_Click()
    PhaseForm ResFile
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmworkspace = Nothing
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "vbpro"
            TreeView1.Nodes.Clear
            CDialog.DialogTitle = "Open Visual Basic Project"
            CDialog.Filter = "Visual Basic Projects(*.vbp)|*.vbp"
            CDialog.ShowOpen
            If Len(CDialog.FileName) <= 0 Then Exit Sub
            If Not UCase(Right(CDialog.FileName, 3)) = "VBP" Then
                MsgBox "This is not a vaild Visual Basic Project", vbCritical, "Invaild Filename"
                Exit Sub
            Else
                OpenVBProject CDialog.FileName
            End If
            
        Case "about"
            frmAbout.Show vbModal
        
        Case "exit"
            Unload Form1
            
        Case "viewcode"
            frmcode.txtcode.Text = LoadFile(ResFile)
            frmcode.Show vbModal, Form1
            
        Case "viewform"
            PhaseForm ResFile
            frmPreView.Show
            
    End Select
    
End Sub

Private Sub TreeView1_Click()

    If Len(TreeView1.SelectedItem.Key) = 0 Then Exit Sub
    FileType = UCase$(Right(TreeView1.SelectedItem.Key, 3))
    ResFile = TreeView1.SelectedItem.Key
    
    If FileType = "FRM" Then
        cmdview.Enabled = True
    Else
        cmdview.Enabled = False
    End If
    
End Sub

Private Sub TreeView1_Collapse(ByVal Node As MSComctlLib.Node)
    cmdview.Enabled = False
End Sub

Private Sub TreeView1_DblClick()
Dim FileType As String
On Error Resume Next
    If Len(TreeView1.SelectedItem.Key) = 0 Then Exit Sub
    FileType = UCase$(Right(TreeView1.SelectedItem.Key, 3))
    ResFile = TreeView1.SelectedItem.Key
    
    Select Case FileType
        Case "FRM"
            Form1.LoadProjectFile VB_FORM, ResFile
        Case "BAS"
            Form1.LoadProjectFile VB_BAS, ResFile
        Case Else
            MsgBox "The option you selected is no available in this version", vbInformation
        End Select
End Sub
