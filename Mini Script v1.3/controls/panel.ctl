VERSION 5.00
Begin VB.UserControl Panel 
   ClientHeight    =   1335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1395
   ControlContainer=   -1  'True
   ScaleHeight     =   1335
   ScaleWidth      =   1395
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   3
      X1              =   1320
      X2              =   1320
      Y1              =   90
      Y2              =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   105
      X2              =   1335
      Y1              =   1245
      Y2              =   1245
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   60
      X2              =   60
      Y1              =   105
      Y2              =   1230
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   75
      X2              =   1305
      Y1              =   75
      Y2              =   75
   End
End
Attribute VB_Name = "Panel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' DM Panel ActiveX Control
' Writen and designed by Ben Jones
' Copyright © 2002 Ben Jones

'Default Property Values:
Const m_def_PanelStyle = 0

Enum TPanel
    Upper = 1
    Lower = 0
End Enum

'Property Variables:
Dim m_PanelStyle As Integer
'Event Declarations:
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."


Private Function DoPanel()
    Select Case m_PanelStyle
        Case 0
            Line1(0).BorderColor = &H808080
            Line1(1).BorderColor = &HFFFFFF
            Line1(2).BorderColor = &H808080
            Line1(3).BorderColor = &HFFFFFF
        Case 1
            Line1(0).BorderColor = &HFFFFFF
            Line1(1).BorderColor = &H808080
            Line1(2).BorderColor = &HFFFFFF
            Line1(3).BorderColor = &H808080
    End Select
End Function

Private Sub UserControl_Resize()
    Line1(0).X1 = 0
    Line1(0).Y1 = 0
    Line1(0).Y2 = 0
    Line1(1).X1 = 0
    Line1(2).X1 = 0
    Line1(2).X2 = 0
    Line1(2).Y1 = 0
    Line1(3).Y1 = 0
    
    Line1(3).X1 = UserControl.Width - 8
    Line1(3).X2 = UserControl.Width - 8
    Line1(3).Y2 = UserControl.Height
    Line1(2).Y2 = UserControl.Height
    Line1(0).X2 = UserControl.Width
    Line1(1).Y1 = UserControl.Height - 8
    Line1(1).Y2 = UserControl.Height - 8
    Line1(1).X2 = UserControl.Width
    
End Sub

Public Property Get PanelStyle() As TPanel
    PanelStyle = m_PanelStyle
    If PanelStyle = Lower Then
        m_PanelStyle = 0
    End If
    If PanelStyle = Upper Then
        m_PanelStyle = 1
    End If
    
End Property
Sub ino()

End Sub
Public Property Let PanelStyle(ByVal New_PanelStyle As TPanel)
    m_PanelStyle = New_PanelStyle
    PropertyChanged "PanelStyle"
    DoPanel
End Property

Private Sub UserControl_InitProperties()
    m_PanelStyle = m_def_PanelStyle
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_PanelStyle = PropBag.ReadProperty("PanelStyle", m_def_PanelStyle)
End Sub

Private Sub UserControl_Show()
    DoPanel
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("PanelStyle", m_PanelStyle, m_def_PanelStyle)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

