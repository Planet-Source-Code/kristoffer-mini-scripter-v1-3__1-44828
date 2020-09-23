Attribute VB_Name = "Functions"
 
 
 
''''''''''''''''''''''''''''''''''''''''''''''''''
'' [Mini Scripter] v1.0 BETA                    ''
''                                              ''
'' This code is created by Kristoffer Sörquist  ''
'' If you want to use this code in your project ''
'' then keep this text, please. zytric@msn.com  ''
''                                              ''
'' Thanks to:                                   ''
'' Ben Jones                                    ''
''                                              ''
''''''''''''''''''''''''''''''''''''''''''''''''''


Option Explicit
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public Const BlueWords = "Const*Else*ElseIf*End If*If*Alias*And*As*Base*Binary*Boolean*Byte*ByVal*Call*Case*CBool*CByte*CCur*CDate*CDbl*CDec*CInt*CLng*Close*Compare*Const*CSng*CStr*Currency*CVar*CVErr*Decimal*Declare*DefBool*DefByte*DefCur*DefDate*DefDbl*DefDec*DefInt*DefLng*DefObj*DefSng*DefStr*DefVar*Dim*Do*Double*Each*Else*ElseIf*End*Enum*Eqv*Erase*Error*Exit*Explicit*False*For*Function*Get*Global*GoSub*GoTo*If*Imp*In*Input*Input*Integer*Is*LBound*Let*Lib*Like*Line*Lock*Long*Loop*LSet*Name*New*Next*Not*Object*On*Open*Option*Or*Output*Print*Private*Property*Public*Put*Random*Read*ReDim*Resume*Return*RSet*Seek*Select*Set*Single*Spc*Static*String*Stop*Sub*Tab*Then*Then*True*Type*UBound*Unlock*Variant*Wend*While*With*Xor*Nothing*To"
Public SyntaxArray() As String


Enum VB_FILETYPE
    VB_BAS = 1
    VB_FORM
    VB_CLASS
End Enum


Enum EDIT_MENU
    M_CUT = 1
    M_COPY
    M_PASTE
    M_SELECTALL
End Enum

Public Function RefreshText()
 If Not Right(Form1.Txt.Text, 2) = vbCrLf Then Form1.Txt.Text = Form1.Txt.Text & vbCrLf 'fix for the loop =)
 Call LockWindowUpdate(Form1.hWnd)
 Call ColorText(Form1.Txt, Form1.Txt.SelStart)
 Form1.SomethingChanged = False
 Call LockWindowUpdate(0&)
End Function


Public Function PreLoad()
'Preloads the array so it dosnt take time on the refrehsments
SyntaxArray = Split(BlueWords, "*")
End Function

Private Function Colorize(Box As RichTextBox, NewText As String, Start As Integer, Length As Integer, Color As String)
'This function just colorising the text....
Box.SelStart = Start
Box.SelLength = Length
Box.SelColor = Color
If Not NewText = "" Then Box.SelText = NewText
End Function

Public Function ColorText(Box As RichTextBox, SelectStart As Integer)
'The main function
Dim Words() As String
Dim TimeStamp As Long, NewLine As String
Dim X As Integer, Y As Integer
Dim StartPos As Integer, Pos1 As Integer, LoopStuck As Integer

'Paint text black
StartPos = 1
Call Colorize(Box, "", 0, Len(Box.Text), vbBlack)

'Loop all the text
Do While InStr(StartPos, Box.Text, vbCrLf) <> 0 And LoopStuck < Len(Box.Text)
LoopStuck = LoopStuck + 1

'Get first line
NewLine = Mid(Box.Text, StartPos, InStr(StartPos, Box.Text, vbCrLf) - StartPos)

'Loop word by word
Words = Split(NewLine, " ")
Pos1 = 0
For X = 0 To UBound(Words)
Pos1 = Pos1 + Len(Words(X)) + 1

'Comment
If InStr(1, NewLine, "'") <> 0 Then
 Call Colorize(Box, "", (InStr(1, NewLine, "'") + StartPos) - 2, Len(NewLine) - InStr(1, NewLine, "'") + 1, &H8000&)
 NewLine = Left(NewLine, InStr(1, NewLine, "'") - 1)
End If
     
'Syntax word
For Y = 0 To UBound(SyntaxArray)
 If LCase(Words(X)) = LCase(SyntaxArray(Y)) Then
 Call Colorize(Box, SyntaxArray(Y), Int((Pos1 + StartPos) - 2) - Len(Words(X)), Len(Words(X)), &H800000)
 End If
Next Y

Next X
StartPos = InStr(StartPos, Box.Text, vbCrLf) + 2
Loop

'Set the pointer back and fix the color
Call Colorize(Box, "", SelectStart, 0, vbBlack)

End Function

Public Function EditMenu(mnu_Command As EDIT_MENU, TxtArea As RichTextBox)
' Code for the edit menus added by ben jones
    Select Case mnu_Command
        Case M_CUT
            Clipboard.SetText TxtArea.SelText
            TxtArea.SelText = ""
            TxtArea.SetFocus
        Case M_COPY
            Clipboard.SetText TxtArea.SelText
            TxtArea.SetFocus
        Case M_PASTE
            TxtArea.SelText = Clipboard.GetText
            TxtArea.SetFocus
        Case M_SELECTALL
            TxtArea.SelStart = 0
            TxtArea.SelLength = Len(TxtArea.Text)
            TxtArea.SetFocus
    End Select
    
End Function

Public Function GetVBVer(lzData As String) As String
Dim lpos As Long, ypos As Long
    lpos = InStr(1, lzData, "VERSION", vbTextCompare)
    If lpos = 0 Then
        GetVBVer = "0.0"
        Exit Function
    Else
        ypos = InStr(lpos, lzData, vbCrLf, vbTextCompare)
        GetVBVer = Trim(Mid(lzData, lpos + 7, ypos - lpos - 7))
        ypos = 0: ypos = 0
    End If
    
End Function

Public Function GetFilename(lzData As String) As String
Dim lpos, ypos As Long
    lpos = InStr(1, lzData, "VB.Form", vbTextCompare)
    If lpos = 0 Then
        GetFilename = ""
        Exit Function
    Else
        ypos = InStr(lpos + 1, lzData, vbCrLf, vbTextCompare)
        GetFilename = Trim(Mid(lzData, lpos + 7, ypos - lpos - 7))
        lpos = 0: ypos = 0
    End If
    
End Function

Public Function GetMainCode(lzData As String, FileType As VB_FILETYPE) As String
Dim lpos As Long, ypos As Long

    Select Case FileType
        Case VB_FORM ' Visual Basic form file
            lpos = InStr(1, lzData, "Private", vbTextCompare)
            GetMainCode = Mid(lzData, lpos, Len(lzData))
        Case VB_BAS ' Visual Basic Bas file
            lpos = InStr(1, lzData, "VB_Name =", vbTextCompare)
            If lpos = 0 Then GetMainCode = "" ' you can change this to a error message if you like
            ypos = InStr(lpos + 1, lzData, vbCrLf)
            GetMainCode = LTrim(Mid(lzData, ypos + 2, Len(lzData)))
    End Select
    
End Function

Function lstVBFunctions(lzCode As String, cboLst As ComboBox)
Dim icnt As Long, Ipart, Lpart As Long, X As Long, Y As Long, ch As Long
Dim LnStr, StrBuff As String, FuncName As String, SubName As String, strln As String
On Error Resume Next
' VB Function names Added by Ben jones
' Ok this does work to a level but may need touching up in parts but as am example it will ok
    cboLst.Clear
    StrBuff = lzCode & vbCrLf
    For icnt = 1 To Len(StrBuff)
        ch = Asc(Mid$(StrBuff, icnt, 1))
        If ch <> 13 Then
            strln = strln & Chr(ch)
        Else
            Ipart = InStr(1, strln, "Function ", vbTextCompare) ' Start of function name
            Lpart = InStr(1, strln, "(")    ' End of function name
            If Ipart > 0 And Lpart > 0 Then
                FuncName = Trim$(Mid$(strln, Ipart + Len("Function"), Lpart - Ipart - Len("Function")))
                cboLst.AddItem FuncName
            End If
            
            X = InStr(1, strln, "Sub ", vbTextCompare)
            Y = InStr(1, strln, "(")
            
            If X > 0 And Y > 0 Then
                SubName = Trim(Mid(strln, X + Len("Sub"), Y - X - Len("Sub")))
                cboLst.AddItem SubName
            End If
            strln = ""
            icnt = icnt + 1
        End If
    Next icnt
    cboLst.ListIndex = 0
    icnt = 0: Ipart = 0: Lpart = 0: X = 0: Y = 0
    FuncName = ""
    StrBuff = ""
    LnStr = ""
    ch = ""
    
End Function

