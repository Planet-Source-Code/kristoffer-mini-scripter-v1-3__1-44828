Attribute VB_Name = "modTools"
Function GetExtension(lzPathFile As String) As String
Dim ipos As Long, I As Long
    For I = Len(lzPathFile) To 1 Step -1
        If InStr(I, lzPathFile, ".", vbBinaryCompare) Then
            ipos = I
            Exit For
        End If
    Next
    
    If ipos = 0 Then
        GetExtension = ""
    Else
        GetExtension = Mid(lzPathFile, I + 1, Len(lzPathFile))
    End If
    I = 0
End Function

Public Function OpenFile(lzFile As String) As String
Dim nFile As Long, fData As String
    nFile = FreeFile ' Pointer to the file
    Open lzFile For Binary As #nFile ' open the file in binary mode
        fData = Space(LOF(nFile))
        Get #nFile, , fData   ' get file contents
    Close #nFile    ' close the file

    OpenFile = fData
    fData = ""
    
End Function
Function GetFileTitle(lzFilename As String)
Dim I As Long, lpos As Long
    For I = 1 To Len(lzFilename)
        If Mid(lzFilename, I, 1) = "\" Then
            lpos = I
        End If
    
    Next
    
    If lpos = 0 Then
        GetFileTitle = ""
        Exit Function
    Else
        GetFileTitle = Mid(lzFilename, lpos + 1, Len(lzFilename))
    End If
    
    
End Function
Function GetItemFileName(StrBuff As String) As String
Dim Ipart As Integer
    Ipart = InStr(StrBuff, "\")
    If Ipart > 0 Then
        GetItemFileName = Mid(StrBuff, Ipart + 1, Len(StrBuff) - Ipart)
    Else
        Ipart = InStr(StrBuff, "=")
        GetItemFileName = Mid(StrBuff, Ipart + 1, Len(StrBuff) - Ipart)
    End If
    Ipart = 0
    
End Function
Function GetItemName(lzFilename As String)
Dim TFile As Long, nPos As Integer
Dim StrItem As String
    TFile = FreeFile
    On Error Resume Next
    Open lzFilename For Input As #TFile
        Do While Not EOF(TFile)
            Input #TFile, StrItem
                nPos = InStr(1, StrItem, "Attribute VB_Name = ")
                If nPos = 1 Then
                    GetItemName = Trim(Mid(StrItem, nPos + 21, Len(StrItem) - nPos - 21))
                End If
            DoEvents
        Loop
    Close #TFile
    nPos = 0
    StrItem = ""
    
End Function
Function GetPath(lzPath As String) As String
Dim Ipart As Integer
    For I = Len(lzPath) To 1 Step -1
        StrC = Mid(lzPath, I, 1)
        If StrC = "\" Then Ipart = I: Exit For
    Next
    GetPath = Trim(Mid(lzPath, 1, Ipart))
    StrC = ""
    I = 0
    Ipart = 0
    
End Function
Public Function GetValue(StrBuff As String) As String
Dim StrB As String
Dim ipos As Long
    
    ipos = InStr(StrBuff, "=")
    StrB = Trim(Mid(StrBuff, ipos + 1, Len(StrBuff)))
    
    If Right(StrB, 1) = Chr(34) Then
        StrB = Left(StrB, Len(StrB) - 1)
        StrB = Right(StrB, Len(StrB) - 1)
        GetValue = StrB
    Else
        GetValue = StrB
    End If
    
    ipos = 0: StrB = ""
    
End Function
