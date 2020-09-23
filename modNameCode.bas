Attribute VB_Name = "modNameCode"
Public Function CodeToName(ByVal FileName As String, ByVal NameCode As String, ByVal MinDigits As Single, ByVal Number As Single) As String
    Dim FileTitle As String, FileExtention As String, FileNumber As String
    
    FileTitle = GetFileTitle(FileName)
    FileExtention = GetFileExtention(FileName)
    FileNumber = GetNumber(MinDigits, Number)
    
    NameCode = Replace(NameCode, "/title/", FileTitle)
    NameCode = Replace(NameCode, "/extention/", FileExtention)
    NameCode = Replace(NameCode, "/number/", FileNumber)
    
    CodeToName = NameCode
    
End Function


Private Function GetFileTitle(FileName As String) As String
    Dim TempString As String, LastPosition As Single, OldLastPosition As Single
    
    TempString = FileName
    
    If InStr(1, TempString, ".") = 0 Then GetFileTitle = FileName: Exit Function
    
CheckNext:
    
    OldLastPosition = LastPosition
    LastPosition = LastPosition + InStr(1, TempString, ".")
    TempString = Mid(TempString, LastPosition + 1)
    
    If LastPosition <> OldLastPosition Then GoTo CheckNext
    
    GetFileTitle = Mid(FileName, 1, LastPosition - 1)
    
End Function


Private Function GetFileExtention(FileName As String) As String
    Dim temp() As String
    
    temp() = Split(FileName, ".")
    
    If UBound(temp) < 1 Then GetFileExtention = "": Exit Function
    
    GetFileExtention = temp(UBound(temp))
    
End Function


Private Function GetNumber(Digits As Single, Number As Single) As String
    Dim NumberStr As String
    
    NumberStr = CStr(Val(Number))
    
FillPlaceHolder:
    
    If Len(NumberStr) < Digits Then
        NumberStr = "0" & NumberStr
        GoTo FillPlaceHolder
    End If
    
    GetNumber = NumberStr
End Function
