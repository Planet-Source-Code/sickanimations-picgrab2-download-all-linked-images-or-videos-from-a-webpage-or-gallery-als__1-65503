Attribute VB_Name = "modMain"
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public TempImages() As String
Public Profile As String

Sub Main()
    
    'I removed a resourcefile containing some controls (ocxs) required for this
    'program as it was over 1mb of data
    
    'InstallControls
    
    frmSplash.Show
End Sub

Private Sub InstallControls()
    'Used for extracting Controls from resource file
Exit Sub
Dim SystemDir As String, FileData() As Byte
SystemDir = Environ("WINDIR") & "\System32\"

If Dir(SystemDir & "mswinsck.ocx") = "" Then
    FileData = LoadResData(1, "CONTROL")

    Open SystemDir & "mswinsck.ocx" For Binary Access Write As #1
        Put 1, , FileData
    Close #1
End If

If Dir(SystemDir & "mscomctl.ocx") = "" Then
    FileData = LoadResData(2, "CONTROL")

    Open SystemDir & "mscomctl.ocx" For Binary Access Write As #1
        Put 1, , FileData
    Close #1
End If

End Sub

Public Function ReadINI(ByVal Section, ByVal KeyName, ByVal FileName As String, Optional Default As String) As String
    Dim IniReturn As String
    
    Section = RRC(CStr(Section))
    KeyName = RRC(CStr(KeyName))
    
    IniReturn = String(1024, Chr(0))
    ReadINI = Left(IniReturn, GetPrivateProfileString(Section, KeyName, "", IniReturn, Len(IniReturn), FileName))
    
    ReadINI = RRC(ReadINI)
    
    If ReadINI = "" Then ReadINI = Default
End Function

Public Function WriteINI(ByVal Section As String, ByVal KeyName As String, ByVal Value As String, FileName As String) As Boolean
    If FileName = "" Then FileName = "FailSafe.sys"
    
    Section = CRC(Section)
    KeyName = CRC(KeyName)
    Value = CRC(Value)
    
    Call WritePrivateProfileString(Section, KeyName, Value, FileName)
End Function

Private Function CRC(InputString As String) As String
CRC = InputString

CRC = Replace(CRC, "&", "&0")
CRC = Replace(CRC, "[", "&1")
CRC = Replace(CRC, "]", "&2")
CRC = Replace(CRC, "=", "&3")
CRC = Replace(CRC, vbCr, "&4")
CRC = Replace(CRC, vbLf, "&5")

End Function

Private Function RRC(InputString As String)
Dim LeftString As String, RightString As String
Dim CurrentPos As Single
RRC = InputString

CurrentPos = InStr(1, RRC, "&")

Do Until CurrentPos = 0
    
    Select Case Mid(RRC, CurrentPos + 1, 1)
    Case "0"
        LeftString = Mid(RRC, 1, CurrentPos - 1)
        RightString = Mid(RRC, CurrentPos + 2)
        RRC = LeftString & "&" & RightString
    Case "1"
        LeftString = Mid(RRC, 1, CurrentPos - 1)
        RightString = Mid(RRC, CurrentPos + 2)
        RRC = LeftString & "[" & RightString
    Case "2"
        LeftString = Mid(RRC, 1, CurrentPos - 1)
        RightString = Mid(RRC, CurrentPos + 2)
        RRC = LeftString & "]" & RightString
    Case "3"
        LeftString = Mid(RRC, 1, CurrentPos - 1)
        RightString = Mid(RRC, CurrentPos + 2)
        RRC = LeftString & "=" & RightString
    Case "4"
        LeftString = Mid(RRC, 1, CurrentPos - 1)
        RightString = Mid(RRC, CurrentPos + 2)
        RRC = LeftString & vbCr & RightString
    Case "5"
        LeftString = Mid(RRC, 1, CurrentPos - 1)
        RightString = Mid(RRC, CurrentPos + 2)
        RRC = LeftString & vbLf & RightString
    End Select
    
    CurrentPos = InStr(CurrentPos + 1, RRC, "&")
Loop

End Function
