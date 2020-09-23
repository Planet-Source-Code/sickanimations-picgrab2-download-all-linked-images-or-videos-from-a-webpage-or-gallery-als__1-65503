Attribute VB_Name = "modHTML"
'Module written by Tim Cinel 2003
'  email: sickanimations@hotmail.com
'website: http://www.sickanimations.cjb.net/

'Steal this code - i dont care!

'You like my code? Email me! Find more of it at http://www.planetsourcecode.com/

Public Type HttpFileHeader
    Content_Length As Long
    Content_Type As String
    Content_Location As String
    HTTP_Proto As String
End Type

Public Type Link
    HREF As String
    face As String
    target As String
End Type

Function HTTP_Protocol(Protocol As String, Output As Long) As String
    On Error Resume Next
   
    Select Case Output
    Case 0  '0 HTTP VERSION
       If CharCount(Protocol, " ") < 2 Then Exit Function
       HTTP_Protocol = Split(Protocol, " ")(0)
    Case 1  '1 HTTP RETURN NUMBER
       If CharCount(Protocol, " ") < 2 Then Exit Function
       HTTP_Protocol = Split(Protocol, " ")(1)
    Case 2  '2 HTTP RETURN STRING
       If CharCount(Protocol, " ") < 2 Then Exit Function
       HTTP_Protocol = Split(Protocol, " ")(2)
    End Select
    
End Function

Private Function AnalyseLink(ByVal LinkTag As String) As Link
    On Error Resume Next
    Dim StartPos As Long, EndPos As Long
    
    'Get the link face
    StartPos = InStr(1, LinkTag, ">")
    EndPos = InStr(StartPos, LCase(LinkTag), "</a")
    
    If EndPos > StartPos Then
        StartPos = StartPos + Len(">")
        AnalyseLink.face = Mid(LinkTag, StartPos, EndPos - StartPos)
    End If
    
    'Remove spaces and make sure all quotes are the same ( " can be used as well as ' in html )
    'This makes things much easier
    LinkTag = Replace(LinkTag, " ", "")
    LinkTag = Replace(LinkTag, Chr(34), "'")
    
    'Get the link reference
    
    StartPos = InStr(1, LCase(LinkTag), "href='")
    EndPos = InStr(StartPos + Len("href='"), LinkTag, "'")
    
    If EndPos > StartPos Then
        StartPos = StartPos + Len("href='")
        AnalyseLink.HREF = Mid(LinkTag, StartPos, EndPos - StartPos)
    End If
    
    'Get the link target
    
    StartPos = InStr(1, LCase(LinkTag), "target='")
    EndPos = InStr(StartPos + Len("target='"), LinkTag, "'")
    
    If EndPos > StartPos Then
        StartPos = StartPos + Len("target='")
        AnalyseLink.target = Mid(LinkTag, StartPos, EndPos - StartPos)
    End If
    
End Function

Function ReadDocument(ByVal Document As String, URL As String, ByRef Title As String, ByRef Links() As Link)
    On Error Resume Next
    Dim StartPos As Long, EndPos As Long, CurrentLine As String, CurrentLink As Link
    Dim FileDirectory As String
    
    'Get the directory of the HTML document (Used when extending links)
    FileDirectory = GetDirectory(URL)
    
    'Find the title
    StartPos = InStr(1, LCase(Document), "<title>")
    EndPos = InStr(1, LCase(Document), "</title>")
    
    'In case there is no title
    If EndPos > StartPos Then
        Title = Mid(Document, StartPos + Len("<title>"), (EndPos - (StartPos + Len("<title>"))))
    Else
        Title = ""
    End If
    'Debug.Print Document
    
    'Dim the array
    ReDim Links(0 To 0)
    
    'Find the start first link
    StartPos = InStr(1, LCase(Document), "<a")
    Do Until StartPos = 0
        'Find the end of the link just found
        EndPos = InStr(StartPos, LCase(Document), "</a>")
            
            'Incase the end of the link wasn't found
            If EndPos > StartPos Then
                EndPos = EndPos + Len("</a>")
                
                'Put the whole tag into a string
                CurrentLine = Mid(Document, StartPos, EndPos - StartPos)
                'Analyse the tag to find the HREF, Face code, and the target.
                CurrentLink = AnalyseLink(CurrentLine)
                
                CurrentLink.HREF = ExtendLink(FileDirectory, CurrentLink.HREF)
                
                If CurrentLink.HREF = "" Then GoTo NextLink
                'Add the link to the Array
                If UBound(Links) = 0 Then
                    ReDim Links(1 To 1)
                    Links(1) = CurrentLink
                Else
                    ReDim Preserve Links(1 To UBound(Links) + 1)
                    Links(UBound(Links)) = CurrentLink
                End If
            
            End If
        'Find start of the next link
NextLink:
        StartPos = InStr(StartPos + 1, LCase(Document), "<a")
    Loop
    'Debug.Print Document
    
End Function

Function ExtendLink(ByVal URL As String, ByVal Link As String) As String
On Error Resume Next
Dim HostName As String

HostName = GetHostName(URL)

CheckLink:
If Left(Link, Len("../")) = "../" Then
    'This refers to a parent directory
    Link = Mid(Link, Len("../*"))
    URL = GetParent(URL)
    GoTo CheckLink
    
ElseIf Left(LCase(Link), Len("/")) = "/" Then
    'Refers to a root directory
    ExtendLink = HostName & Link
ElseIf Left(Link, Len("javascript:")) = "javascript:" Then
    'This is a javascript link
    ExtendLink = ""
ElseIf Left(Link, Len("#")) = "#" Then
    'This is a link to an anchor
    ExtendLink = ""
ElseIf Left(Link, Len("mailto:")) = "mailto:" Then
    'This is a email link
    ExtendLink = ""
ElseIf Left(LCase(Link), Len("http://")) = "http://" Then
    'This is an absolute link
    ExtendLink = Link
Else
    'It is a relative link
    ExtendLink = URL & Link
End If

End Function

Public Function CharCount(InputStr As String, Character As String) As Long
    On Error Resume Next
    Dim CurrentPos As Long
    On Error GoTo ErrorHandler
    Character = Mid(Character, 1, 1)
    
    CurrentPos = InStr(1, InputStr, Character)
    Do Until CurrentPos = 0
        CharCount = CharCount + 1
        CurrentPos = InStr(CurrentPos, InputStr, Character)
    Loop
    
ErrorHandler:
End Function

Private Function GetHostName(ByVal URL As String) As String
    On Error Resume Next
    If Left(URL, Len("http://")) = "http://" Then URL = Mid(URL, Len("http://") + 1)
    If InStr(1, URL, "/") = 0 Then URL = URL & "/"
    
    GetHostName = Split(URL, "/")(0)
    
End Function

Function GetDirectory(ByVal URL As String) As String
    On Error Resume Next
    If Left(URL, Len("http://")) <> "http://" Then URL = "http://" & URL
    If InStr(Len("http://*"), URL, "/") = 0 Then URL = URL & "/"
    
    Dim temp() As String
    temp() = Split(URL, "/")
    
    For i = 0 To UBound(temp) - 1
        GetDirectory = GetDirectory & temp(i) & "/"
    Next i
End Function

Function GetParent(ByVal URL As String) As String
    On Error Resume Next
    Dim temp() As String
    temp() = Split(URL, "/")
    
    If UBound(temp) < 1 Then GetParent = URL: Exit Function
    
    For i = 0 To UBound(temp) - 2
        GetParent = GetParent & temp(i) & "/"
    Next i
End Function
