VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHttpClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class Module written by Tim Cinel 2003
'  email: sickanimations@hotmail.com
'website: http://www.sickanimations.cjb.net/

'Steal this code - i dont care!

'If you want to use this class module, you will -need- to do the following:

'Add the reference '%SystemRoot%\System32\mswinsock.ocx'
'(Click: Project, References, Browse..., Navigate to System32 in your windows
'directory, type 'mswinsock.ocx', click OK)
'If you have the winsock control as a component, I -think- you have to remove it.

'Declare the Class Module in your code like this:
'<Scope> WithEvents <Name> as clsHttpClient
'Example:
'Public WithEvents HttpDownloader as clsHttpClient

'You like my code? Email me! Find more of it at http://www.planetsourcecode.com/

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private WithEvents Socket As MSWinsockLib.Winsock
Attribute Socket.VB_VarHelpID = -1

Private Cancel As Boolean
Private Busy As Boolean

Const CrLf As String = vbCrLf

Private DataBuffer As String

Event DownloadBegin(Location As String, FileLength As Long)
Event DownloadProgress(Downloaded As Long, Total As Long, Percent As Single)
Event DownloadComplete(URL As String, TimeMs As Long, BytesDownloaded As Long)
Event DownloadError(Number As Integer, Description As String, URL As String)
Event DownloadRedirect(OldUrl As String, NewUrl As String)

Function GetFileHeader(URL As String, Optional Port As Long = 80, Optional ByRef ReturnCode As String, Optional ByRef ContentLength As Long, Optional ByRef ContentType As String, Optional ByRef Location As String)
    On Error Resume Next
    Dim Host As String, File As String, ClientHeader As String
    
    Cancel = False
    
    Host = GetHostName(URL)
    File = GetFileName(URL)
    
    ClientHeader = MakeClientHeader(Host, File, "HEAD")
       
    Socket.Close
    Socket.Connect Host, Port
    
Dim StartTime As Long
    
    StartTime = GetTickCount
    Do Until Socket.State = sckConnected
        If Cancel = True Then RaiseEvent DownloadError(9, "The process was cancelled.", URL): GoTo Error
        
        If Socket.State = 8 Or Socket.State = 9 Or Socket.State = 0 Then Exit Function
        If GetTickCount - StartTime > 10000 Then Exit Function
        DoEvents
    Loop
     
       
    DataBuffer = ""
    Socket.SendData ClientHeader
    
    StartTime = GetTickCount
    
    Do While DataBuffer = ""
        If Cancel = True Then RaiseEvent DownloadError(9, "The process was cancelled.", URL): GoTo Error
        
        If GetTickCount - StartTime > 15000 Then RaiseEvent DownloadError(2, "Timed out waiting for data to arrive.", URL): Exit Function
        DoEvents
    Loop
    
    Socket.Close
    
    Dim Lines() As String, CurrentLine() As String
    
    If InStr(1, DataBuffer, vbCrLf) = 0 Then Exit Function
    
    Lines() = Split(DataBuffer, vbCrLf)
    
    For i = 0 To UBound(Lines)
        CurrentLine() = Split(Lines(i), " ", 2)
        If Left(CurrentLine(0), Len("HTTP")) = "HTTP" Then
            ReturnCode = CurrentLine(1)
        ElseIf Left(CurrentLine(0), Len("Content-Length:")) = "Content-Length:" Then
            ContentLength = Val(CurrentLine(1))
        ElseIf Left(CurrentLine(0), Len("Content-Type:")) = "Content-Type:" Then
            ContentType = CurrentLine(1)
        ElseIf Left(CurrentLine(0), Len("Location:")) = "Location:" Then
            Location = CurrentLine(1)
        End If
        
    Next i
    
    GetFileHeader = 1
    Busy = False
    
    Exit Function
Error:

    Socket.Close
    Busy = False
End Function

Function DownloadFile(ByRef URL As String, Optional Port As Long = 80, Optional ByRef DownloadData As String, Optional SaveToFile As Boolean = False, Optional Destination As String, Optional ContentAllowed As String = "*") As Long
    On Error GoTo Error
    Dim Host As String, File As String, ClientHeader As String
    Dim StartTime As Long, DownloadStart As Long
    Dim Data As String, temp As String
    
    'If Busy = True Then Exit Function
    
    Busy = True
    Cancel = False

Begin:
    Host = GetHostName(URL)
    File = GetFileName(URL)
    
    ClientHeader = MakeClientHeader(Host, File, "GET")
    
    Socket.Close
    Socket.Connect Host, Port
    
    StartTime = GetTickCount
    Do Until Socket.State = sckConnected
        If Cancel = True Then DownloadFile = 9:  RaiseEvent DownloadError(9, "The process was cancelled.", URL): GoTo Error
        
        If Socket.State = 8 Or Socket.State = 9 Or Socket.State = 0 Then RaiseEvent DownloadError(0, "Error connecting to host.", URL): GoTo Error
        If GetTickCount - StartTime > 20000 Then DownloadFile = 1: RaiseEvent DownloadError(1, "Connection timeout.", URL): GoTo Error
        DoEvents
    Loop
    
    StartTime = GetTickCount
       
    DataBuffer = ""
    Socket.SendData ClientHeader
    
    Do While DataBuffer = ""
        If Cancel = True Then DownloadFile = 9:  RaiseEvent DownloadError(9, "The process was cancelled.", URL): GoTo Error
        
        If GetTickCount - StartTime > 20000 Then DownloadFile = 2: RaiseEvent DownloadError(2, "Timed out waiting for data to arrive.", URL): GoTo Error
        DoEvents
    Loop
    
    'Debug.Print DataBuffer
    
    Dim ResponseEnd As Long, FileLen As Long, Downloaded As Long, Saved As Long
    ResponseEnd = InStr(1, DataBuffer, CrLf & CrLf) + Len(CrLf & CrLf)
      
    If ResponseEnd > 0 Then
        Dim Response As String, HttpCode As Long, ContentType As String, Allowed() As String
        
        
        Response = Mid(DataBuffer, 1, ResponseEnd)
        'Debug.Print Response
        DataBuffer = Mid(DataBuffer, ResponseEnd)
        'Debug.Print Data
        
        HttpCode = Val(QueryResponseData(Response, 0))
        Select Case Int(HttpCode / 100)
        Case 2
            'OK!
            FileLen = Val(QueryResponseData(Response, 1))
            ContentType = QueryResponseData(Response, 2)
            
            'Check if the content type is allowed
            If ContentAllowed <> "*" Then
                If InStr(1, ContentAllowed, ",") = 0 Then
                    If InStr(1, LCase(ContentType), LCase(ContentAllowed)) = 0 Then DownloadFile = 10: RaiseEvent DownloadError(10, "Content type not allowed. (" & ContentType & ")", URL): GoTo Error
                Else
                    Allowed() = Split(ContentAllowed, ",")
                    For i = 0 To UBound(Allowed)
                        If InStr(1, LCase(ContentType), LCase(Allowed(i))) = 0 Then DownloadFile = 10: RaiseEvent DownloadError(10, "Content type not allowed. (" & ContentType & ")", URL): GoTo Error
                    Next i
                End If
            End If
        Case 3
            '(redirection)
            temp = QueryResponseData(Response, 3)
            temp = ExtendLink(URL, temp)
            RaiseEvent DownloadRedirect(URL, temp)
            URL = temp
            GoTo Begin
        Case 4
            '(client error) - probably file not found
            RaiseEvent DownloadError(4, "Server returned error " & HttpCode & ".", URL)
            DownloadFile = 4
            GoTo Error
        Case 5
            'Server has error OR is incapible of satisfing request
            RaiseEvent DownloadError(5, "Server error " & HttpCode & ".", URL)
            DownloadFile = 5
            GoTo Error
        End Select
    Else
        
    End If
    
    If SaveToFile = True Then
        Dim FileNumber As Integer
        FileNumber = FreeFile
        If Dir(Destination) <> "" Then Kill Destination
        Open Destination For Binary Access Write As FileNumber
        Put FileNumber, , DataBuffer
        Saved = Saved + Len(DataBuffer)
        DataBuffer = ""
    Else
        DownloadData = DownloadData & DataBuffer
        DataBuffer = ""
    End If
    
    
    RaiseEvent DownloadBegin(URL, FileLen)
    DownloadStart = GetTickCount
    
    StartTime = GetTickCount
    Do Until Socket.State <> sckConnected
        If Cancel = True Then DownloadFile = 9:  RaiseEvent DownloadError(9, "The process was cancelled.", URL): GoTo Error
        If GetTickCount - StartTime > 60000 Then DownloadFile = 11: RaiseEvent DownloadError(11, "Timed out waiting for data.", URL)
        Downloaded = Downloaded + Len(DataBuffer)
        
        If Downloaded > Saved Then
            'Data has been received
            If SaveToFile = True Then
                Put FileNumber, , DataBuffer
            Else
                DownloadData = DownloadData & DataBuffer
            End If
            Saved = Saved + Len(DataBuffer)
            DataBuffer = ""
            
            If FileLen > 0 Then
                RaiseEvent DownloadProgress(Downloaded, FileLen, (100 / FileLen) * Downloaded)
            Else
                RaiseEvent DownloadProgress(Downloaded, FileLen, 0)
            End If
            
            StartTime = GetTickCount
        End If
        
        If Saved >= FileLen And FileLen <> 0 Then
            Socket.Close
        End If
        
        DoEvents
    Loop
    
    DoEvents
    
    If SaveToFile = True Then
        If DataBuffer <> "" Then Put FileNumber, , DataBuffer
        Close FileNumber
    Else
        If DataBuffer <> "" Then DownloadData = DownloadData & DataBuffer
    End If
    
    Busy = False
    DownloadFile = 0
    RaiseEvent DownloadComplete(URL, GetTickCount - DownloadStart, Downloaded)
    
    Exit Function

Error:
    On Error Resume Next
       
    Socket.Close
    
    Close FileNumber
    If Destination <> "" Then
        If Dir(Destination) <> "" Then Kill Destination
    End If
    
    Busy = False
    
    If Err.Number > 0 Then
        RaiseEvent DownloadError(Err.Number, Err.Description, URL)
    End If
End Function

Function CancelOperations()
    Cancel = True
End Function

Private Function GetHostName(ByVal URL As String) As String
    On Error Resume Next
    If Left(URL, Len("http://")) = "http://" Then URL = Mid(URL, Len("http://") + 1)
    If InStr(1, URL, "/") = 0 Then URL = URL & "/"
    
    GetHostName = Split(URL, "/")(0)
    
End Function

Private Function GetFileName(ByVal URL As String) As String
    On Error Resume Next
    If Left(URL, Len("http://")) = "http://" Then URL = Mid(URL, Len("http://") + 1)
    If InStr(1, URL, "/") = 0 Then URL = URL & "/"
    
    GetFileName = Split(URL, "/", 2)(1)
    
End Function

Private Function MakeClientHeader(Host As String, File As String, Request As String)
    On Error Resume Next
    MakeClientHeader = Request & " /" & File & " HTTP/1.1" & CrLf & _
                       "Accept: */*" & CrLf & _
                       "Connection: close" & CrLf & _
                       "User-Agent: Mozilla/5.0 Gecko/20030624" & CrLf & _
                       "Host: " & Host & CrLf & CrLf

End Function

Private Function QueryResponseData(Response As String, Output As Long) As String
    On Error Resume Next
    Dim Lines() As String, CurrentLine() As String
    
    'OUTPUT
    ' 0 - HTTP Response Code
    ' 1 - Content Length
    ' 2 - Content Type
    ' 3 - Location
    
    If InStr(1, Response, vbCrLf) = 0 Then Exit Function
    
    Lines() = Split(Response, vbCrLf)
    
    For i = 0 To UBound(Lines)
        CurrentLine() = Split(Lines(i), " ", 2)
        If LCase(Left(CurrentLine(0), Len("HTTP"))) = "http" And Output = 0 Then
            QueryResponseData = CurrentLine(1)
            Exit Function
        ElseIf LCase(Left(CurrentLine(0), Len("Content-Length:"))) = "content-length:" And Output = 1 Then
            QueryResponseData = Val(CurrentLine(1))
            Exit Function
        ElseIf LCase(Left(CurrentLine(0), Len("Content-Type:"))) = "content-type:" And Output = 2 Then
            QueryResponseData = CurrentLine(1)
            Exit Function
        ElseIf LCase(Left(CurrentLine(0), Len("Location:"))) = "location:" And Output = 3 Then
            QueryResponseData = CurrentLine(1)
            Exit Function
        End If
        
    Next i

End Function

Private Function ExtendLink(ByVal URL As String, ByVal Link As String) As String
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

Private Sub Class_Initialize()
    On Error Resume Next
    Set Socket = New MSWinsockLib.Winsock
End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    Dim temp As String
    
    Socket.GetData temp
    DataBuffer = DataBuffer & temp
    
End Sub
