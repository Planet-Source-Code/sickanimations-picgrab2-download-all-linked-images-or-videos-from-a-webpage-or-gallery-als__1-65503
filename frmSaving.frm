VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSaving 
   BackColor       =   &H8000000A&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PicGrab - Downloading Pics"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6000
   Icon            =   "frmSaving.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkLog 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "Show Log"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3240
      TabIndex        =   12
      Top             =   2560
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   5520
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "Text Files (*.txt)|*.txt"
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pause"
      Height          =   255
      Left            =   4560
      TabIndex        =   11
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox txtStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1920
      Width           =   5760
   End
   Begin VB.Timer tmrActivity 
      Interval        =   25
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picProgressSRC 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   0
      Picture         =   "frmSaving.frx":1042
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   405
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.PictureBox picFileProgressSRC 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   0
      Picture         =   "frmSaving.frx":1DA1
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   413
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   4560
      TabIndex        =   6
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdSkip 
      Caption         =   "Skip"
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   2280
      Width           =   1335
   End
   Begin VB.PictureBox picFileProgress 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   202
      TabIndex        =   4
      Top             =   2280
      Width           =   3065
   End
   Begin VB.PictureBox picProgress 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   230
      Left            =   0
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   398
      TabIndex        =   3
      Top             =   3240
      Width           =   6000
   End
   Begin VB.PictureBox picActivitySRC 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   0
      Picture         =   "frmSaving.frx":2927
      ScaleHeight     =   315
      ScaleWidth      =   6315
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   6375
   End
   Begin VB.PictureBox picActivity 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   230
      Left            =   0
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   1
      Top             =   1500
      Width           =   6000
   End
   Begin VB.PictureBox picSaving 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1500
      Left            =   0
      Picture         =   "frmSaving.frx":2AF9
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   0
      Top             =   0
      Width           =   6000
   End
   Begin VB.Label lblTimeRemaining 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2640
      Width           =   3015
   End
   Begin VB.Label lblTotalProgress 
      BackStyle       =   0  'Transparent
      Caption         =   "00000 out of 00000 images."
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   2055
   End
End
Attribute VB_Name = "frmSaving"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Image1Left As Single, Image2Left As Single
Private CurrentDownload As String, CurrentFileLength As Long, CurrentArrayRef As Long
Private Images() As String

'Error detectors
Private ConnectionErrors As Long, Redirects As Long

'Variables for Statistics
Private TotalBytes As Long, TotalSaved As Long, TotalErrors As Long, TotalDownloadTime As Long, TotalImages As Long, CurrentImage As Long, CurrentDownloadStart As Long, TotalCompleted As Long

'Variables for loops
Private CancelSaving As Boolean, PauseSaving As Boolean

Private WithEvents HTTP As clsHttpClient
Attribute HTTP.VB_VarHelpID = -1

Public Function DownloadFiles(URLs() As String, NameCode As String, MinDigits As Single, Number As Single, DestDir As String, Optional StartFrom As Long, Optional EndAt As Long, Optional ContentType As String = "image")
On Error GoTo ErrorHandler
    Dim FileName As String
       
    Set HTTP = New clsHttpClient
    Images() = URLs()
    
    ReDim URLs(0 To 0)
    
    If Right(DestDir, 1) <> "\" Then DestDir = DestDir & "\"
    If Dir(DestDir, vbDirectory) = "" Then MakePath DestDir
    
    SaveStart = GetTickCount
    CancelSaving = False
    PauseSaving = False
    TotalBytes = 0
    TotalSaved = 0
    TotalErrors = 0
    TotalDownloadTime = 0
    TotalImages = 0
    CurrentImage = 0
    CurrentDownloadStart = 0

    If StartFrom < LBound(Images) Or StartFrom > UBound(Images) Then StartFrom = LBound(Images)
    If EndAt = 0 Or EndAt > UBound(Images) Or EndAt < StartFrom Then EndAt = UBound(Images)
    
    TotalImages = (EndAt - StartFrom) + 1
    
    CurrentImage = 0
    
    lblTotalProgress.Caption = CurrentImage & " out of " & TotalImages & " files."
    DrawProgress CSng(TotalImages), CSng(CurrentImage), picProgress, picProgressSRC
    
    frmLog.LogEvent "Downloading " & TotalImages & " files. Started at " & Time, 6
    For CurrentArrayRef = StartFrom To EndAt
        
        CurrentImage = CurrentImage + 1
        
        If PauseSaving = True Then frmLog.LogEvent "Paused file download process (Image " & CurrentImage & " of " & TotalImages & " at " & Time, 7
        
        Do While PauseSaving = True
            txtStatus.Text = "Downloading of pictures is paused."
            DoEvents
        Loop
        
        FileName = GetFileName(Images(CurrentArrayRef))
        FileName = CodeToName(FileName, NameCode, MinDigits, Number)
        
        
        If CancelSaving = True Then GoTo Cancel
        
        txtStatus.Text = "Connecting to host."
        frmLog.LogEvent "Connecting: '" & Images(CurrentArrayRef) & "' (File " & CurrentImage & " out of " & TotalImages & ")", 1
        
        CurrentDownloadStart = GetTickCount
        
        If HTTP.DownloadFile(Images(CurrentArrayRef), , , True, DestDir & FileName, ContentType) = 0 Then
            Number = Number + 1
            TotalCompleted = TotalCompleted + 1
        Else
            TotalErrors = TotalErrors + 1
        End If
        
        If CancelSaving = True Then GoTo Cancel
               
    lblTotalProgress.Caption = CurrentImage & " out of " & TotalImages & " files."
    DrawProgress CSng(TotalImages), CSng(CurrentImage), picProgress, picProgressSRC
    
    Next CurrentArrayRef
    
    Set HTTP = Nothing
    
    cmdSkip.Enabled = False
    cmdCancel.Caption = "Exit"
    
    On Error Resume Next
Cancel:
    
    On Error Resume Next
    WriteINI "Settings", "StartFrom", Number, Profile
    
    txtStatus.Text = TotalCompleted & " files downloaded successfuly. " & TotalErrors & " files not saved. Downloaded " & TotalSaved & "kb in " & Int((TotalDownloadTime) / 1000) & ". Average transfer speed: " & Round(TotalSaved / ((TotalDownloadTime) / 1000), 2) & "kb/s."
    frmLog.LogEvent "Downloaded " & TotalCompleted & " of " & TotalImages & " files. Downloaded " & TotalSaved & "kb in " & Int((TotalDownloadTime) / 1000) & ". Average transfer speed: " & Round(TotalSaved / ((TotalDownloadTime) / 1000), 2) & "kb/s. Process completed at: " & Time, 8
    
    tmrActivity.Enabled = False
    cmdSkip.Enabled = False
    cmdPause.Enabled = False
    cmdCancel.Caption = "Exit"
    frmLog.Hide
    
    Exit Function

ErrorHandler:
    frmLog.LogEvent "Error #" & Err.Number & " in frmSaving.DownloadFiles(): " & Err.Description, 3
    Resume Next
End Function

Private Sub chkLog_Click()
    Select Case CBool(chkLog.Value)
    Case True
        frmLog.Show , Me
    Case False
        frmLog.Hide
    End Select
End Sub

Private Sub cmdCancel_Click()
    On Error Resume Next
    Select Case cmdCancel.Caption
    Case "Cancel"
        On Error GoTo ErrorHandler
        
        If MsgBox("Are you sure you want to cancel?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        If MsgBox("Do you want to save the URLS of remaining files to a list so you can continue saving these pictures?", vbQuestion + vbYesNo) = vbYes Then
            Dim FileName As String, FileNumber As String
            
            cd.ShowSave
            
            FileName = cd.FileName
            FileNumber = FreeFile
            
            If Dir(FileName) <> "" Then Kill FileName
            Open FileName For Binary Access Write As FileNumber
                For i = CurrentArrayRef To UBound(Images)
                    Put FileNumber, , Images(i) & vbLf
                Next i
            Close FileNumber
        End If
        
        HTTP.CancelOperations
        CancelSaving = True
        PauseSaving = False
        cmdPause.Enabled = False
        cmdCancel.Caption = "Exit"
    Case "Exit"
        Unload Me
    End Select
Exit Sub
ErrorHandler:
    MsgBox "Image List was not saved succesfully." & vbCrLf & "(Error #" & Err.Number & ": " & Err.Description & ")", vbExclamation + vbOKOnly
End Sub

Private Sub cmdPause_Click()
    On Error Resume Next
    Select Case cmdPause.Caption
    Case "Pause"
        PauseSaving = True
        cmdPause.Caption = "Unpause"
    Case "Unpause"
        PauseSaving = False
        cmdPause.Caption = "Pause"
        frmLog.LogEvent "Unpaused image download process at " & Time, 6
    End Select
End Sub

Private Sub cmdSkip_Click()
    On Error Resume Next
    HTTP.CancelOperations
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Image1Left = 200
    Image2Left = -200
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    
    If CurrentArrayRef < TotalImages And CancelSaving = False Then
        If MsgBox("It is not recommended closing this window unless all the pictures are saved or you have cancelled the process." & vbCrLf & "Do you want to close it anyway?", vbYesNo + vbExclamation) = vbNo Then Cancel = 1: Exit Sub
    End If
    
    frmLog.Hide
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PauseSaving = False
    CancelSaving = True
    frmMain.Show
End Sub

Private Sub HTTP_DownloadBegin(Location As String, FileLength As Long)
    On Error Resume Next
    Redirects = 0
    ConnectionErrors = 0
    
    CurrentDownload = Location
    CurrentFileLength = FileLength
    txtStatus.Text = "Download started (" & Round(FileLength / 1024, 2) & "kb)."
    
    frmLog.LogEvent "Download of '" & Location & "' has begun.", 5

    DrawProgress 1, 0, picFileProgress, picFileProgressSRC
End Sub

Private Sub HTTP_DownloadComplete(URL As String, TimeMs As Long, BytesDownloaded As Long)
    On Error Resume Next
    DrawProgress 1, 1, picFileProgress, picFileProgressSRC
    
    txtStatus.Text = "Download completed."
    frmLog.LogEvent "Download complete: '" & URL & "' (" & Round(BytesDownloaded / 1024) & "kb) in " & Round(TimeMs / 1000) & " seconds. Average rate of " & Round((BytesDownloaded / 1024) / (TimeMs / 1000), 2) & "kb/s.", 2

    TotalDownloadTime = TotalDownloadTime + (GetTickCount - CurrentDownloadStart)
    TotalSaved = TotalSaved + (BytesDownloaded / 1024)
    
    Dim Average As Long
    Average = Round((TotalDownloadTime / 1000) / CurrentImage)
    Average = Average * ((TotalImages) - CurrentImage)
    lblTimeRemaining.Caption = "Estimated time remaining: " & Format(DateAdd("s", Average, 0), "hh:mm:ss")
    
End Sub

Private Sub HTTP_DownloadError(Number As Integer, Description As String, URL As String)
    On Error Resume Next
    frmLog.LogEvent "Error #:" & Number & ": '" & Description & "'", 3
    If Number = 0 Or Number = 1 Then ConnectionErrors = ConnectionErrors + 1
    
    TotalDownloadTime = TotalDownloadTime + (GetTickCount - CurrentDownloadStart)
    
    If ConnectionErrors >= 50 Then
        'There have been 50 connection errors in a row. The internet is probably disconnected
        cmdPause_Click
        
        If MsgBox("There have been 50 connection errors in a row." & vbCrLf & "A possible cause for this is the Internet Connection being terminated." & vbCrLf & "The download proocess has been paused, do you want PicGrab to go back 50 files?", vbYesNo + vbQuestion) = vbYes Then
            CurrentArrayRef = CurrentArrayRef - 50
        End If
            ConnectionErrors = 0
    End If
    
End Sub

Private Sub HTTP_DownloadProgress(Downloaded As Long, Total As Long, Percent As Single)
    On Error Resume Next
    txtStatus.Text = Round(Percent, 1) & "% (" & Int(Downloaded / 1024) & "kb of " & Int(Total / 1024) & "kb)."
    DrawProgress CSng(Total), CSng(Downloaded), picFileProgress, picFileProgressSRC
End Sub


Private Sub HTTP_DownloadRedirect(OldUrl As String, NewUrl As String)
    frmLog.LogEvent "Redirected to '" & NewUrl & "'", 4
    Redirects = Redirects + 1
    If Redirects >= 10 Then HTTP.CancelOperations: Redirects = 0
End Sub

Private Sub tmrActivity_Timer()
If PauseSaving = True Then Exit Sub
If CancelSaving = True Then Exit Sub

    If Image1Left >= 400 Then Image1Left = -400
    If Image2Left >= 400 Then Image2Left = -400
    
    picActivity.Cls
    
    BitBlt picActivity.hDC, Image1Left, 0, 400, 15, picActivitySRC.hDC, 0, 0, vbSrcCopy
    BitBlt picActivity.hDC, Image2Left, 0, 400, 15, picActivitySRC.hDC, 0, 0, vbSrcCopy
    
    picActivity.Refresh
    
    Image1Left = Image1Left + 5
    Image2Left = Image2Left + 5
    
    DoEvents
End Sub

Private Function DrawProgress(Max As Single, Value As Single, picDestination As PictureBox, picSource As PictureBox)
Dim ImageWidth As Single, ImageHeight As Single

ImageWidth = Progress(Max, Value, picDestination.ScaleWidth)
ImageHeight = picDestination.ScaleHeight

picDestination.Cls
BitBlt picDestination.hDC, 0, 0, ImageWidth, ImageHeight, picSource.hDC, 0, 0, vbSrcCopy
picDestination.Refresh

End Function

Private Function Progress(Maximum As Single, Value As Single, MaxWidth As Single) As Single
Dim Percentage As Single, Width As Single

If Maximum <= 0 Or MaxWidth <= 0 Then
    Progress = 0
    Exit Function
End If

If Value >= Maximum Then
    Progress = MaxWidth
    Exit Function
End If


Percentage = (Value * 100) / Maximum

SetProgress:

Progress = (Percentage / 100) * MaxWidth

End Function

Private Function GetFileName(ByVal URL As String) As String
    Dim temp() As String
    If Left(URL, Len("http://")) = "http://" Then URL = Mid(URL, Len("http://") + 1)
    If InStr(1, URL, "/") = 0 Then URL = URL & "/"
    
    temp() = Split(URL, "/")
    GetFileName = temp(UBound(temp))
    
    
End Function

Private Function MakePath(Path As String)
    Dim temp() As String, CurrentDir As String
        
        On Error Resume Next
        
        If InStr(1, Path, "\") = 0 Then Exit Function
        temp() = Split(Path, "\")
        CurrentDir = temp(0) & "\"
        
    For i = 1 To UBound(temp)
        CurrentDir = CurrentDir & temp(i) & "\"
        If Dir(CurrentDir, vbDirectory) = "" Then
            MkDir CurrentDir
        End If
    Next i
    
End Function
