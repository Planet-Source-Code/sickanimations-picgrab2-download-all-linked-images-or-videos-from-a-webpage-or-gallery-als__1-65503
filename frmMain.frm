VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PicGrab2"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtImagesFound 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   6000
      Width           =   615
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   255
      Left            =   4800
      TabIndex        =   28
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   255
      Left            =   3600
      TabIndex        =   27
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Frame fmeSettings 
      Height          =   2055
      Left            =   120
      TabIndex        =   5
      Top             =   6240
      Width           =   5895
      Begin VB.TextBox txtContentType 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4440
         TabIndex        =   35
         Text            =   "image"
         ToolTipText     =   $"frmMain.frx":038A
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtFileTypes 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   34
         Text            =   "jpg,png,jpeg"
         ToolTipText     =   $"frmMain.frx":0412
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CheckBox chkContentType 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Content Types: "
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3000
         TabIndex        =   33
         ToolTipText     =   "Enable extention control"
         Top             =   1680
         Width           =   1815
      End
      Begin VB.CheckBox chkFileTypes 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "File Types:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   32
         ToolTipText     =   "Enable extention control"
         Top             =   1680
         Width           =   1815
      End
      Begin VB.CheckBox chkViewLog 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "View HTTP Log"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3960
         TabIndex        =   31
         Top             =   1050
         Width           =   1695
      End
      Begin VB.CommandButton cmdLoadProfile 
         Caption         =   "Load Profile"
         Height          =   255
         Left            =   4200
         TabIndex        =   25
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdSaveProfile 
         Caption         =   "Save Profile"
         Height          =   255
         Left            =   3000
         TabIndex        =   24
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chkDisallow 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Disallow extentions:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   23
         ToolTipText     =   "Enable extention control"
         Top             =   1370
         Width           =   1815
      End
      Begin VB.TextBox txtDisallow 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         TabIndex        =   22
         Text            =   "com/"
         ToolTipText     =   $"frmMain.frx":049A
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdGrabAndGo 
         Caption         =   "Grab and Go"
         Height          =   255
         Left            =   480
         TabIndex        =   21
         ToolTipText     =   "Specify saving details before grabbing, this way you can leave the computer unnatended if it will be a big grab."
         Top             =   720
         Width           =   2415
      End
      Begin VB.CommandButton cmdLoadImageList 
         Caption         =   "Load Image List"
         Height          =   255
         Left            =   3000
         TabIndex        =   16
         ToolTipText     =   "Load an image list saved previously for Downloading."
         Top             =   480
         Width           =   2415
      End
      Begin VB.CommandButton cmdSaveImages 
         Caption         =   "Save Images"
         Height          =   255
         Left            =   480
         TabIndex        =   9
         ToolTipText     =   "Save all images in the list above."
         Top             =   480
         Width           =   2415
      End
      Begin VB.CommandButton cmdSaveImageList 
         Caption         =   "Save Image List"
         Height          =   255
         Left            =   3000
         TabIndex        =   17
         ToolTipText     =   "Save all images in the list above into a file that you can load back into this list."
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton cmdStartGrabbing 
         Caption         =   "Start Grabbing"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         ToolTipText     =   "Start crawling through the links looking for pictures."
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox txtExtention 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         TabIndex        =   15
         Text            =   "htm,html,/"
         ToolTipText     =   $"frmMain.frx":0522
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox chkExtention 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Allow extentions:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   14
         ToolTipText     =   "Enable extention control"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtLinkDepth 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5520
         MaxLength       =   1
         TabIndex        =   7
         Text            =   "2"
         ToolTipText     =   "Maximum amount of parent webpages a webpage can have."
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Maximum Link Depth:"
         Height          =   255
         Left            =   3960
         TabIndex        =   6
         Top             =   1320
         Width           =   1575
      End
   End
   Begin MSComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   8325
      Width           =   6150
      _ExtentX        =   10848
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8229
            Text            =   "SickAnimations PicGrab2 Pro - Tim Cinel 2003"
            TextSave        =   "SickAnimations PicGrab2 Pro - Tim Cinel 2003"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Ready"
            TextSave        =   "Ready"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "Text Files (*.txt)|*.txt"
   End
   Begin MSComctlLib.ImageList il 
      Left            =   5520
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":05B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0816
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0AA3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fmeURL 
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   5895
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add URL"
         Height          =   255
         Left            =   4680
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtURL 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4455
      End
   End
   Begin MSComctlLib.TreeView tvLinks 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   8916
      _Version        =   393217
      Indentation     =   9
      LineStyle       =   1
      PathSeparator   =   "/|¤|\"
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "il"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Frame fmeGrabbing 
      Height          =   1215
      Left            =   120
      TabIndex        =   10
      Top             =   6240
      Visible         =   0   'False
      Width           =   5895
      Begin VB.CommandButton cmdSkip 
         Caption         =   "Skip"
         Height          =   255
         Left            =   4320
         TabIndex        =   26
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtCurrentDepth 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Height          =   285
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   840
         Width           =   495
      End
      Begin VB.CheckBox chkLog 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Show Log Window"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   820
         Width           =   1815
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   255
         Left            =   4320
         TabIndex        =   13
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   170
         Width           =   5655
      End
      Begin MSComctlLib.ProgressBar pb 
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblCurrentDepth 
         Caption         =   "Current Depth:"
         Height          =   255
         Left            =   2640
         TabIndex        =   19
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Images found:"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   6000
      Width           =   1095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents HTTP As clsHttpClient
Attribute HTTP.VB_VarHelpID = -1

Private Sub chkDisallow_Click()
    WriteINI "Settings", "DisallowEnabled", chkDisallow.Value, Profile
End Sub

Private Sub chkExtention_Click()
    WriteINI "Settings", "ExtentionEnabled", chkExtention.Value, Profile
End Sub

Private Sub chkLog_Click()
    Select Case CBool(chkLog.Value)
    Case True
        frmLog.Show , Me
    Case False
        frmLog.Hide
    End Select
    chkViewLog.Value = chkLog.Value
End Sub

Private Sub chkViewLog_Click()
    chkLog.Value = chkViewLog.Value
    chkLog_Click
End Sub

Private Sub cmdAdd_Click()
    Dim CurrentNode As Node
    If txtURL.Text <> "" Then
        Set CurrentNode = tvLinks.Nodes.Add
        CurrentNode.Tag = txtURL.Text
        CurrentNode.Text = txtURL.Text
        CurrentNode.Image = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    If MsgBox("Are you sure you want to cancel?", vbYesNo + vbQuestion) = vbYes Then
        HTTP.CancelOperations
        fmeGrabbing.Visible = False
        fmeSettings.Visible = True
        fmeURL.Enabled = True
    End If
End Sub

Private Sub cmdClear_Click()
    If MsgBox("Are you sure you want to clear all entries?", vbYesNo + vbQuestion) = vbYes Then tvLinks.Nodes.Clear
End Sub

Private Sub cmdLoadImageList_Click()
    On Error GoTo ErrorHandler
    Dim FileName As String, FileNumber As Integer, FileData As String, temp() As String, CurrentNode As Node
    Dim ParentIndex As Long, i As Integer
    
    cd.ShowOpen
    
    FileName = cd.FileName
    FileNumber = FreeFile
    
    Open FileName For Binary Access Read As FileNumber
        FileData = String(LOF(FileNumber), Chr(0))
        Get FileNumber, , FileData
    Close FileNumber
    
    temp() = Split(FileData, vbLf)
    
    Set CurrentNode = tvLinks.Nodes.Add
    CurrentNode.Text = GetFileName(FileName)
    CurrentNode.Tag = FileName
    CurrentNode.Image = 3
    ParentIndex = CurrentNode.Index
    
    For i = 0 To UBound(temp)
        If temp(i) = "" Then GoTo NextLine
        temp(i) = RRC(temp(i))
        Set CurrentNode = tvLinks.Nodes.Add(ParentIndex, tvwChild)
        CurrentNode.Text = temp(i)
        CurrentNode.Tag = temp(i)
        CurrentNode.Image = 2
NextLine:
    Next i

    sb.Panels(1).Text = "Loaded '" & GetFileName(FileName) & "'."
    sb.Panels(2).Text = UBound(temp) + 1 & " New entry(s)"
Exit Sub
ErrorHandler:
    MsgBox "Image list was not loaded succesfully." & vbCrLf & "(Error #" & Err.Number & ": " & Err.Description & ")", vbExclamation + vbOKOnly
End Sub

Private Sub cmdLoadProfile_Click()
    frmSave.cmdLoad_Click
    Unload frmSave
End Sub

Private Sub cmdRemove_Click()
    On Error Resume Next
    tvLinks.Nodes.Remove tvLinks.SelectedItem.Index
End Sub

Private Sub cmdSaveImageList_Click()
    On Error GoTo ErrorHandler
    Dim FileName As String, FileNumber As Integer, FileData As String, CurrentNode As Node, LastSib As Long
    Dim i As Integer
    
    If tvLinks.Nodes.Count = 0 Then GoTo ErrorHandler

    cd.ShowSave
    
    FileName = cd.FileName
    FileNumber = FreeFile
    
    If Dir(FileName) <> "" Then Kill FileName
    
    Open FileName For Binary Access Write As FileNumber
    
    For i = 1 To tvLinks.Nodes.Count
        Set CurrentNode = tvLinks.Nodes(i)
        If CurrentNode.Image = 2 Then
            Put FileNumber, , CRC(CurrentNode.Text) & vbLf
        End If
    Next i
    
    Close FileNumber
    
    Exit Sub
ErrorHandler:
    MsgBox "Image list was not saved succesfully." & vbCrLf & "(Error #" & Err.Number & ": " & Err.Description & ")", vbExclamation + vbOKOnly
End Sub

Private Sub cmdSaveImages_Click()
    Dim CurrentNode As Node
    Dim i As Integer
    
    ReDim TempImages(0 To 0)
    For i = 1 To tvLinks.Nodes.Count
        Set CurrentNode = tvLinks.Nodes(i)
        If CurrentNode.Image = 2 Then
            If UBound(TempImages) = 0 Then
                ReDim TempImages(1 To 1)
                TempImages(1) = CurrentNode.Tag
            Else
                ReDim Preserve TempImages(1 To UBound(TempImages) + 1)
                TempImages(UBound(TempImages)) = CurrentNode.Tag
            End If
        End If
    Next i
    
    If UBound(TempImages) = 0 Then MsgBox "No images have been grabbed.", vbOKOnly + vbInformation: Exit Sub
       
    Unload frmSave
    
    frmSave.Show vbModal, Me
    
End Sub


Private Sub cmdSaveProfile_Click()
    frmSave.cmdSave_Click
    Unload frmSave
End Sub

Private Sub cmdSkip_Click()
    HTTP.CancelOperations
End Sub

Private Sub cmdStartGrabbing_Click()
    On Error GoTo ErrorHandler
    
    Dim CurrentNode As Node
    Dim Links() As modHTML.Link, DocumentData As String, URL As String, MaxDepth As Long, CurrentDepth As Long, Extentions() As String, CustomExt As Boolean, Disallows() As String, Disallow As Boolean
    Dim PageTitle As String, ContentType As String, ImagesFound As Long, FileTypes() As String
    Dim i As Integer, k As Integer, x As Integer, ftype As Integer
    
       
    If tvLinks.Nodes.Count = 0 Then MsgBox "You need to specify at least one URL before Grabbing!", vbInformation + vbOKOnly: Exit Sub
    
    Set HTTP = New clsHttpClient
    
    If chkLog.Value = 1 Then frmLog.Show , Me
    
    txtImagesFound.Text = "0"
    fmeSettings.Visible = False
    fmeGrabbing.Visible = True
    fmeURL.Enabled = False
    
    MaxDepth = Val(txtLinkDepth.Text)
    
    If chkExtention.Value = 1 Then
        CustomExt = True
        If InStr(1, txtExtention.Text, ",") = 0 Then
            ReDim Extentions(0 To 0)
            Extentions(0) = txtExtention.Text
        Else
            Extentions() = Split(txtExtention.Text, ",")
        End If
    Else
        CustomExt = False
    End If
    
    If chkFileTypes.Value = 1 Then
        If InStr(1, txtFileTypes.Text, ",") = 0 Then
            ReDim FileTypes(0 To 0)
            FileTypes(0) = txtFileTypes.Text
        Else
            FileTypes() = Split(txtFileTypes.Text, ",")
        End If
    Else
            ReDim FileTypes(0 To 2)
            FileTypes(0) = "jpeg"
            FileTypes(1) = "jpg"
            FileTypes(2) = "png"
    End If
    
    
    If chkDisallow.Value = 1 Then
        Disallow = True
        If InStr(1, txtDisallow.Text, ",") = 0 Then
            ReDim Disallows(0 To 0)
            Disallows(0) = txtExtention.Text
        Else
            Disallows() = Split(txtDisallow.Text, ",")
        End If
    Else
        Disallow = False
    End If
    
    frmLog.LogEvent "Started Grab process at " & Time, 6
    i = 1
    Do Until i >= tvLinks.Nodes.Count + 1
        Set CurrentNode = tvLinks.Nodes(i)
        DocumentData = ""
        
        If CurrentNode.Image <> 1 Then GoTo NextPage
        
        
        CurrentDepth = GetParents(CurrentNode)
        If CurrentDepth >= MaxDepth Then GoTo NextPage
        txtCurrentDepth.Text = CurrentDepth + 1
        
        URL = CurrentNode.Tag
        
        Set HTTP = New clsHttpClient
        
        sb.Panels(1).Text = URL
        frmLog.LogEvent "Connecting to '" & URL & "'", 1
        txtStatus.Text = "Connecting to " & URL
        If HTTP.DownloadFile(URL, , DocumentData, , , "text") <> 0 Then GoTo NextPage
        
        If fmeGrabbing.Visible = False Then GoTo Done
        
GetLinks:
        modHTML.ReadDocument DocumentData, URL, PageTitle, Links()
        If PageTitle <> "" Then CurrentNode.Text = PageTitle
        
        Dim LinkNode As Node
        For k = 1 To UBound(Links)
            If Links(k).HREF = "" Then GoTo NextLink
            
            For ftype = 0 To UBound(FileTypes)
            
                If LCase(Right(Links(k).HREF, Len(FileTypes(ftype)))) = FileTypes(ftype) Then
                    Set LinkNode = tvLinks.Nodes.Add(CurrentNode.Index, tvwChild, , Links(k).HREF, 2)
                    LinkNode.Tag = Links(k).HREF
                    ImagesFound = ImagesFound + 1
                    txtImagesFound.Text = ImagesFound
                    GoTo NextLink
                End If
            
            Next ftype
            
            If CustomExt = True Then
                Dim qPos As Long, TempLink As String
                For x = 0 To UBound(Extentions)
                    TempLink = Links(k).HREF
                    If Right(TempLink, Len(Extentions(x))) = Extentions(x) Then GoTo CheckDisallows
                    qPos = LastPos(1, TempLink, "?")
                    If qPos <> 0 Then TempLink = Mid(TempLink, 1, qPos - 1)
                    If Right(TempLink, Len(Extentions(x))) = Extentions(x) Then GoTo CheckDisallows
                Next x
                GoTo NextLink
            End If
            
CheckDisallows:
            If Disallow = True Then
                For x = 0 To UBound(Disallows)
                    TempLink = Links(k).HREF
                    If Right(TempLink, Len(Disallows(x))) = Disallows(x) Then GoTo NextLink
                Next x
            End If
AddLink:
            Set LinkNode = tvLinks.Nodes.Add(CurrentNode.Index, tvwChild, , Links(k).HREF, 1)
            LinkNode.Tag = Links(k).HREF

NextLink:
        Next k
           
NextPage:
    i = i + 1
    If fmeGrabbing.Visible = False Then GoTo Done
    Loop
    
Done:
    frmLog.LogEvent "Grab process ended at " & Time, 8
    
    sb.Panels(1).Text = "SickAnimations PicGrab2 Pro - Tim Cinel 2003."
    sb.Panels(2).Text = tvLinks.Nodes.Count & " Entry(s)"
    
    fmeSettings.Visible = True
    fmeGrabbing.Visible = False
    fmeURL.Enabled = True
    
    Exit Sub

ErrorHandler:
    frmLog.LogEvent "Error #" & Err.Number & " in frmMain.cmdStartGrabbing_Click(): " & Err.Description, 3
    Resume Next

End Sub

Function GetParents(ByVal N As Node) As Long
CheckAgain:
    If InStr(1, N.FullPath, "/|¤|\") = 0 Then Exit Function
    Set N = N.Parent
    GetParents = GetParents + 1
    GoTo CheckAgain
End Function

Private Sub cmdGrabAndGo_Click()
    Dim i As Integer, CType As String
    ReDim TempImages(0 To 0)
    
    If tvLinks.Nodes.Count = 0 Then MsgBox "You need to specify at least one URL before Grabbing!", vbInformation + vbOKOnly: Exit Sub
    
    frmSave.cmdSaveAll.Caption = "Grab and Go!"
    frmSave.txtDownloadFrom.Enabled = False
    frmSave.txtDownloadTo.Enabled = False
    
    frmSave.Show vbModal, Me
    
    cmdStartGrabbing_Click
    
    Dim CurrentNode As Node
    
    ReDim TempImages(0 To 0)
    For i = 1 To tvLinks.Nodes.Count
        Set CurrentNode = tvLinks.Nodes(i)
        If CurrentNode.Image = 2 Then
            If UBound(TempImages) = 0 Then
                ReDim TempImages(1 To 1)
                TempImages(1) = CurrentNode.Tag
            Else
                ReDim Preserve TempImages(1 To UBound(TempImages) + 1)
                TempImages(UBound(TempImages)) = CurrentNode.Tag
            End If
        End If
    Next i
    
    If UBound(TempImages) = 0 Then MsgBox "No images have been grabbed.", vbOKOnly + vbInformation: Exit Sub
    
    frmMain.Hide
    frmSaving.Show
    
    If chkContentType.Value = 1 Then
        CType = txtContentType.Text
    Else
        CType = "image"
    End If
    
    frmSaving.DownloadFiles TempImages(), frmSave.txtName.Text, Val(frmSave.txtMinimumDigits.Text), Val(frmSave.txtStartFrom.Text), frmSave.Directory.Path & "\", , , CType

    Unload frmSave

End Sub

Private Sub Form_Load()
On Error Resume Next
    Profile = GetSetting("PicGrab", "Settings", "Profile", App.Path & "\default.ini")
    frmSave.Form_Load
    Unload frmSave
    
    chkDisallow.Value = Val(ReadINI("Settings", "DisallowEnabled", Profile, "0"))
    txtDisallow.Text = ReadINI("Settings", "Disallow", Profile, "com/")
    
    txtLinkDepth.Text = Val(ReadINI("Settings", "LinkDepth", Profile, "2"))
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub HTTP_DownloadBegin(Location As String, FileLength As Long)
    On Error Resume Next
    txtStatus.Text = "Download has begun."
    frmLog.LogEvent "Downloading '" & Location & "'", 5
    pb.Max = FileLength
    sb.Panels(1).Text = Location
    sb.Panels(2).Text = "0%"
End Sub

Private Sub HTTP_DownloadComplete(URL As String, TimeMs As Long, BytesDownloaded As Long)
    On Error Resume Next
    txtStatus.Text = "Download complete."
    pb.Value = 0
    sb.Panels(1).Text = ""
    sb.Panels(2).Text = ""
    
    frmLog.LogEvent "Download complete: '" & URL & "' (" & Round(BytesDownloaded / 1024) & "kb) in " & Round(TimeMs / 1000) & " seconds. Average rate of " & Round((BytesDownloaded / 1024) / (TimeMs / 1000), 2) & "kb/s.", 2
End Sub

Private Sub HTTP_DownloadError(Number As Integer, Description As String, URL As String)
    frmLog.LogEvent "Error #:" & Number & ": '" & Description & "' while downloading '" & URL & "'", 3
End Sub

Private Sub HTTP_DownloadProgress(Downloaded As Long, Total As Long, Percent As Single)
    On Error Resume Next
    pb.Value = Downloaded
    txtStatus.Text = "Downloaded " & Round(Downloaded / 1024, 2) & "kb of " & Round(Total / 1024, 2) & "kb."
    sb.Panels(2).Text = Round(Percent, 2) & "%"
End Sub

Function LastPos(Start, String1, String2) As Long
    Dim temp
    temp = InStr(Start, String1, String2)
    Do Until temp = 0
        LastPos = temp
        temp = InStr(temp + 1, String1, String2)
    Loop
End Function

Private Function GetFileName(FilePath As String) As String
    Dim temp() As String
    If InStr(1, FilePath, "\") = 0 Then GetFileName = FilePath: Exit Function
    
    temp() = Split(FilePath, "\")
    
    GetFileName = temp(UBound(temp))
    
End Function

Private Sub HTTP_DownloadRedirect(OldUrl As String, NewUrl As String)
    Dim CustomEx As Boolean, Extentions() As String, Disallows() As String, Disallow As Boolean, i As Integer, x As Integer
    frmLog.LogEvent "Redirected to '" & NewUrl & "'", 4
    
    If chkExtention.Value = 1 Then
        CustomEx = True
        If InStr(1, txtExtention.Text, ",") = 0 Then
            ReDim Extentions(0 To 0)
            Extentions(0) = txtExtention.Text
        Else
            Extentions() = Split(txtExtention.Text, ",")
        End If
    Else
        CustomEx = False
    End If
    
    If chkDisallow.Value = 1 Then
        Disallow = True
        If InStr(1, txtDisallow.Text, ",") = 0 Then
            ReDim Disallows(0 To 0)
            Disallows(0) = txtExtention.Text
        Else
            Disallows() = Split(txtDisallow.Text, ",")
        End If
    End If

    If CustomEx = True Then
        Dim qPos As Long, TempLink As String
        For x = 0 To UBound(Extentions)
            TempLink = NewUrl
            If Right(TempLink, Len(Extentions(x))) = Extentions(x) Then GoTo CheckIfDisallowed
            qPos = LastPos(1, TempLink, "?")
            If qPos <> 0 Then TempLink = Mid(TempLink, 1, qPos - 1)
            If Right(TempLink, Len(Extentions(x))) = Extentions(x) Then GoTo CheckIfDisallowed
        Next x
        GoTo Cancel
    Else
        GoTo Allow
    End If
CheckIfDisallowed:
    If Disallow = True Then
        For x = 0 To UBound(Disallows)
            TempLink = NewUrl
            If Right(TempLink, Len(Disallows(x))) = Disallows(x) Then GoTo Cancel
            qPos = LastPos(1, TempLink, "?")
            If qPos <> 0 Then TempLink = Mid(TempLink, 1, qPos - 1)
            If Right(TempLink, Len(Disallows(x))) = Disallows(x) Then GoTo Cancel
        Next x
        GoTo Allow
    End If
    
Cancel:
    HTTP.CancelOperations
Exit Sub
Allow:
End Sub


Private Sub txtDisallow_LostFocus()
    WriteINI "Settings", "Disallow", txtDisallow.Text, Profile
End Sub

Private Sub txtExtention_LostFocus()
    WriteINI "Settings", "Extention", txtExtention.Text, Profile
End Sub

Private Sub txtLinkDepth_LostFocus()
    If Val(txtLinkDepth.Text) <= 0 Then txtLinkDepth.Text = 1
    WriteINI "Settings", "LinkDepth", CStr(Val(txtLinkDepth.Text)), Profile
End Sub

Private Function CRC(InputString As String) As String
'Correct Reserved Characters - Encodes characters incase they are reserved.
CRC = InputString

CRC = Replace(CRC, "&", "&0")
CRC = Replace(CRC, vbLf, "&1")

End Function

Private Function RRC(InputString As String)
'Decodes characters encoded by CRC
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
        RRC = LeftString & vbLf & RightString
    End Select
    
    CurrentPos = InStr(CurrentPos + 1, RRC, "&")
Loop
End Function

