VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSave 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Save"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSaveAll 
      Caption         =   "Start Download"
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   5160
      Width           =   4575
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   4080
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "INI Files (*.ini)|*.ini"
   End
   Begin VB.Frame fmePresets 
      Caption         =   "Presets"
      Height          =   615
      Left            =   0
      TabIndex        =   13
      Top             =   4440
      Width           =   4575
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save Profile"
         Height          =   255
         Left            =   2280
         TabIndex        =   15
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Load Profile"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.DriveListBox Drive 
      Height          =   315
      Left            =   0
      TabIndex        =   12
      Top             =   120
      Width           =   4575
   End
   Begin VB.Frame fmeName 
      Caption         =   "Name"
      Height          =   1215
      Left            =   0
      TabIndex        =   6
      Top             =   2160
      Width           =   4575
      Begin VB.CommandButton cmdDefault 
         Caption         =   "Default"
         Height          =   255
         Left            =   2400
         TabIndex        =   11
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "Help"
         Height          =   255
         Left            =   3480
         TabIndex        =   9
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Text            =   "/title/ #/number/./extention/"
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label lblHelp 
         Caption         =   "If confused, click default."
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblTypeName 
         Caption         =   "Name Code:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame fmeNumber 
      Caption         =   "Number:"
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   3360
      Width           =   4575
      Begin VB.TextBox txtDownloadTo 
         Height          =   285
         Left            =   2400
         TabIndex        =   20
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtDownloadFrom 
         Height          =   285
         Left            =   1320
         TabIndex        =   17
         Text            =   "1"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtMinimumDigits 
         Height          =   285
         Left            =   2880
         MaxLength       =   1
         TabIndex        =   5
         Text            =   "4"
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtStartFrom 
         Height          =   285
         Left            =   960
         MaxLength       =   4
         TabIndex        =   3
         Text            =   "0"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblTo 
         Alignment       =   2  'Center
         Caption         =   "to"
         Height          =   255
         Left            =   1920
         TabIndex        =   19
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblFrom 
         BackStyle       =   0  'Transparent
         Caption         =   "Download from"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Minimum Digits:"
         Height          =   255
         Left            =   1680
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblStartOn 
         Caption         =   "Start From:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.DirListBox Directory 
      Height          =   1665
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   4575
   End
End
Attribute VB_Name = "frmSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurrentFileNumber As Single

Private Sub cmdDefault_Click()
    txtName.Text = "/title/ #/number/./extention/"
End Sub

Private Sub cmdHelp_Click()
    frmNameCodeHelp.Show vbModal, Me
End Sub

Sub cmdLoad_Click()
    On Error GoTo ErrorHandler
    Dim FileName As String
    
    cd.ShowOpen
    FileName = cd.FileName
    
    SaveSetting "PicGrab", "Settings", "Profile", FileName
    Profile = FileName
    Form_Load
    Exit Sub
ErrorHandler:
    MsgBox "Profile was not loaded succesfully." & vbCrLf & "(Error #" & Err.Number & ": " & Err.Description & ")", vbExclamation + vbOKOnly
End Sub

Private Sub cmdRemove_Click()
    On Error Resume Next
    lstFiles.RemoveItem lstFiles.ListIndex
End Sub

Sub cmdSave_Click()
    On Error GoTo ErrorHandler
    Dim FileName As String
    
    cd.ShowSave
    FileName = cd.FileName
    
    WriteINI "Settings", "NameCode", txtName.Text, FileName
    WriteINI "Settings", "MinDigits", txtMinimumDigits.Text, FileName
    WriteINI "Settings", "StartFrom", txtStartFrom.Text, FileName
    WriteINI "Settings", "Path", Directory.Path, FileName
    WriteINI "Settings", "ExtentionEnabled", frmMain.chkExtention.Value, FileName
    WriteINI "Settings", "Extention", frmMain.txtExtention.Text, FileName
    WriteINI "Settings", "DisallowEnabled", frmMain.chkDisallow.Value, FileName
    WriteINI "Settings", "Disallow", frmMain.txtDisallow.Text, FileName
    
    SaveSetting "PicGrab", "Settings", "Profile", FileName
    Profile = FileName
    
    Exit Sub
ErrorHandler:
    MsgBox "Profile was not saved succesfully." & vbCrLf & "(Error #" & Err.Number & ": " & Err.Description & ")", vbExclamation + vbOKOnly
End Sub

Private Sub cmdSaveAll_Click()
    Dim CType As String
    
    If frmMain.chkContentType.Value = 1 Then
        CType = frmMain.txtContentType.Text
    Else
        CType = "image"
    End If

    
    If UBound(TempImages) = 0 Then Me.Hide: Exit Sub
    
    txtDownloadFrom.Text = Val(txtDownloadFrom.Text)
    If Val(txtDownloadFrom.Text) < 0 Then txtDownloadFrom.Text = 1
    
    Me.Hide
    frmMain.Hide
    frmSaving.Show
    frmSaving.DownloadFiles TempImages(), txtName.Text, Val(txtMinimumDigits.Text), Val(txtStartFrom.Text), Directory.Path & "\", Val(txtDownloadFrom.Text), Val(txtDownloadTo.Text), CType
    Unload Me
    
End Sub

Private Sub Drive_Change()
    On Error Resume Next
    Directory.Path = Drive.Drive
End Sub

Sub Form_Load()
    On Error Resume Next
    txtDownloadTo.Text = UBound(modMain.TempImages)
    
    Me.Caption = "Save - Profile '" & GetFileName(Profile) & "'"
    frmMain.Caption = "PicGrab2 - Profile '" & GetFileName(Profile) & "'"
    
    txtName.Text = ReadINI("Settings", "NameCode", Profile, "/title/ #/number/./extention/")
    txtMinimumDigits.Text = ReadINI("Settings", Profile, "MinDigits", "4")
    txtStartFrom.Text = ReadINI("Settings", "StartFrom", Profile, "0")
    Directory.Path = ReadINI("Settings", "Path", Profile, App.Path)
    Drive.Drive = Mid(Directory.Path, 1, 2)
    frmMain.txtLinkDepth.Text = Val(ReadINI("Settings", "LinkDepth", Profile, "2"))
    frmMain.chkExtention.Value = ReadINI("Settings", "ExtentionEnabled", Profile, "0")
    frmMain.txtExtention.Text = ReadINI("Settings", "Extention", Profile, "html,htm,/")
    frmMain.chkDisallow.Value = ReadINI("Settings", "DisallowEnabled", Profile, "0")
    frmMain.txtDisallow.Text = ReadINI("Settings", "Disallow", Profile, "com/")

End Sub

Private Sub txtDownloadFrom_LostFocus()
    txtDownloadFrom.Text = Val(txtDownloadFrom.Text)
    If Val(txtDownloadFrom.Text) < LBound(modMain.TempImages) Then txtDownloadFrom.Text = LBound(modMain.TempImages)
    If Val(txtDownloadFrom.Text) > Val(txtDownloadTo.Text) Then txtDownloadFrom.Text = Val(txtDownloadTo.Text)
End Sub

Private Sub txtDownloadTo_LostFocus()
    txtDownloadTo.Text = Val(txtDownloadTo.Text)
    If Val(txtDownloadTo.Text) > UBound(modMain.TempImages) Then txtDownloadTo.Text = UBound(modMain.TempImages)
    If Val(txtDownloadTo.Text) < Val(txtDownloadFrom.Text) Then txtDownloadTo.Text = Val(txtDownloadFrom.Text)
End Sub

Private Sub txtMinimumDigits_LostFocus()
    txtMinimumDigits.Text = Val(txtMinimumDigits.Text)
    If Val(txtMinimumDigits.Text) > 5 Then txtMinimumDigits.Text = "5"
    
    SaveSetting "PicGrab", "Settings", "MinDigits", Val(txtMinimumDigits.Text)
End Sub

Private Sub txtStartFrom_LostFocus()
    txtStartFrom.Text = Val(txtStartFrom.Text)
End Sub

Private Function GetFileName(FilePath As String) As String
    Dim temp() As String
    If InStr(1, FilePath, "\") = 0 Then GetFileName = FilePath: Exit Function
    
    temp() = Split(FilePath, "\")
    
    GetFileName = temp(UBound(temp))
    
End Function
