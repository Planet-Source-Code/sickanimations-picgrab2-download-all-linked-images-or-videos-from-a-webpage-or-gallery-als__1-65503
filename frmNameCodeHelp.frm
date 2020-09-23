VERSION 5.00
Begin VB.Form frmNameCodeHelp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "NameCode Help"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   1560
      MaxLength       =   40
      TabIndex        =   13
      Text            =   "GM_Car.jpg"
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox txtStartNumber 
      Height          =   285
      Left            =   1560
      MaxLength       =   9
      TabIndex        =   12
      Text            =   "16"
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox txtPlaceHolders 
      Height          =   285
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   11
      Text            =   "4"
      Top             =   3240
      Width           =   495
   End
   Begin VB.TextBox txtOutputName 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "Image GM_Car #0016.jpg"
      Top             =   4080
      Width           =   2895
   End
   Begin VB.TextBox txtNameCode 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Text            =   "Image /title/ #/number/./extention/"
      Top             =   3600
      Width           =   2895
   End
   Begin VB.Label lblSyntax 
      Caption         =   $"frmNameCodeHelp.frx":0000
      Height          =   975
      Left            =   0
      TabIndex        =   10
      Top             =   1200
      Width           =   4575
   End
   Begin VB.Label lblIntro 
      Caption         =   $"frmNameCodeHelp.frx":013F
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   360
      Width           =   4695
   End
   Begin VB.Label lblActualName 
      Alignment       =   1  'Right Justify
      Caption         =   "Output File Name:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label lblNameCode 
      Alignment       =   1  'Right Justify
      Caption         =   "Name Code:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label lblDigits 
      Alignment       =   1  'Right Justify
      Caption         =   "Minimum Digits"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label lblNumber 
      Alignment       =   1  'Right Justify
      Caption         =   "Start number:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label lblFileName 
      Alignment       =   1  'Right Justify
      Caption         =   "File name:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label lblExample 
      Caption         =   "Example:"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label lblPicGrab 
      Caption         =   "PicGrab uses a speacial, simple code for naming files."
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmNameCodeHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
MsgBox "If you still do not understand, just click the 'Default' button.", vbInformation + vbOKOnly, "Help"
End Sub

Private Sub txtFileName_Change()
txtNameCode_Change
End Sub

Private Sub txtNameCode_Change()
txtOutputName.Text = modNameCode.CodeToName(txtFileName.Text, txtNameCode.Text, Val(txtPlaceHolders.Text), Val(txtStartNumber.Text))
End Sub

Private Sub txtPlaceHolders_Change()
txtPlaceHolders.Text = Val(txtPlaceHolders.Text)
txtNameCode_Change
End Sub

Private Sub txtStartNumber_Change()
txtStartNumber.Text = Val(txtStartNumber.Text)
txtNameCode_Change
End Sub
