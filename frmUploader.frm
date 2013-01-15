VERSION 5.00
Begin VB.Form frmUploader 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frameUpload 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   495
         Left            =   2400
         Picture         =   "frmUploader.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3000
         Width           =   1215
      End
      Begin VB.DirListBox DirListBox 
         Height          =   2115
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1935
      End
      Begin VB.FileListBox FileListBox 
         Height          =   2040
         Left            =   2160
         TabIndex        =   2
         Top             =   720
         Width           =   1815
      End
      Begin VB.DriveListBox DriveListBox 
         Height          =   315
         Left            =   840
         TabIndex        =   1
         Top             =   240
         Width           =   2535
      End
      Begin VB.CommandButton cmdUpload 
         Caption         =   "&Select"
         Height          =   495
         Left            =   480
         Picture         =   "frmUploader.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label lblPickpath 
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmUploader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim whoCalledUploader As String

Dim picpath As String
Dim filename As String
Dim filepath As String

Private Sub cmdCancel_Click()
picpath = "xxxx" 'setting the picpath as null when cancelled button is hit
Me.lblPickpath.Caption = picpath
frmUploader.Visible = False

If (whoCalledUploader = "driver") Then
    frmDriverRegistration.Enabled = True
ElseIf (whoCalledUploader = "cab") Then
        frmCabRegistration.Enabled = True
End If
End Sub

Private Sub cmdUpload_Click()

'Uploading the picture to the picture frame
'picpath = filepath & "\" & filename
'Me.lblPickpath.Caption = picpath
frmUploader.Visible = False
'MsgBox picpath

If (whoCalledUploader = "driver") Then
    frmDriverRegistration.Enabled = True
ElseIf (whoCalledUploader = "cab") Then
        frmCabRegistration.Enabled = True
End If

End Sub

Private Sub DirListBox_Change()
Me.FileListBox.Path = Me.DirListBox.Path
End Sub

Private Sub DriveListBox_Change()

'Code to handle if any errors occur during driver selection

Dim msg As String
Dim result As String

On Error GoTo Error
DirListBox.Path = DriveListBox.Drive
Exit Sub
Error: msg = "Error: " & Err.Number & ": " & Err.Description
result = MsgBox(msg, vbOKCancel + vbExclamation, "No Data")

If result = vbOK Then
Resume
Else
DriveListBox.Drive = DirListBox.Path
Err.Clear
Exit Sub
End If
End Sub

Private Sub FileListBox_Click()

filepath = FileListBox.Path
filename = FileListBox.filename

picpath = filepath & "\" & filename
Me.lblPickpath.Caption = picpath

End Sub

Public Sub setWhoCalled(frmObject As String)
whoCalledUploader = frmObject
End Sub

Private Sub Form_Load()
picpath = ""
End Sub
