VERSION 5.00
Begin VB.Form frmMap 
   BorderStyle     =   0  'None
   ClientHeight    =   11190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15285
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11190
   ScaleWidth      =   15285
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frameMapCmd 
      Height          =   975
      Left            =   8280
      TabIndex        =   10
      Top             =   240
      Width           =   5655
      Begin VB.CommandButton cmdGoBack 
         Caption         =   "Go Back"
         Height          =   495
         Left            =   3840
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "Show Map"
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdAddMap 
         Caption         =   "Add Map"
         Height          =   495
         Left            =   2040
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame frameUpload 
      Height          =   3375
      Left            =   8280
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   5655
      Begin VB.DriveListBox DriveListBox 
         Height          =   315
         Left            =   1440
         TabIndex        =   9
         Top             =   240
         Width           =   2655
      End
      Begin VB.FileListBox FileListBox 
         Height          =   1845
         Left            =   2760
         TabIndex        =   8
         Top             =   720
         Width           =   2655
      End
      Begin VB.DirListBox DirListBox 
         Height          =   1890
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   2415
      End
      Begin VB.CommandButton cmdMapCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   3120
         TabIndex        =   6
         Top             =   2760
         Width           =   1815
      End
      Begin VB.CommandButton cmdMapUpload 
         Caption         =   "Upload"
         Height          =   375
         Left            =   600
         TabIndex        =   5
         Top             =   2760
         Width           =   1815
      End
   End
   Begin VB.Frame frameLocator 
      Caption         =   "Locator"
      Height          =   975
      Left            =   1440
      TabIndex        =   2
      Top             =   240
      Width           =   6735
      Begin VB.ComboBox comboAreaSelection 
         Height          =   315
         Left            =   2760
         TabIndex        =   0
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Area Selection"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.PictureBox picAreaMap 
      Height          =   7695
      Left            =   1440
      Picture         =   "frmMap.frx":0000
      ScaleHeight     =   7635
      ScaleWidth      =   12435
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   12495
   End
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim flgA As Boolean

Private Sub cmdAddMap_Click()

frameLocator.Enabled = False
Me.cmdGoBack.Enabled = False
Me.cmdView.Enabled = False
'Me.cmdAddMap.Enabled = False

frameUpload.Visible = True
End Sub


Private Sub cmdGoBack_Click()
Unload Me
Call EnableCmd
End Sub

Private Sub cmdMapCancel_Click()

Me.cmdAddMap.Enabled = True
Me.cmdGoBack.Enabled = True
Me.cmdView.Enabled = True
frameLocator.Enabled = True
frameUpload.Visible = False
End Sub

Private Sub cmdMapUpload_Click()

Me.cmdAddMap.Enabled = True
Me.cmdGoBack.Enabled = True
Me.cmdView.Enabled = True
frameLocator.Enabled = True
frameUpload.Visible = False


End Sub

Private Sub cmdView_Click()

If (flgA) Then
    Me.picAreaMap.Visible = True
    flgA = False
    cmdView.Caption = "Hide Map"
    Me.cmdAddMap.Enabled = False
    Me.cmdGoBack.Enabled = False
    frameLocator.Enabled = False
Else
    Me.picAreaMap.Visible = False
    flgA = True
    cmdView.Caption = "Show Map"
    Me.cmdAddMap.Enabled = True
    Me.cmdGoBack.Enabled = True
    frameLocator.Enabled = True
End If

End Sub

Private Sub Form_Load()
flgA = True

End Sub
