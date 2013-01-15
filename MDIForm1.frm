VERSION 5.00
Begin VB.MDIForm MDIMain 
   BackColor       =   &H80000014&
   Caption         =   "Blue Line Cabs - AClick Solutions - HimSoft Inc"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox MDIPic 
      Align           =   1  'Align Top
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   1200
      Left            =   0
      ScaleHeight     =   1200
      ScaleWidth      =   15240
      TabIndex        =   0
      Top             =   0
      Width           =   15240
      Begin VB.CommandButton cmdMain 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   8
         Left            =   11640
         Picture         =   "MDIForm1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdMain 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   240
         Picture         =   "MDIForm1.frx":0742
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdMain 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   2040
         Picture         =   "MDIForm1.frx":0B84
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdMain 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   3240
         Picture         =   "MDIForm1.frx":0FC6
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdMain 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   7
         Left            =   10440
         Picture         =   "MDIForm1.frx":1408
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdMain 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   5
         Left            =   7440
         Picture         =   "MDIForm1.frx":184A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdMain 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   6
         Left            =   8640
         Picture         =   "MDIForm1.frx":1C8C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdMain 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   3
         Left            =   4440
         Picture         =   "MDIForm1.frx":20CE
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdMain 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   4
         Left            =   5640
         Picture         =   "MDIForm1.frx":2510
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdMain 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   9
         Left            =   13440
         Picture         =   "MDIForm1.frx":2952
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ENQUIRY"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11640
         TabIndex        =   19
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "LOG OFF"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   13440
         TabIndex        =   18
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "FARE"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10440
         TabIndex        =   17
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "DRIVER"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8640
         TabIndex        =   16
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CABS"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7440
         TabIndex        =   15
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "INVOICE"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5640
         TabIndex        =   14
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "MAP VIEW"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   13
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "BOOKING"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   12
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CUSTOMER"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ADMIN"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   1215
      End
      Begin VB.Image Image1 
         Height          =   9870
         Left            =   0
         Picture         =   "MDIForm1.frx":2D94
         Top             =   -8520
         Width           =   15795
      End
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdMain_Click(Index As Integer)
Select Case Index
    Case 0
    'MsgBox "Admin Module Under Construction"
    Unload frmHome
    Call MeCenter(5010, 3705, frmManageUsers)
    Call DisableCmd(Index)
    
    Case 1
    Call CenterMe(frmCustList)
    frmCustList.Show
    Call DisableCmd(Index)
    
    Case 2
    Unload frmHome
    Call CenterMe(frmBookingList)
    frmBookingList.Show
    Call DisableCmd(Index)
    
    Case 3
    Call CenterMe(frmMap)
    frmMap.Show
    Call DisableCmd(Index)
    
    Case 4
    Call CenterMe(frmInvoiceList)
    frmInvoiceList.Show
    Call DisableCmd(Index)
    
    Case 5
    Call CenterMe(frmCabList)
    Unload frmHome
    frmCabList.Show
    Call DisableCmd(Index)
    
    Case 6
    Call CenterMe(frmDriverList)
    Unload frmHome
    frmDriverList.Show
    Call DisableCmd(Index)
    
    Case 7
    Call CenterMe(frmFareList)
    Unload frmHome
    frmFareList.Show
    Call DisableCmd(Index)
    
    Case 8
    MsgBox "Enquiry Module Under Construction"
    
    Case 9
    Unload frmHome
    For i = 0 To 9
        MDIMain.cmdMain(i).Enabled = False
    Next i
        
    Call MeCenter(4515, 2010, frmLogin)
           
    End Select
    
End Sub

Private Sub MDIForm_Load()
login = True

For i = 0 To 9
    cmdMain(i).Enabled = False
Next i

Call MeCenter(4515, 2010, frmLogin)
'frmHome.Show
End Sub


'Public Sub DisableCmd(Index As Integer)
 '   Dim i As Integer
    
  '  For i = 0 To 9
   '     If (i <> Index) Then
            'cmdMain(i).Enabled = False
    '    End If
   ' Next
'End Sub


'Public Sub EnableCmd()
'    Dim i As Integer
'
'    For i = 0 To 9 Step 1
'        cmdMain(i).Enabled = True
'    Next
'
'End Sub
