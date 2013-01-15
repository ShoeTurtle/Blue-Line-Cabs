VERSION 5.00
Begin VB.Form frmStreetList 
   Caption         =   "Street List - Blue Line"
   ClientHeight    =   8700
   ClientLeft      =   120
   ClientTop       =   435
   ClientWidth     =   11205
   LinkTopic       =   "Form1"
   ScaleHeight     =   8700
   ScaleWidth      =   11205
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "ADD AREA"
      Height          =   615
      Left            =   3120
      TabIndex        =   9
      Top             =   7080
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   4080
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   2880
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CLOSE"
      Height          =   615
      Left            =   8640
      TabIndex        =   4
      Top             =   7080
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SAVE"
      Height          =   615
      Left            =   5880
      TabIndex        =   3
      Top             =   7080
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   1680
      Width           =   3015
   End
   Begin VB.Label Label8 
      Caption         =   "Street Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label7 
      Caption         =   "Street Id :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Area No :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Street List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "frmStreetList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub

Private Sub Label1_Click()

End Sub
