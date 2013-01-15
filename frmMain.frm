VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Blue-Line-Cab"
   ClientHeight    =   9270
   ClientLeft      =   1035
   ClientTop       =   750
   ClientWidth     =   12165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   618
   ScaleMode       =   0  'User
   ScaleWidth      =   811
   Begin VB.CommandButton Command7 
      Caption         =   "Exit"
      Height          =   495
      Left            =   10320
      TabIndex        =   6
      Top             =   8400
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Driver"
      Height          =   495
      Left            =   10320
      TabIndex        =   5
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Customer"
      Height          =   495
      Left            =   10320
      TabIndex        =   4
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Taxi"
      Height          =   495
      Left            =   10320
      TabIndex        =   3
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Invoice"
      Height          =   495
      Left            =   10320
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Trip Status"
      Height          =   495
      Left            =   10320
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Booking"
      Height          =   495
      Left            =   10320
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   4140
      Left            =   2400
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   5513.514
      ScaleMode       =   0  'User
      ScaleWidth      =   6735.849
      TabIndex        =   8
      Top             =   3120
      Width           =   6180
   End
   Begin VB.Label Label1 
      Caption         =   "Bue Line Cab"
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
      Left            =   3840
      TabIndex        =   7
      Top             =   840
      Width           =   3255
   End
   Begin VB.Menu Entry 
      Caption         =   "Entry"
      Begin VB.Menu Booking 
         Caption         =   "Booking"
      End
      Begin VB.Menu Invoice 
         Caption         =   "Invoice"
      End
      Begin VB.Menu DriverPayment 
         Caption         =   "Driver Payment"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Maintenance 
      Caption         =   "Maintenance"
      Begin VB.Menu Taxi 
         Caption         =   "Taxi"
      End
      Begin VB.Menu Customer 
         Caption         =   "Customer"
      End
      Begin VB.Menu Driver 
         Caption         =   "Driver"
      End
   End
   Begin VB.Menu Report 
      Caption         =   "Report"
      Begin VB.Menu DriverList 
         Caption         =   "Driver List"
      End
      Begin VB.Menu CustomerList 
         Caption         =   "Customer List"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
