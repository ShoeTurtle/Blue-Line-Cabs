VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmDriverRegistration 
   BorderStyle     =   0  'None
   Caption         =   "Driver Registration - Blue Line"
   ClientHeight    =   11190
   ClientLeft      =   240
   ClientTop       =   -60
   ClientWidth     =   15285
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmDriverRegistration.frx":0000
   ScaleHeight     =   11190
   ScaleWidth      =   15285
   ShowInTaskbar   =   0   'False
   Begin MSMask.MaskEdBox txtDriverLicNo 
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   3840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   6
      Mask            =   "######"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtDriverPhNo 
      Height          =   375
      Left            =   8040
      TabIndex        =   5
      Top             =   2160
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##########"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtDriverPin 
      Height          =   375
      Left            =   7680
      TabIndex        =   19
      Top             =   7080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   6
      Mask            =   "######"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdDriverRegClose 
      Caption         =   "Go &Back"
      Height          =   735
      Left            =   7440
      Picture         =   "frmDriverRegistration.frx":58FC
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton cmdCustListRefresh 
      Caption         =   "&Refresh"
      Height          =   735
      Left            =   3480
      Picture         =   "frmDriverRegistration.frx":5A46
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton cmdDriverRegSave 
      Caption         =   "&Add"
      Height          =   735
      Left            =   5520
      Picture         =   "frmDriverRegistration.frx":5FD0
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton cmdBrowsePlus 
      Height          =   375
      Left            =   12000
      Picture         =   "frmDriverRegistration.frx":6C9A
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3600
      Width           =   2175
   End
   Begin VB.TextBox txtDriverBlock 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3000
      TabIndex        =   16
      Top             =   6600
      Width           =   2655
   End
   Begin VB.TextBox txtDriverCity 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7680
      TabIndex        =   17
      Top             =   6600
      Width           =   2775
   End
   Begin VB.TextBox txtDriverStreet 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7680
      TabIndex        =   15
      Top             =   6120
      Width           =   2775
   End
   Begin VB.TextBox txtDriverState 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3000
      TabIndex        =   18
      Top             =   7080
      Width           =   2655
   End
   Begin VB.TextBox txtDriverHouseNo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3000
      TabIndex        =   14
      Top             =   6120
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   7920
      ScaleHeight     =   240
      ScaleWidth      =   2235
      TabIndex        =   40
      Top             =   4560
      Width           =   2295
      Begin VB.CommandButton cmdTrainedOn 
         Height          =   240
         Left            =   2010
         Picture         =   "frmDriverRegistration.frx":7224
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         Width           =   225
      End
      Begin VB.Label lblTrainedOn 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   0
         TabIndex        =   42
         Top             =   0
         Width           =   45
      End
      Begin VB.Label Label14 
         Caption         =   "Drop Down Selection"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.PictureBox picCombo 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   3120
      ScaleHeight     =   240
      ScaleWidth      =   2235
      TabIndex        =   37
      Top             =   4560
      Width           =   2295
      Begin VB.CommandButton cmdLicType 
         Height          =   240
         Left            =   2010
         Picture         =   "frmDriverRegistration.frx":75C6
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Width           =   225
      End
      Begin VB.Label lblLicType 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   0
         TabIndex        =   39
         Top             =   0
         Width           =   45
      End
      Begin VB.Label Label12 
         Caption         =   "Drop Down Selection"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   38
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.TextBox txtDriverFN 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3000
      TabIndex        =   1
      Top             =   1680
      Width           =   2655
   End
   Begin VB.TextBox txtDriverEmail 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3000
      TabIndex        =   4
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox txtDriverLN 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8040
      TabIndex        =   2
      Top             =   1680
      Width           =   2655
   End
   Begin VB.TextBox txtDriverId 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1200
      Width           =   1935
   End
   Begin VB.ComboBox lstComDriverBlood 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "frmDriverRegistration.frx":7968
      Left            =   8040
      List            =   "frmDriverRegistration.frx":7984
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2640
      Width           =   1695
   End
   Begin VB.ListBox lstDriverTrainedOn 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      ItemData        =   "frmDriverRegistration.frx":79C2
      Left            =   7920
      List            =   "frmDriverRegistration.frx":79D5
      Style           =   1  'Checkbox
      TabIndex        =   13
      Top             =   4920
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.ListBox lstDriverLicType 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      ItemData        =   "frmDriverRegistration.frx":7A28
      Left            =   3120
      List            =   "frmDriverRegistration.frx":7A38
      Style           =   1  'Checkbox
      TabIndex        =   11
      Top             =   4920
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSComCtl2.DTPicker dtPickDriverDateOfJoin 
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   2640
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   20578307
      CurrentDate     =   40827
   End
   Begin MSComCtl2.DTPicker dtPickDriverDOB 
      Height          =   375
      Left            =   8040
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   20578307
      CurrentDate     =   40827
   End
   Begin MSComCtl2.DTPicker dtPickDriverLicValidity 
      Height          =   375
      Left            =   8040
      TabIndex        =   9
      Top             =   3960
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   20578307
      CurrentDate     =   40827
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00004000&
      BorderStyle     =   4  'Dash-Dot
      X1              =   3360
      X2              =   8760
      Y1              =   8160
      Y2              =   8160
   End
   Begin VB.Line Line15 
      BorderColor     =   &H00004000&
      BorderStyle     =   4  'Dash-Dot
      X1              =   3360
      X2              =   8760
      Y1              =   9120
      Y2              =   9120
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00004000&
      BorderStyle     =   4  'Dash-Dot
      X1              =   3360
      X2              =   3360
      Y1              =   8160
      Y2              =   9120
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00004000&
      BorderStyle     =   4  'Dash-Dot
      X1              =   8760
      X2              =   8760
      Y1              =   8160
      Y2              =   9120
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000C&
      X1              =   720
      X2              =   10920
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Line Line11 
      BorderColor     =   &H8000000C&
      X1              =   720
      X2              =   10920
      Y1              =   7800
      Y2              =   7800
   End
   Begin VB.Line Line10 
      BorderColor     =   &H8000000C&
      X1              =   720
      X2              =   720
      Y1              =   7800
      Y2              =   5880
   End
   Begin VB.Line Line9 
      BorderColor     =   &H8000000C&
      X1              =   10920
      X2              =   10920
      Y1              =   5880
      Y2              =   7800
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000C&
      X1              =   720
      X2              =   10920
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000C&
      X1              =   720
      X2              =   10920
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000C&
      X1              =   720
      X2              =   720
      Y1              =   5280
      Y2              =   3720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      X1              =   10920
      X2              =   10920
      Y1              =   3720
      Y2              =   5280
   End
   Begin VB.Line Line5 
      BorderColor     =   &H8000000C&
      X1              =   720
      X2              =   10920
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line6 
      BorderColor     =   &H8000000C&
      X1              =   720
      X2              =   10920
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line7 
      BorderColor     =   &H8000000C&
      X1              =   720
      X2              =   720
      Y1              =   3240
      Y2              =   960
   End
   Begin VB.Line Line8 
      BorderColor     =   &H8000000C&
      X1              =   10920
      X2              =   10920
      Y1              =   960
      Y2              =   3240
   End
   Begin VB.Image picDriver 
      BorderStyle     =   1  'Fixed Single
      Height          =   2535
      Left            =   11640
      Stretch         =   -1  'True
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "State :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   375
      Left            =   960
      TabIndex        =   48
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Block / Sector :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   375
      Left            =   960
      TabIndex        =   47
      Top             =   6600
      Width           =   1935
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Pin No :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   255
      Left            =   6120
      TabIndex        =   46
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "City :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   255
      Left            =   6120
      TabIndex        =   45
      Top             =   6600
      Width           =   855
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Street :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   375
      Left            =   6120
      TabIndex        =   44
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "House No :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   375
      Left            =   960
      TabIndex        =   43
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Licence No :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   255
      Left            =   960
      TabIndex        =   36
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Licence Validity : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   255
      Left            =   6120
      TabIndex        =   35
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Type Of Licence :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   255
      Left            =   960
      TabIndex        =   34
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Trained On :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   255
      Left            =   6120
      TabIndex        =   33
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Email :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   375
      Left            =   960
      TabIndex        =   32
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   375
      Left            =   6120
      TabIndex        =   31
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   375
      Left            =   6120
      TabIndex        =   30
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   375
      Left            =   960
      TabIndex        =   29
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Driver ID  :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   375
      Left            =   960
      TabIndex        =   28
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Of Birth :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   375
      Left            =   6120
      TabIndex        =   27
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Blood Group :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   255
      Left            =   6120
      TabIndex        =   26
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Of Joining :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   375
      Left            =   960
      TabIndex        =   25
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Driver Registration"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   615
      Left            =   720
      TabIndex        =   24
      Top             =   120
      Width           =   10215
   End
End
Attribute VB_Name = "frmDriverRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mypicpath As String
Dim edit As Boolean
Dim myKey As String 'Primary key to identify the editing row

Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim rsLicType As ADODB.Recordset
Dim rsTrainedOn As ADODB.Recordset

Dim filepath As String
Dim filename As String
Dim picpath As String

Dim focA As Boolean
Dim focB As Boolean

Dim flgUploader As Boolean 'this flag is used to control the frmUploader

Dim flgB As Boolean
Dim flgA As Boolean


Private Sub cmdCancel_Click()

End Sub

Private Sub cmdDriverRegClose_Click()
Unload Me
Call CenterMe(frmDriverList)
frmDriverList.Show
End Sub

Private Sub cmdDriverRegSave_Click()

'***********************************
'***SAVING / EDITING DRIVER TABLE***
'***********************************


'MsgBox "myPicPath: " & mypicpath
'MsgBox "picpath: " & picpath

Dim i As Integer
Dim X(4) As String 'Licence Type array
Dim Y(5) As String 'Trained On array
Dim licVal1, licVal2 As Long

Dim currLicNo, updatedLicNo As String
Dim SQLText As String

Dim myDateOfJoin, myDateOfBirth, myLicValidity As Date


If (picpath = "" And edit = False) Then
    picpath = "C:\Documents and Settings\Binaya\Desktop\BlueLineTake6\Drivers\unavailable.jpeg"
End If

If (Me.txtDriverFN.Text = "" Or Me.txtDriverLN.Text = "") Then
    MsgBox "Enter the driver's Full Name"
    Me.txtDriverFN.SetFocus
    Exit Sub
End If

If (Me.txtDriverEmail = "") Then
    MsgBox "Enter the Email Address"
    Exit Sub
End If

If (InStr(1, Me.txtDriverEmail.Text, "@", vbTextCompare) = 0) Then
    MsgBox "Invalid Email Address"
    Me.txtDriverEmail.Text = ""
    Me.txtDriverEmail.SetFocus
    Exit Sub
End If


If (Me.lstComDriverBlood.Text = "") Then
    MsgBox "Enter the Blood Group"
    Me.lstComDriverBlood.SetFocus
    Exit Sub
End If


If (Me.txtDriverHouseNo.Text = "" Or Me.txtDriverStreet.Text = "" Or Me.txtDriverCity.Text = "" Or Me.txtDriverState.Text = "" Or Me.txtDriverBlock = "" Or Val(Me.txtDriverPin) = 0) Then
    Status = MsgBox("Please fill out the address details.", vbOKOnly, "Incomplete Address Info")
    Me.txtDriverHouseNo.SetFocus
    Exit Sub
End If

If (Len(Me.txtDriverPin.ClipText) < 6) Then
    MsgBox "Invalid Pin Code"
    Me.txtDriverPin.SetFocus
    Exit Sub
End If

If IsNumeric(Me.txtDriverHouseNo.Text) = False Then
    MsgBox ("Please enter a number for House No")
    txtDriverHouseNo.Text = ""
    txtDriverHouseNo.SetFocus
    Exit Sub
End If

If (Val(Me.txtDriverPhNo.Text) = 0) Then
    MsgBox "Enter Customer Phone No"
    Me.txtDriverPhNo.SetFocus
    Exit Sub
End If


If (Len(Me.txtDriverPhNo.ClipText) < 10) Then
    MsgBox "Invalid Phone No"
    Me.txtDriverPhNo.SetFocus
    Exit Sub
End If

If (Len(Me.txtDriverLicNo.ClipText) < 6) Then
    MsgBox "Invalid Licence No"
    txtDriverLicNo.Text = ""
    Me.txtDriverLicNo.SetFocus
    Exit Sub
End If



'Verifying the licence type selection
  
If ((lstDriverLicType.Selected(0) = False) And (lstDriverLicType.Selected(1) = False) _
And (lstDriverLicType.Selected(2) = False) And (lstDriverLicType.Selected(3) = False)) Then
    MsgBox "Enter the Licence Type"
    Me.cmdLicType.SetFocus
    Exit Sub
End If

If (edit = False) Then

    SQLText = "SELECT * FROM driver"
    rs.Open SQLText, cn, adOpenKeyset
    licVal1 = Val(Me.txtDriverLicNo.Text)

    Do While (rs.EOF = False)
        licVal2 = Val(rs("DRIVER_LICENCE_NO"))
        If (licVal1 = licVal2) Then
            MsgBox "Licence No already Exists"
            rs.Close
            Exit Sub
        End If
        rs.MoveNext
    Loop
rs.Close
End If


If (edit = True) Then
    SQLText = "SELECT driver_licence_no FROM driver where driver_id = " & myKey
    'MsgBox SQLText
    
    rs.Open SQLText, cn, adOpenKeyset
    currLicNo = rs("DRIVER_LICENCE_NO")
    updatedLicNo = Me.txtDriverLicNo.Text
    rs.Close
    
    'MsgBox "Current Licence No: " & currLicNo & "New: " & updatedLicNo
    
    SQLText = "SELECT driver_licence_no FROM driver"
    rs.Open SQLText
            
    If (currLicNo <> updatedLicNo) Then
        Do While (rs.EOF = False)
            If (updatedLicNo = rs("DRIVER_LICENCE_NO")) Then
                MsgBox "Licence No Already Exists"
                rs.Close
                Exit Sub
            End If
            rs.MoveNext
        Loop
    End If
    rs.Close
    
End If


'***Date of Join should be not be greter than todays date
'***Date of birth difference with now should be atleast 19 years
'***Licence Validity date should be active


myDateOfJoin = Format(Me.dtPickDriverDateOfJoin.Value, "MM-dd-yyyy")


If (DateDiff("n", Now, myDateOfJoin) >= 0) Then
    MsgBox "Invalid Joining Date"
    Me.dtPickDriverDateOfJoin.Value = Now
    Me.dtPickDriverDateOfJoin.SetFocus
    Exit Sub
End If


myDateOfBirth = Format(Me.dtPickDriverDOB.Value, "MM-dd-yyyy")

If (DateDiff("y", myDayeOfBirth, Now) < 18) Then
    MsgBox "Driver Not Qualified To Driver"
    Me.dtPickDriverDOB.Value = Now
    Me.dtPickDriverDateOfJoin.SetFocus
    Exit Sub
End If


myLicValidity = Format(Me.dtPickDriverLicValidity.Value, "MM-dd-yyyy")

If (DateDiff("n", Now, myLicValidity) <= 0) Then
    MsgBox "Licence Has Expired"
    Me.dtPickDriverLicValidity.Value = Now
    Me.dtPickDriverLicValidity.SetFocus
    Exit Sub
End If


'Filling up the licence type array x()
For i = 0 To Me.lstDriverLicType.ListCount - 1
    If (Me.lstDriverLicType.Selected(i)) Then
        X(i) = "Y"
    Else
        X(i) = "N"
    End If
Next

'Filling up the trained on array y()
For i = 0 To Me.lstDriverTrainedOn.ListCount - 1
    If (Me.lstDriverTrainedOn.Selected(i)) Then
        Y(i) = "Y"
    Else
        Y(i) = "N"
    End If
Next

If (edit = False) Then
    
    SQLTextA = "insert into driver values(" _
    & "" & Val(Me.txtDriverId.Text) & "," _
    & "'" & Trim(Me.txtDriverFN.Text) & "', " _
    & "'" & Trim(Me.txtDriverLN.Text) & "', " _
    & "to_date('" & Format(Me.dtPickDriverDOB.Value, "dd-mm-yyyy") & "', 'dd-mm-yyyy'), " _
    & "'" & Trim(Me.txtDriverEmail.Text) & "'," _
    & "" & Val(Me.txtDriverPhNo.Text) & ", " _
    & "to_date('" & Format(Me.dtPickDriverDateOfJoin.Value, "dd-mm-yyyy") & "', 'dd-mm-yyyy'), " _
    & "'" & Me.lstComDriverBlood.Text & "', " _
    & "" & Val(Me.txtDriverLicNo.Text) & ", " _
    & "to_date('" & Format(Me.dtPickDriverLicValidity.Value, "dd-mm-yyyy") & "', 'dd-mm-yyyy'), " _
    & "" & Val(Me.txtDriverHouseNo.Text) & ", " _
    & "'" & Trim(Me.txtDriverStreet.Text) & "', " _
    & "'" & Trim(Me.txtDriverBlock.Text) & "', " _
    & "'" & Trim(Me.txtDriverCity.Text) & "', " _
    & "'" & Trim(Me.txtDriverState.Text) & "', " _
    & "" & Val(Me.txtDriverPin.Text) & ", " _
    & "'" & picpath & "', " _
    & "'False')"
    
    
    SQLTextB = "insert into licencetype values (" & Val(Me.txtDriverId.Text) & ", '" & X(0) & "', '" & X(1) & "', '" & X(2) & "', '" & X(3) & "' )"
    SQLTextC = "insert into trainedon values (" & Val(Me.txtDriverId.Text) & ", '" & Y(0) & "', '" & Y(1) & "', '" & Y(2) & "', '" & Y(3) & "', '" & Y(4) & "' )"

    
    'MsgBox SQLTextA

    cn.Execute SQLTextA
    cn.Execute SQLTextB
    cn.Execute SQLTextC
    
    MsgBox "Driver Details Saved - Successful"

Else
    
    SQLTextD = "update driver set " _
    & "DRIVER_ID = " & Val(Me.txtDriverId.Text) & ", " _
    & "DRIVER_FIRST_NAME = '" & Trim(Me.txtDriverFN.Text) & "', " _
    & "DRIVER_LAST_NAME = '" & Trim(Me.txtDriverLN.Text) & "', " _
    & "DRIVER_DOB = to_date('" & Format(Me.dtPickDriverDOB.Value, "dd-mm-yyyy") & "', 'dd-mm-yyyy'), " _
    & "DRIVER_EMAIL = '" & Trim(Me.txtDriverEmail.Text) & "', " _
    & "DRIVER_PHONE = " & Val(Me.txtDriverPhNo.Text) & ", " _
    & "DRIVER_DATE_OF_JOIN = to_date('" & Format(Me.dtPickDriverDateOfJoin.Value, "dd-mm-yyyy") & "', 'dd-mm-yyyy'), " _
    & "DRIVER_BLOOD_GROUP = '" & Trim(Me.lstComDriverBlood.Text) & "', " _
    & "DRIVER_LICENCE_NO = " & Val(Me.txtDriverLicNo.Text) & ", " _
    & "DRIVER_LICENCE_VALIDITY = to_date('" & Format(Me.dtPickDriverLicValidity.Value, "dd-mm-yyyy") & "', 'dd-mm-yyyy')," _
    & "DRIVER_HOUSE_NO = " & Val(Me.txtDriverHouseNo.Text) & ", " _
    & "DRIVER_STREET = '" & Trim(Me.txtDriverStreet.Text) & "', " _
    & "DRIVER_BLOCK = '" & Trim(Me.txtDriverBlock.Text) & "', " _
    & "DRIVER_CITY = '" & Trim(Me.txtDriverCity.Text) & "', " _
    & "DRIVER_STATE = '" & Trim(Me.txtDriverState.Text) & "', " _
    & "DRIVER_PIN_NO = " & Val(Me.txtDriverPin.Text) & ", " _
    & "DRIVER_PIC = '" & picpath & "' " _
    & "where driver_id = " & Val(myKey) & ""
   
   
    SQLTextE = "update licencetype set " _
    & "HEAVY_MOTOR_VEICHLE = '" & X(0) & "', " _
    & "LIGHT_MOTOR_VEICHLE = '" & X(1) & "', " _
    & "CARGO = '" & X(2) & "', " _
    & "PUBLIC_TRANSPORT = '" & X(3) & "' " _
    & "where driver_id = " & Val(myKey) & ""

    
    SQLTextF = "update trainedon set " _
    & "NIGHT_DRIVING = '" & Y(0) & "', " _
    & "DAY_DRIVING = '" & Y(1) & "', " _
    & "LONG_DISTANCE = '" & Y(2) & "', " _
    & "SHORT_DISTANCE = '" & Y(3) & "', " _
    & "INTER_STATE = '" & Y(4) & "' " _
    & "where driver_id = " & Val(myKey) & ""
    
    cn.Execute SQLTextD
    cn.Execute SQLTextE
    cn.Execute SQLTextF
    
  
    MsgBox "Driver Details Updated - Successful"
End If


cn.Close
    
Unload Me
Call CenterMe(frmDriverList)
frmDriverList.Show

End Sub

Private Sub cmdLicType_Click()

If (flgA) Then
    Me.lstDriverLicType.Visible = True
    Me.lstDriverLicType.SetFocus
    flgA = False
    focA = False
    If (flgB = False) Then
        Me.lstDriverTrainedOn.Visible = False
        flgB = True
        focB = False
    End If
       
Else
    Me.lstDriverLicType.Visible = False
    Call init_flg

End If

End Sub

Private Sub cmdBrowsePlus_Click()

'**************************************
'***Calls the frmUploader form, sets***
'***the picture to the picture box*****
'**************************************

mypicpath = frmUploader.lblPickpath.Caption
If (mypicpath = "xxxx") Then
    flgUploader = False
    Unload frmUploader
End If

If (flgUploader = False) Then
    frmDriverRegistration.Enabled = False
    Call frmUploader.setWhoCalled("driver")
    frmUploader.Top = 4080
    frmUploader.Left = 11040
    frmUploader.Show
    flgUploader = True
Else
    flgUploader = False
    'mypicpath = frmUploader.lblPickpath.Caption
    Unload frmUploader
    
    On Error GoTo skip
    If (mypicpath <> "") Then
        Me.picDriver.Picture = LoadPicture(mypicpath)
        picpath = mypicpath
    Else
        GoTo skip
    End If
    Exit Sub
skip:
    MsgBox "Invalid Picture Format"
End If
End Sub

Private Sub cmdTrainedOn_Click()

If (flgB) Then
    Me.lstDriverTrainedOn.Visible = True
    Me.lstDriverTrainedOn.SetFocus
    flgB = False
    focB = False
    If (flgA = False) Then
        Me.lstDriverLicType.Visible = False
        flgA = True
        focA = False
    End If
Else
    Me.lstDriverTrainedOn.Visible = False
    Call init_flg

End If

End Sub


Private Sub Form_Click()
If (flgA = False) Then
    flgA = True
    Me.lstDriverLicType.Visible = False
End If

If (flgB = False) Then
    flgB = True
    Me.lstDriverTrainedOn.Visible = False
End If


End Sub


Private Sub lstDriverLicType_LostFocus()
If (flgA = False) And (focA = False) Then
    flgA = False
    Me.lstDriverLicType.Visible = False
End If
End Sub


Private Sub lstDriverTrainedOn_LostFocus()
If (flgB = False) And (focB = False) Then
    flgB = False
    Me.lstDriverTrainedOn.Visible = False
End If
  
End Sub

Private Sub Form_Load()
Call init_flg
edit = False

flgUploader = False 'Boolean to control the frmUploader

Me.dtPickDriverDateOfJoin.Value = Now
Me.dtPickDriverDOB.Value = Now
Me.dtPickDriverLicValidity.Value = Now

Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset

cn.Open "provider=MSDASQL.1;DRIVER=MICROSOFT ODBC FOR ORACLE;UID=SCOTT;PWD=TIGER"
IdGenerate

Me.Visible = True
Me.txtDriverFN.SetFocus


End Sub


Private Sub init_flg()
flgA = True
flgB = True

focA = False
focB = False

End Sub

Private Sub IdGenerate()

'***********************
'**Driver Id Generator**
'***********************

Dim rscustID As ADODB.Recordset
Set rsdriverID = New ADODB.Recordset
rsdriverID.Open "select * from driver order by driver_id desc", cn, adOpenKeyset
If rsdriverID.EOF = False Then
    Me.txtDriverId = rsdriverID("driver_id") + 1
Else
    Me.txtDriverId.Text = 50101
End If
       
End Sub

Public Sub FillDriverForm(key As String)

'***************************
'**Driver Form Fill Method**
'***************************

Dim i As Integer

myKey = key
edit = True
Me.cmdDriverRegSave.Caption = "Update"
Me.txtDriverId.Locked = True

SQLTextA = "Select * from driver where(driver_id = " & key & ")"
SQLTextB = "Select * from licencetype where(driver_id = " & key & ")"
SQLTextC = "Select * from trainedon where(driver_id = " & key & ")"

Set rsLicType = New ADODB.Recordset
Set rsTrainedOn = New ADODB.Recordset


rs.Open SQLTextA, cn, adOpenKeyset
rsLicType.Open SQLTextB, cn, adOpenKeyset
rsTrainedOn.Open SQLTextC, cn, adOpenKeyset


With Me
    .txtDriverId = rs("driver_id")
    .txtDriverFN = rs("driver_first_name")
    .txtDriverLN = rs("driver_last_name")
    .txtDriverEmail = rs("driver_email")
    .txtDriverPhNo = rs("driver_phone")
    .txtDriverLicNo = rs("driver_licence_no")
    .txtDriverHouseNo = rs("driver_house_no")
    .txtDriverBlock = rs("driver_block")
    .txtDriverStreet = rs("driver_street")
    .txtDriverCity = rs("driver_city")
    .txtDriverState = rs("driver_state")
    .txtDriverPin = rs("driver_pin_no")
    
    picpath = rs("driver_pic")
    
    On Error GoTo skip
    .picDriver.Picture = LoadPicture(rs("driver_pic"))
skip:
    .dtPickDriverDOB.Value = Format(rs("driver_dob"), "dd-mm-yyyy")
    .dtPickDriverDateOfJoin.Value = Format(rs("driver_date_of_join"), "dd-mm-yyyy")
    .dtPickDriverLicValidity.Value = Format(rs("driver_licence_validity"), "dd-mm-yyyy")

    .lstComDriverBlood.Text = rs("driver_blood_group")
    
    
    'Filling the Licence Type List Box
    
    If (rsLicType("HEAVY_MOTOR_VEICHLE") = "Y") Then
         .lstDriverLicType.Selected(0) = True
    End If
    
    If (rsLicType("LIGHT_MOTOR_VEICHLE") = "Y") Then
         .lstDriverLicType.Selected(1) = True
    End If
    
    If (rsLicType("CARGO") = "Y") Then
         .lstDriverLicType.Selected(2) = True
    End If
   
    If (rsLicType("PUBLIC_TRANSPORT") = "Y") Then
         .lstDriverLicType.Selected(3) = True
    End If
   
   
    'Filling the Trained On List Box
   
    If (rsTrainedOn("NIGHT_DRIVING") = "Y") Then
         .lstDriverTrainedOn.Selected(0) = True
    End If
   
    If (rsTrainedOn("DAY_DRIVING") = "Y") Then
         .lstDriverTrainedOn.Selected(1) = True
    End If
    
    If (rsTrainedOn("LONG_DISTANCE") = "Y") Then
        .lstDriverTrainedOn.Selected(2) = True
    End If
    
    If (rsTrainedOn("SHORT_DISTANCE") = "Y") Then
         .lstDriverTrainedOn.Selected(3) = True
    End If
    
    If (rsTrainedOn("INTER_STATE") = "Y") Then
         .lstDriverTrainedOn.Selected(4) = True
    End If
    
    
End With

rs.Close
rsLicType.Close
rsTrainedOn.Close

End Sub

