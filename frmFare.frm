VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFare 
   BorderStyle     =   0  'None
   ClientHeight    =   11190
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   15285
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmFare.frx":0000
   ScaleHeight     =   11190
   ScaleWidth      =   15285
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Add"
      Height          =   735
      Left            =   7440
      Picture         =   "frmFare.frx":58FC
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton cmdResfresh 
      Caption         =   "&Refresh"
      Height          =   735
      Left            =   5400
      Picture         =   "frmFare.frx":65C6
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Go &Back"
      Height          =   735
      Left            =   9360
      Picture         =   "frmFare.frx":6B50
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7200
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpickFareValid 
      Height          =   375
      Left            =   9360
      TabIndex        =   3
      Top             =   2520
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd/MMM/yyyy"
      Format          =   66322435
      CurrentDate     =   40842
   End
   Begin VB.TextBox txtAbove15 
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
      Left            =   9960
      TabIndex        =   9
      Top             =   5760
      Width           =   1335
   End
   Begin VB.TextBox txtUpto15 
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
      Left            =   9960
      TabIndex        =   7
      Top             =   5280
      Width           =   1335
   End
   Begin VB.TextBox txtAc 
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
      Left            =   5280
      TabIndex        =   6
      Top             =   5160
      Width           =   1335
   End
   Begin VB.TextBox txtNightService 
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
      Left            =   5280
      TabIndex        =   8
      Top             =   5760
      Width           =   1335
   End
   Begin VB.TextBox txtAboveHundred 
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
      Left            =   6480
      TabIndex        =   5
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox txtFareId 
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
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox txtRatePerKm 
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
      Left            =   6480
      TabIndex        =   4
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox txtMinCharge 
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
      Left            =   6480
      TabIndex        =   2
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Line Line8 
      BorderColor     =   &H8000000C&
      X1              =   12480
      X2              =   12480
      Y1              =   4800
      Y2              =   6360
   End
   Begin VB.Line Line7 
      BorderColor     =   &H8000000C&
      X1              =   3120
      X2              =   12480
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line Line6 
      BorderColor     =   &H8000000C&
      X1              =   3120
      X2              =   3120
      Y1              =   4800
      Y2              =   6360
   End
   Begin VB.Line Line5 
      BorderColor     =   &H8000000C&
      X1              =   3120
      X2              =   12480
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000C&
      X1              =   11400
      X2              =   11400
      Y1              =   1440
      Y2              =   3720
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000C&
      X1              =   4560
      X2              =   11400
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000C&
      X1              =   4560
      X2              =   4560
      Y1              =   1440
      Y2              =   3720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      X1              =   4560
      X2              =   11400
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00004000&
      BorderStyle     =   4  'Dash-Dot
      X1              =   10680
      X2              =   10680
      Y1              =   7080
      Y2              =   8040
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00004000&
      BorderStyle     =   4  'Dash-Dot
      X1              =   5280
      X2              =   5280
      Y1              =   7080
      Y2              =   8040
   End
   Begin VB.Line Line15 
      BorderColor     =   &H00004000&
      BorderStyle     =   4  'Dash-Dot
      X1              =   5280
      X2              =   10680
      Y1              =   8040
      Y2              =   8040
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00004000&
      BorderStyle     =   4  'Dash-Dot
      X1              =   5280
      X2              =   10680
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Above 15 min :"
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
      Left            =   8280
      TabIndex        =   22
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Upto 15 min :"
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
      Left            =   8280
      TabIndex        =   21
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Per 5 min"
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
      Left            =   11400
      TabIndex        =   20
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "A/C :"
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
      Left            =   3360
      TabIndex        =   19
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Night Service :"
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
      Left            =   3360
      TabIndex        =   18
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Above 100 Km :"
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
      Left            =   4800
      TabIndex        =   17
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Rate Per Km :"
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
      Left            =   4800
      TabIndex        =   16
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Fare ID :"
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
      Left            =   4800
      TabIndex        =   15
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Min Charge :"
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
      Left            =   4800
      TabIndex        =   14
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Valid Till :"
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
      Left            =   8160
      TabIndex        =   13
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fare Details"
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
      Left            =   4560
      TabIndex        =   0
      Top             =   600
      Width           =   6855
   End
End
Attribute VB_Name = "frmFare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim myKey As String
Dim edit As Boolean
Dim mcon As ADODB.Connection

Private Sub cmdClose_Click()
Unload Me
Call CenterMe(frmFareList)
frmFareList.Show
End Sub

Private Sub cmdSave_Click()

'*****************************
'***SAVE / EDIT Fare Record***
'*****************************

Dim fareValid As Date

If IsNumeric(Me.txtMinCharge.Text) = False Then
    MsgBox ("Please enter a number for Minimum Charge")
    txtMinCharge.Text = ""
    txtMinCharge.SetFocus
    Exit Sub
End If

If IsNumeric(Me.txtRatePerKm.Text) = False Then
    MsgBox ("Please enter a number for Rate Per Km")
    txtRatePerKm.Text = ""
    txtRatePerKm.SetFocus
    Exit Sub
End If


If IsNumeric(Me.txtAboveHundred.Text) = False Then
    MsgBox ("Please enter a number for Above 100Km Charge")
    txtAboveHundred.Text = ""
    txtAboveHundred.SetFocus
    Exit Sub
End If


If IsNumeric(Me.txtAc.Text) = False Then
    MsgBox ("Please enter a number for A/C Charges")
    txtAc.Text = ""
    txtAc.SetFocus
    Exit Sub
End If


If IsNumeric(Me.txtNightService.Text) = False Then
    MsgBox ("Please enter a number for Night Charges")
    txtNightService.Text = ""
    txtNightService.SetFocus
    Exit Sub
End If


If IsNumeric(Me.txtUpto15.Text) = False Then
    MsgBox ("Please enter a number for Upto 15 min Charge")
    txtUpto15.Text = ""
    txtUpto15.SetFocus
    Exit Sub
End If

If IsNumeric(Me.txtAbove15.Text) = False Then
    MsgBox ("Please enter a number above 15 min Charge")
    txtAbove15.Text = ""
    txtAbove15.SetFocus
    Exit Sub
End If


fareValid = Format(Me.dtpickFareValid.Value, "MM-dd-yyyy")

If (DateDiff("n", Now, fareValid) <= 0) Then
    MsgBox "Invalid Fare Validity"
    Me.dtpickFareValid.Value = Now
    Me.dtpickFareValid.SetFocus
    Exit Sub
End If


If (edit = False) Then
    SQLTextA = "insert into fare values(" _
    & "" & Val(Me.txtFareId) & ", " _
    & "" & Val(Me.txtMinCharge.Text) & ", " _
    & "" & Val(Me.txtRatePerKm.Text) & ", " _
    & "" & Val(Me.txtAboveHundred.Text) & ", " _
    & "" & Val(Me.txtAc.Text) & ", " _
    & "" & Val(Me.txtNightService.Text) & ", " _
    & "to_date('" & Format(Me.dtpickFareValid.Value, "dd-mm-yyyy") & "', 'dd-mm-yyyy')," _
    & "" & Val(Me.txtAc.Text) & ", " _
    & "" & Val(Me.txtAc.Text) & ")"

        
    'MsgBox SQLTextA
    mcon.Execute SQLTextA
Else
    SQLTextB = "update fare set " _
    & "FARE_ID = " & Val(Me.txtFareId.Text) & ", " _
    & "MIN_CHARGE = " & Val(Me.txtMinCharge.Text) & ", " _
    & "PER_KM_CHARGE = " & Val(Me.txtRatePerKm.Text) & ", " _
    & "ABOVE_HUNDRED_CHARGE = " & Val(Me.txtAboveHundred.Text) & ", " _
    & "AC_CHARGE = " & Val(Me.txtAc) & ", " _
    & "NIGHT_CHARGE = " & Val(Me.txtNightService.Text) & ", " _
    & "FARE_VALIDITY = to_date('" & Format(Me.dtpickFareValid.Value, "dd-mm-yyyy") & "', 'dd-mm-yyyy'), " _
    & "UPTO_15 = " & Val(Me.txtUpto15.Text) & ", " _
    & "ABOVE_15 = " & Val(Me.txtAbove15.Text) & " " _
    & "WHERE FARE_ID = " & Val(myKey) & ""

    'MsgBox SQLTextB
    mcon.Execute SQLTextB
End If

Unload Me
Call CenterMe(frmFareList)
frmFareList.Show
End Sub

Private Sub Form_Load()

'*********************
'***Form Load Event***
'*********************

edit = False
Me.dtpickFareValid.Value = Now

Set mcon = New ADODB.Connection
mcon.Open "provider=MSDASQL.1;DRIVER=MICROSOFT ODBC FOR ORACLE;UID=SCOTT;PWD=TIGER"
Call FareIdGenerate

Me.Visible = True
Me.txtMinCharge.SetFocus
End Sub

Private Sub FareIdGenerate()

'**********************
'***Generate Fare Id***
'**********************

Dim rsfareID As ADODB.Recordset
Set rsfareID = New ADODB.Recordset
rsfareID.Open "select * from fare order by fare_id desc", mcon, adOpenKeyset
If rsfareID.EOF = False Then
    Me.txtFareId = rsfareID("fare_id") + 1
Else
    Me.txtFareId.Text = 10200
End If
End Sub

Public Sub FillFareForm(key As String)

'***************************
'***Fare Form Fill Method***
'***************************

Dim i As Integer
Dim rs As ADODB.Recordset

myKey = key
edit = True
Me.cmdSave.Caption = "Update"
Me.txtFareId.Locked = True

SQLTextA = "Select * from FARE where (FARE_ID = " & key & ")"




Set rs = New ADODB.Recordset

rs.Open SQLTextA, mcon, adOpenKeyset

With Me
    .txtFareId = rs("FARE_ID")
    .txtMinCharge = rs("MIN_CHARGE")
    .txtRatePerKm = rs("PER_KM_CHARGE")
    .txtAboveHundred = rs("ABOVE_HUNDRED_CHARGE")
    .txtAc = rs("AC_CHARGE")
    .txtNightService = rs("NIGHT_CHARGE")
    .dtpickFareValid = rs("FARE_VALIDITY")
    .txtUpto15 = rs("UPTO_15")
    .txtAbove15 = rs("ABOVE_15")
End With

rs.Close

End Sub


