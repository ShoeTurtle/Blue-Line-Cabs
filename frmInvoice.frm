VERSION 5.00
Begin VB.Form frmInvoice 
   BorderStyle     =   0  'None
   Caption         =   "Street List - Blue Line"
   ClientHeight    =   11190
   ClientLeft      =   300
   ClientTop       =   -75
   ClientWidth     =   15285
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmInvoice.frx":0000
   ScaleHeight     =   11190
   ScaleMode       =   0  'User
   ScaleWidth      =   18833.63
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtMin 
      Height          =   285
      Left            =   10320
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox txtNight 
      Height          =   285
      Left            =   10320
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox txtAc 
      Height          =   285
      Left            =   10320
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox txtDist 
      Height          =   285
      Left            =   10320
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox txtWait 
      Height          =   285
      Left            =   10320
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   2280
      Width           =   1335
   End
   Begin VB.ComboBox comboBookingId 
      Height          =   315
      Left            =   4440
      TabIndex        =   1
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txtCustName 
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   2880
      Width           =   3135
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Go &Back"
      Height          =   735
      Left            =   9960
      Picture         =   "frmInvoice.frx":58FC
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   735
      Left            =   3240
      Picture         =   "frmInvoice.frx":5A46
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Add"
      Height          =   735
      Left            =   4920
      Picture         =   "frmInvoice.frx":5FD0
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7680
      Width           =   1215
   End
   Begin VB.TextBox txtSource 
      Height          =   1005
      Left            =   3600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   4800
      Width           =   2775
   End
   Begin VB.TextBox txtDestination 
      Height          =   1005
      Left            =   8640
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   4800
      Width           =   2655
   End
   Begin VB.TextBox txtDistance 
      Height          =   285
      Left            =   8640
      TabIndex        =   8
      Top             =   6000
      Width           =   1575
   End
   Begin VB.TextBox txtTotalAmount 
      Height          =   285
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   6480
      Width           =   1575
   End
   Begin VB.ComboBox comboRate 
      Height          =   315
      Left            =   3600
      TabIndex        =   7
      Top             =   6000
      Width           =   1575
   End
   Begin VB.TextBox txtWaitTime 
      Height          =   285
      Left            =   3600
      TabIndex        =   9
      Top             =   6480
      Width           =   1575
   End
   Begin VB.TextBox txtInvoiceNo 
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox txtBookTime 
      Height          =   285
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox txtBookDate 
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "&Calculate"
      Height          =   735
      Left            =   8280
      Picture         =   "frmInvoice.frx":6C9A
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   735
      Left            =   6600
      Picture         =   "frmInvoice.frx":6DE4
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Min' Charge :"
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
      Left            =   8880
      TabIndex        =   40
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "min'"
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
      Left            =   5280
      TabIndex        =   39
      Top             =   6480
      Width           =   495
   End
   Begin VB.Line Line17 
      BorderColor     =   &H8000000C&
      X1              =   10645.9
      X2              =   14638.11
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Charge Sheet"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   375
      Left            =   8640
      TabIndex        =   38
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Line Line16 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      X1              =   10645.9
      X2              =   14638.11
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line15 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      X1              =   10645.9
      X2              =   10645.9
      Y1              =   1320
      Y2              =   4080
   End
   Begin VB.Line Line14 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      X1              =   14638.11
      X2              =   14638.11
      Y1              =   1320
      Y2              =   4080
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      X1              =   10645.9
      X2              =   14638.11
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Night Charge:"
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
      Left            =   8880
      TabIndex        =   37
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "A/C Charge"
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
      Left            =   8880
      TabIndex        =   36
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Dist Charge"
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
      Left            =   8880
      TabIndex        =   35
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Wait Charge :"
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
      Left            =   8880
      TabIndex        =   34
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00004000&
      BorderStyle     =   4  'Dash-Dot
      X1              =   3844.352
      X2              =   13898.81
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00004000&
      BorderStyle     =   4  'Dash-Dot
      X1              =   3844.352
      X2              =   13898.81
      Y1              =   8520
      Y2              =   8520
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00004000&
      BorderStyle     =   4  'Dash-Dot
      X1              =   3844.352
      X2              =   3844.352
      Y1              =   7560
      Y2              =   8520
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00004000&
      BorderStyle     =   4  'Dash-Dot
      X1              =   13898.81
      X2              =   13898.81
      Y1              =   7560
      Y2              =   8520
   End
   Begin VB.Line Line8 
      BorderColor     =   &H8000000C&
      X1              =   2661.475
      X2              =   2661.475
      Y1              =   4440
      Y2              =   7080
   End
   Begin VB.Line Line7 
      BorderColor     =   &H8000000C&
      X1              =   14638.11
      X2              =   2661.475
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Line Line6 
      BorderColor     =   &H8000000C&
      X1              =   14638.11
      X2              =   14638.11
      Y1              =   4440
      Y2              =   7080
   End
   Begin VB.Line Line5 
      BorderColor     =   &H8000000C&
      X1              =   2661.475
      X2              =   14638.11
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000C&
      X1              =   10350.18
      X2              =   10350.18
      Y1              =   1320
      Y2              =   4080
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000C&
      X1              =   2661.475
      X2              =   10350.18
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000C&
      X1              =   2661.475
      X2              =   2661.475
      Y1              =   1320
      Y2              =   4080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      X1              =   2661.475
      X2              =   10350.18
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Rs."
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
      Left            =   10320
      TabIndex        =   33
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Source :"
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
      Left            =   2400
      TabIndex        =   32
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Destination :"
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
      Left            =   6960
      TabIndex        =   31
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Rate :"
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
      Left            =   2400
      TabIndex        =   30
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Distance :"
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
      Left            =   6960
      TabIndex        =   29
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Wait Time :"
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
      Left            =   2400
      TabIndex        =   28
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount :"
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
      Left            =   6960
      TabIndex        =   27
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name :"
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
      Left            =   2400
      TabIndex        =   26
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Number :"
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
      Left            =   2400
      TabIndex        =   25
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Time :"
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
      Left            =   5640
      TabIndex        =   24
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Date :"
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
      Left            =   2400
      TabIndex        =   23
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Booking ID :"
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
      Left            =   2400
      TabIndex        =   22
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Generator"
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
      Left            =   2160
      TabIndex        =   21
      Top             =   480
      Width           =   9735
   End
End
Attribute VB_Name = "frmInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim myInvoiceId As String
Dim edit As Boolean

Dim rs As ADODB.Recordset
Dim rsTmp As ADODB.Recordset

Dim myBookId As String
Dim myCustId As String
Dim myFareId As String
Dim mytime As Integer

Dim tmin, tdist, tac, tnight, twait As Double


Dim ac, night As Boolean
Dim dist, waitTime As Double

Dim total As Long

Dim mcon As ADODB.Connection

Private Sub cmdCalculate_Click()

If (Me.comboBookingId.Text = "") Then
    MsgBox "Select a Booking Id"
    comboBookingId.SetFocus
    Exit Sub
End If

Call Calculate_Fare

End Sub

Private Sub cmdClose_Click()
Unload Me
Call CenterMe(frmInvoiceList)
frmInvoiceList.Show
End Sub

Private Sub cmdSave_Click()

If (Me.comboBookingId.Text = "") Then
    MsgBox "Select a Booking Id"
    comboBookingId.SetFocus
    Exit Sub
End If

If IsNumeric(Me.txtWaitTime) = False Then
    MsgBox ("Please enter a number for Wait Time")
    txtWaitTime.Text = ""
    txtWaitTime.SetFocus
    Exit Sub
End If

If IsNumeric(Me.txtDistance) = False Then
    MsgBox ("Please enter a number for Distance")
    txtDistance.Text = ""
    txtDistance.SetFocus
    Exit Sub
End If

If (Me.txtTotalAmount = "") Then
    MsgBox "Calcualte the toal amount to save"
    Me.cmdCalculate.SetFocus
    Exit Sub
End If

If (edit = False) Then
    
    SQLText = "INSERT INTO invoice VALUES(" _
    & Val(txtInvoiceNo.Text) & ", " _
    & Val(comboBookingId.Text) & ", " _
    & Val(comboRate.Text) & ", " _
    & "to_date('" & Format(Now, "dd-mm-yyyy") & "', 'dd-mm-yyyy')," _
    & Val(txtDistance.Text) & ", " _
    & Val(txtWaitTime.Text) & ", " _
    & Val(txtTotalAmount.Text) & ")"
            
    'MsgBox SQLText
    mcon.Execute SQLText
    
    SQLText = "UPDATE trip SET invoice = 'True' where trip_id = " & myBookId
    'MsgBox SQLText
    mcon.Execute SQLText
    
    
Else
    Call Calculate_Fare
    SQLTextB = "update invoice set " _
    & "FARE_REF = " & Val(Me.comboRate.Text) & ", " _
    & "DISTANCE = " & Val(Me.txtDistance.Text) & ", " _
    & "WAIT_TIME = " & Val(Me.txtWaitTime.Text) & ", " _
    & "TOTAL_AMOUNT = " & Val(Me.txtTotalAmount.Text) & _
    " WHERE invoice_id = " & myInvoiceId
    'MsgBox SQLTextB
    mcon.Execute SQLTextB
End If



'***Enter Values into Charge Sheet*****

If (edit = False) Then
    SQLText = "INSERT INTO charge_sheet VALUES(" _
    & Val(Me.comboBookingId.Text) & ", " _
    & Val(Me.txtMin.Text) & ", " _
    & Val(Me.txtDist.Text) & ", " _
    & Val(Me.txtAc.Text) & ", " _
    & Val(Me.txtNight.Text) & ", " _
    & Val(Me.txtWait.Text) & ")"
        
    'MsgBox SQLText
    mcon.Execute SQLText
    

Else
    SQLText = "update charge_sheet set " _
    & "TNIGHT = " & Val(Me.comboRate.Text) & ", " _
    & "TMIN = " & Val(Me.txtDistance.Text) & ", " _
    & "TDIST= " & Val(Me.txtWaitTime.Text) & ", " _
    & "TAC= " & Val(Me.txtTotalAmount.Text) & _
    " WHERE booking_id = " & Me.comboBookingId.Text
    
    'MsgBox SQLText
    mcon.Execute SQLText
    
End If

Unload Me
Call CenterMe(frmInvoiceList)
frmInvoiceList.Show

End Sub

Private Sub comboBookingId_Click()
Dim myName As String
Dim myBookDate, myBookTime As String
Dim myAddSource, myAddDest As String

myBookId = Me.comboBookingId.Text

SQLText = "SELECT * from booking where booking_id = " & myBookId
'MsgBox SQLText

rsTmp.Open SQLText, mcon, adOpenKeyset
myCustId = rsTmp("CUST_ID")
myBookDate = Format(rsTmp("BOOKING_DATE_TIME"), "dd-MMM-yyyy")
myBookTime = Format(rsTmp("BOOKING_DATE_TIME"), "hh:mm AM/PM")

myAddSource = rsTmp("PICKUP_AREA") & ", " & rsTmp("PICKUP_STREET") & ", " & rsTmp("PICKUP_BLOCK")
myAddDest = rsTmp("DROP_AREA") & ", " & rsTmp("DROP_STREET") & ", " & rsTmp("DROP_BLOCK")
rsTmp.Close

SQLText = "SELECT cust_first_name, cust_last_name from customer where cust_id = " & myCustId
rs.Open SQLText, mcon, adOpenKeyset
myName = rs("cust_first_name") & " " & rs("cust_last_name")
rs.Close

Me.txtBookDate = myBookDate
Me.txtBookTime = myBookTime
Me.txtCustName = myName
Me.txtSource = myAddSource
Me.txtDestination = myAddDest



End Sub



Private Sub comboRate_Click()

'It will clear the total amount if a different fare id is chosen than that from
'the original value

If (Me.comboRate.Text <> myFareId) Then
    Me.txtTotalAmount.Text = ""
End If

End Sub

Private Sub Form_Load()
'*******************
'**Form Load Event**
'*******************

edit = False

Set mcon = New ADODB.Connection
Set rs = New ADODB.Recordset
Set rsTmp = New ADODB.Recordset

total = 0
tmin = tdist = twait = tac = tnight = 0

mcon.Open "provider=MSDASQL.1;DRIVER=MICROSOFT ODBC FOR ORACLE;UID=SCOTT;PWD=TIGER"
InvoiceIdGenerate

Me.Visible = True
Me.comboBookingId.SetFocus

SQLTextA = "select * from TRIP where (STATUS = 'DONE' AND INVOICE = 'False') ORDER BY TRIP_ID"
'SQLTextA = "Select * from TRIP"
rsTmp.Open SQLTextA, mcon, adOpenKeyset
     
    Do While (rsTmp.EOF = False)
        'If (rsTmp("INVOICE") <> True) Then
        comboBookingId.AddItem rsTmp("TRIP_ID")
        'End If
        rsTmp.MoveNext
    Loop
rsTmp.Close

comboBookingId.listindex = -1

'***This filters out all the expired fare id******

SQLText = "SELECT fare_id FROM fare WHERE " _
& "(fare_validity > to_date('" & Format(Now, "dd-mm-yyyy") & "', 'dd-mm-yyyy')) "

rs.Open SQLText, mcon, adOpenKeyset

Do While (rs.EOF = False)
    Me.comboRate.AddItem rs("FARE_ID")
    rs.MoveNext
Loop
rs.Close

End Sub

Private Sub InvoiceIdGenerate()

Dim rsinvoiceID As ADODB.Recordset

Set rsinvoiceID = New ADODB.Recordset

rsinvoiceID.Open "select * from invoice order by invoice_id desc", mcon, adOpenKeyset
If rsinvoiceID.EOF = False Then
    Me.txtInvoiceNo.Text = rsinvoiceID("invoice_id") + 1
Else
    Me.txtInvoiceNo.Text = 70200
End If
End Sub

Public Sub FillInvoiceForm(flxref As MSFlexGrid, row As Integer)

'*****************************
'**Customer Form Fill Method**
'*****************************

edit = True

Me.cmdSave.Caption = "Update"
comboBookingId.Enabled = False


myBookId = flxref.TextMatrix(row, 1)
myInvoiceId = flxref.TextMatrix(row, 0)
myFareId = flxref.TextMatrix(row, 4)

SQLText = "SELECT *FROM charge_sheet WHERE booking_id = " & myBookId
'MsgBox SQLText

rs.Open SQLText, mcon, adOpenKeyset
Do While (rs.EOF = False)
    Me.txtAc = rs("TAC")
    Me.txtDist = rs("TDIST")
    Me.txtMin = rs("TMIN")
    Me.txtNight = rs("TNIGHT")
    Me.txtWait = rs("TWAIT")
    rs.MoveNext
Loop
rs.Close


SQLText = "SELECT * from booking where booking_id = " & myBookId

rsTmp.Open SQLText, mcon, adOpenKeyset
myCustId = rsTmp("CUST_ID")
myBookDate = Format(rsTmp("BOOKING_DATE_TIME"), "dd-MMM-yyyy")
myBookTime = Format(rsTmp("BOOKING_DATE_TIME"), "hh:mm AM/PM")

myAddSource = rsTmp("PICKUP_AREA") & ", " & rsTmp("PICKUP_STREET") & ", " & rsTmp("PICKUP_BLOCK")
myAddDest = rsTmp("DROP_AREA") & ", " & rsTmp("DROP_STREET") & ", " & rsTmp("DROP_BLOCK")
rsTmp.Close

With Me
    .txtInvoiceNo = flxref.TextMatrix(row, 0)
    .comboBookingId = flxref.TextMatrix(row, 1)
    .txtCustName = flxref.TextMatrix(row, 2)
    .txtBookDate = myBookDate
    .txtBookTime = myBookTime
    .comboRate = flxref.TextMatrix(row, 4)
    .txtWaitTime = flxref.TextMatrix(row, 6)
    .txtDistance = flxref.TextMatrix(row, 5)
    .txtTotalAmount = flxref.TextMatrix(row, 8)
    .txtSource = myAddSource
    .txtDestination = myAddDest
End With
End Sub

Public Sub Calculate_Fare()

If (Me.comboRate.Text = "") Then
    MsgBox "Choose a Rate from the drop-down"
    Me.comboRate.SetFocus
    Exit Sub
End If

If IsNumeric(Me.txtWaitTime) = False Then
    MsgBox ("Please enter a number for Wait Time")
    txtWaitTime.Text = ""
    txtWaitTime.SetFocus
    Exit Sub
End If

If IsNumeric(Me.txtDistance) = False Then
    MsgBox ("Please enter a number for Distance")
    txtDistance.Text = ""
    txtDistance.SetFocus
    Exit Sub
End If

total = 0

Dim myCabId As Integer

night = False

dist = Val(Me.txtDistance.Text)
waitTime = Val(Me.txtWaitTime.Text)
mytime = Format(Me.txtBookTime, "hhmm")

If (mytime > 2200) Then
    night = True
End If

If (mytime < 500) Then
    night = True
End If

SQLText = "SELECT cab_id from trip where trip_id = " & Me.comboBookingId.Text
rs.Open SQLText, mcon, adOpenKeyset
myCabId = rs("CAB_ID")
rs.Close

SQLText = "SELECT cab_ac from cab where cab_id = " & myCabId
rs.Open SQLText, mcon, adOpenKeyset
ac = rs("CAB_AC")
rs.Close

SQLText = "SELECT * FROM fare WHERE fare_id = " & Me.comboRate.Text
rs.Open SQLText, mcon, adOpenKeyset

Dim minCharge, perKmCharge, above100Charge, acCharge, nightCharge, upto15, above15 As Integer

minCharge = Val(rs("min_charge"))
perKmCharge = Val(rs("per_km_charge"))
above100Charge = Val(rs("above_hundred_charge"))
acCharge = Val(rs("ac_charge"))
nightCharge = Val(rs("night_charge"))
upto15 = Val(rs("upto_15"))
above15 = Val(rs("above_15"))

rs.Close
tmin = minCharge

If (dist <= 100) Then
    tdist = (dist * perKmCharge)
Else
    tdist = ((dist - 100) * above100Charge) + 100 * perKmCharge
End If

If (ac) Then
    tac = acCharge
Else
    tac = 0
End If

If (night) Then
    tnight = nightCharge
Else
    tnight = 0
End If

If (waitTime <= 15) Then
    twait = upto15
Else
    twait = upto15 + (above15 * (waitTime - 15))
End If

total = (tmin + tdist + tac + tnight + twait)

Me.txtAc.Text = tac
Me.txtDist = tdist
Me.txtMin = tmin
Me.txtNight = tnight
Me.txtWait = twait

Me.txtTotalAmount.Text = total

End Sub

