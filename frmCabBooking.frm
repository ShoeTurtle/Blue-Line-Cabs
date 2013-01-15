VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBooking 
   BorderStyle     =   0  'None
   Caption         =   "Cab Registration - Blue Line "
   ClientHeight    =   9825
   ClientLeft      =   1965
   ClientTop       =   2370
   ClientWidth     =   15285
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   Picture         =   "frmCabBooking.frx":0000
   ScaleHeight     =   9825
   ScaleWidth      =   15285
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox comboNoPassenger 
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
      ItemData        =   "frmCabBooking.frx":58FC
      Left            =   4560
      List            =   "frmCabBooking.frx":590F
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Add"
      Height          =   735
      Left            =   6960
      Picture         =   "frmCabBooking.frx":5930
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7920
      Width           =   1215
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   735
      Left            =   4920
      Picture         =   "frmCabBooking.frx":65FA
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7920
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Go &Back"
      Height          =   735
      Left            =   8880
      Picture         =   "frmCabBooking.frx":6B84
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7920
      Width           =   1215
   End
   Begin VB.TextBox txtDestArea 
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
      Left            =   9600
      TabIndex        =   10
      Top             =   3600
      Width           =   2895
   End
   Begin VB.TextBox txtDestStreet 
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
      Left            =   9600
      TabIndex        =   11
      Top             =   4080
      Width           =   2895
   End
   Begin VB.TextBox txtDestBlock 
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
      Left            =   9600
      TabIndex        =   12
      Top             =   4560
      Width           =   2895
   End
   Begin VB.TextBox txtPickupArea 
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
      Left            =   9600
      TabIndex        =   7
      Top             =   1800
      Width           =   2895
   End
   Begin VB.TextBox txtPickupStreet 
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
      Left            =   9600
      TabIndex        =   8
      Top             =   2280
      Width           =   2895
   End
   Begin VB.TextBox txtPickupBlock 
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
      Left            =   9600
      TabIndex        =   9
      Top             =   2760
      Width           =   2895
   End
   Begin VB.TextBox txtBookingId 
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
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CheckBox checkReturn 
      BackColor       =   &H80000014&
      Caption         =   "Return Trip"
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
      TabIndex        =   5
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox txtCustName 
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
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   2640
      Width           =   2175
   End
   Begin VB.ComboBox comboCustID 
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
      Left            =   3720
      TabIndex        =   0
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox txtPickupDirection 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7920
      MaxLength       =   400
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   5640
      Width           =   4575
   End
   Begin MSComCtl2.DTPicker dtPickBookDate 
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   4920
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   66387971
      CurrentDate     =   40838
   End
   Begin MSComCtl2.DTPicker dtPickTime 
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   5640
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "hh:mm "
      Format          =   66387970
      CurrentDate     =   40838
   End
   Begin VB.Line Line25 
      BorderColor     =   &H8000000C&
      X1              =   12720
      X2              =   12720
      Y1              =   5040
      Y2              =   7080
   End
   Begin VB.Line Line24 
      BorderColor     =   &H8000000C&
      X1              =   16560
      X2              =   16560
      Y1              =   1320
      Y2              =   2280
   End
   Begin VB.Line Line23 
      BorderColor     =   &H8000000C&
      X1              =   7680
      X2              =   12720
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Line Line22 
      BorderColor     =   &H8000000C&
      X1              =   7680
      X2              =   7680
      Y1              =   5040
      Y2              =   7080
   End
   Begin VB.Line Line20 
      BorderColor     =   &H00004000&
      BorderStyle     =   4  'Dash-Dot
      X1              =   10200
      X2              =   10200
      Y1              =   7800
      Y2              =   8760
   End
   Begin VB.Line Line19 
      BorderColor     =   &H00004000&
      BorderStyle     =   4  'Dash-Dot
      X1              =   4800
      X2              =   4800
      Y1              =   7800
      Y2              =   8760
   End
   Begin VB.Line Line18 
      BorderColor     =   &H00004000&
      BorderStyle     =   4  'Dash-Dot
      X1              =   4800
      X2              =   10200
      Y1              =   8760
      Y2              =   8760
   End
   Begin VB.Line Line17 
      BorderColor     =   &H00004000&
      BorderStyle     =   4  'Dash-Dot
      X1              =   4800
      X2              =   10200
      Y1              =   7800
      Y2              =   7800
   End
   Begin VB.Line Line16 
      BorderColor     =   &H8000000C&
      X1              =   12720
      X2              =   12720
      Y1              =   3480
      Y2              =   5040
   End
   Begin VB.Line Line15 
      BorderColor     =   &H8000000C&
      X1              =   7680
      X2              =   7680
      Y1              =   3480
      Y2              =   5040
   End
   Begin VB.Line Line14 
      BorderColor     =   &H8000000C&
      X1              =   12720
      X2              =   12720
      Y1              =   1560
      Y2              =   3480
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000C&
      X1              =   7680
      X2              =   7680
      Y1              =   1560
      Y2              =   3480
   End
   Begin VB.Line Line9 
      BorderColor     =   &H8000000C&
      X1              =   7680
      X2              =   12720
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line8 
      BorderColor     =   &H8000000C&
      X1              =   6960
      X2              =   6960
      Y1              =   4560
      Y2              =   7080
   End
   Begin VB.Line Line7 
      BorderColor     =   &H8000000C&
      X1              =   2040
      X2              =   6960
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Line Line6 
      BorderColor     =   &H8000000C&
      X1              =   2040
      X2              =   2040
      Y1              =   4560
      Y2              =   7080
   End
   Begin VB.Line Line5 
      BorderColor     =   &H8000000C&
      X1              =   2040
      X2              =   6960
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000C&
      X1              =   2040
      X2              =   6960
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000C&
      X1              =   2040
      X2              =   2040
      Y1              =   1560
      Y2              =   3960
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000C&
      X1              =   6960
      X2              =   6960
      Y1              =   3960
      Y2              =   1560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      X1              =   2040
      X2              =   6960
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label lblDestArea 
      BackStyle       =   0  'Transparent
      Caption         =   "Dest' Area :"
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
      Left            =   7920
      TabIndex        =   30
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label lblDestStreet 
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
      Left            =   7920
      TabIndex        =   29
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label lblDestBlock 
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
      Left            =   7920
      TabIndex        =   28
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Pickup Area : "
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
      Left            =   7920
      TabIndex        =   27
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Street  :"
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
      Left            =   7920
      TabIndex        =   26
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblPickupBlock 
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
      Left            =   7920
      TabIndex        =   25
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label2 
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
      Left            =   2280
      TabIndex        =   24
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label3 
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
      Height          =   255
      Left            =   2280
      TabIndex        =   23
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label Label4 
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
      Left            =   2280
      TabIndex        =   22
      Top             =   5640
      Width           =   855
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "No Of Passangers :"
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
      Left            =   2280
      TabIndex        =   21
      Top             =   6360
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Cust Name :"
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
      Left            =   2280
      TabIndex        =   20
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Cust ID:"
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
      Left            =   2280
      TabIndex        =   19
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Pickup Direction :"
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
      Left            =   9360
      TabIndex        =   18
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Booking Information"
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
      Left            =   2040
      TabIndex        =   17
      Top             =   600
      Width           =   10695
   End
End
Attribute VB_Name = "frmBooking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs As ADODB.Recordset
Dim rsTmp As ADODB.Recordset
Dim mcon As ADODB.Connection

Dim myKey As String
Dim edit As Boolean

Private Sub cmdClose_Click()
Unload Me
Call CenterMe(frmBookingList)
frmBookingList.Show
End Sub



Private Sub cmdSave_Click()

'*****************************
'***SAVE / EDIT Fare Record***
'*****************************


Dim mydate, mytime, myDateTime As String
Dim flg As Integer
Dim bookDateTime As Date
Dim message As String
Dim SQLText As String

mydate = Format(Me.dtPickBookDate.Value, "dd-mm-yyyy")
mytime = Format(Me.dtPickTime.Value, "hh-mm")
myDateTime = mydate & " " & mytime


'********Checking status before UPDATING RECORDS***********
If (edit) Then

    SQLText = "SELECT * FROM booking where booking_id = " & myKey
    'MsgBox SQLText

    rs.Open SQLText, mcon, adOpenKeyset


    If (rs("BOOKING_STATUS") = "DONE") Then
        MsgBox "Trip is complete unable to edit"
        Exit Sub
        rs.Close
    End If

    If (rs("BOOKING_STATUS") = "ONGO") Then
        MsgBox "Cab has been dispatched, to abort goto Command Center"
        Exit Sub
        rs.Close
    End If
    rs.Close

    SQLText = "SELECT to_char(booking_date_time, 'MM-dd-yyyy HH24:MI') FROM booking WHERE " _
    & "booking_id = " & myKey
    rs.Open SQLText, mcon, adOpenKeyset
    'MsgBox SQLText

    bookDateTime = rs(0)
    rs.Close
    
    'MsgBox bookDateTime

    
    If (DateDiff("n", Now, bookDateTime) < 0) Then
        message = "Trip has Expired, Do you want to make changes ?"
        response = MsgBox(message, vbYesNo + vbQuestion, "Confirmation")

                If response = 6 Then
                    GoTo ProcedeToEdit
                Else
                    Exit Sub
                End If
    End If
End If

ProcedeToEdit:

'*************END OF STATUS CHECK FOR UPDATION************


xx = Format(Me.dtPickBookDate.Value, "MM-dd-yyyy")
yy = Format(Me.dtPickTime.Value, "hh:mm")
bookDateTime = Format(xx & " " & yy, "MM-dd-yyyy hh:mm")

If (DateDiff("n", Now, bookDateTime) <= 0) Then
    MsgBox "Invalid Booking Date Time"
    Me.dtPickBookDate.SetFocus
    Exit Sub
End If

If (Me.comboNoPassenger.Text = "") Then
    MsgBox "Select the no of Passengers"
    Exit Sub
End If

 
If (Me.txtPickupArea.Text = "" Or Me.txtPickupBlock = "" Or Me.txtPickupStreet.Text = "" _
Or Me.txtDestArea.Text = "" Or Me.txtDestBlock.Text = "" Or Me.txtDestStreet = "" Or _
Me.txtPickupDirection.Text = "") Then
    Status = MsgBox("Please fill out the address details.", vbOKOnly, "Incomplete Address Info")
    Me.txtPickupArea.SetFocus
    Exit Sub
End If


If (edit = False) Then
    SQLTextA = "INSERT into BOOKING values(" _
    & "" & Val(Me.txtBookingId) & ", " _
    & "to_date('" & myDateTime & "', 'dd-mm-yyyy hh24-mi')," _
    & "" & Val(Me.checkReturn.Value) & ", " _
    & "" & Val(Me.comboCustID.Text) & ", " _
    & "'" & Trim(Me.comboNoPassenger.Text) & "', " _
    & "'" & Trim(Me.txtPickupArea.Text) & "', " _
    & "'" & Trim(Me.txtPickupStreet.Text) & "', " _
    & "'" & Trim(Me.txtPickupBlock.Text) & "', " _
    & "'" & Trim(Me.txtDestArea.Text) & "', " _
    & "'" & Trim(Me.txtDestStreet.Text) & "', " _
    & "'" & Trim(Me.txtDestBlock.Text) & "', " _
    & "'" & Trim(Me.txtPickupDirection.Text) & "', " _
    & "'Pending' )"

    'MsgBox SQLTextA
    mcon.Execute SQLTextA

Else
    
    SQLTextB = "UPDATE BOOKING set " _
    & "BOOKING_DATE_TIME = to_date('" & myDateTime & "', 'dd-mm-yyyy hh24-mi'), " _
    & "RETURN_TRIP = " & Val(Me.checkReturn.Value) & ", " _
    & "PASSENGER_NO = '" & Trim(Me.comboNoPassenger.Text) & "' , " _
    & "PICKUP_AREA = '" & Trim(Me.txtPickupArea.Text) & "' , " _
    & "PICKUP_STREET = '" & Trim(Me.txtPickupStreet.Text) & "' , " _
    & "PICKUP_BLOCK = '" & Trim(Me.txtPickupBlock.Text) & "' , " _
    & "DROP_AREA = '" & Trim(Me.txtDestArea.Text) & "' , " _
    & "DROP_STREET = '" & Trim(Me.txtDestStreet.Text) & "' , " _
    & "DROP_BLOCK = '" & Trim(Me.txtDestBlock.Text) & "',  " _
    & "DIRECTION = '" & Trim(Me.txtPickupDirection.Text) & "' " _
    & "WHERE BOOKING_ID = " & Val(myKey) & ""

    'MsgBox SQLTextB
    mcon.Execute SQLTextB

End If
 
Unload Me
Call CenterMe(frmBookingList)
frmBookingList.Show

End Sub

Private Sub comboCustID_Click()

rsTmp.MoveFirst
rsTmp.Move (Me.comboCustID.listindex)
CustFullName = rsTmp("CUST_FIRST_NAME") & " " & rsTmp("CUST_LAST_NAME")
txtCustName.Text = CustFullName

End Sub

Private Sub Form_Load()

'*********************
'***From Load Event***
'*********************

flgcmdGoCust = False 'some shitty flag
edit = False
Me.dtPickBookDate.Value = Now
Me.dtPickTime.Value = Now
Me.comboNoPassenger.listindex = -1

Set rsTmp = New ADODB.Recordset
Set mcon = New ADODB.Connection
Set rs = New ADODB.Recordset

mcon.Open "provider=MSDASQL.1;DRIVER=MICROSOFT ODBC FOR ORACLE;UID=SCOTT;PWD=TIGER"
Call BookingIdGenerate

Me.Visible = True
Me.comboCustID.SetFocus

'Code to populate the combo drop down
 
SQLTextA = "select * from CUSTOMER Order by CUST_ID"
rsTmp.Open SQLTextA, mcon, adOpenKeyset
    rsTmp.MoveFirst
    Do Until rsTmp.EOF
        comboCustID.AddItem rsTmp("CUST_ID")
        rsTmp.MoveNext
    Loop
            
comboCustID.listindex = 0

End Sub

Public Sub FillBookingForm(keyCust As String, keyBook As String)

'******************************
'***Booking Form Fill Method***
'******************************
edit = True
myKey = keyBook

Me.cmdSave.Caption = "Update"

Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset


Me.txtBookingId.Locked = True
Me.txtCustName.Locked = True
Me.comboCustID.Locked = True


SQLTextA = "Select * from CUSTOMER where(CUST_ID = " & keyCust & ")"
rs.Open SQLTextA, mcon, adOpenKeyset
fullName = rs("CUST_FIRST_NAME") & " " & rs("CUST_LAST_NAME")
rs.Close

SQLTextB = "Select * from BOOKING where (BOOKING_ID = " & keyBook & ")"
rs.Open SQLTextB, mcon, adOpenKeyset


With frmBooking
    .comboCustID.Text = rs("CUST_ID")
    .txtBookingId.Text = rs("BOOKING_ID")
    .txtCustName.Text = fullName
    .comboNoPassenger.Text = rs("PASSENGER_NO")
    
    .dtPickBookDate.Value = Format(rs("BOOKING_DATE_TIME"), "mm-dd-yyyy")
    .dtPickTime.Value = Format(rs("BOOKING_DATE_TIME"), "hh:mm:ss")
    
    .checkReturn.Value = rs("RETURN_TRIP")
              
    .txtPickupArea = rs("PICKUP_AREA")
    .txtPickupStreet = rs("PICKUP_STREET")
    .txtPickupBlock = rs("PICKUP_BLOCK")
    
    .txtDestArea = rs("DROP_AREA")
    .txtDestStreet = rs("DROP_STREET")
    .txtDestBlock = rs("DROP_BLOCK")
    .txtPickupDirection = rs("DIRECTION")
    
End With
rs.Close

End Sub

Private Sub BookingIdGenerate()

'*************************
'***Generate Booking Id***
'*************************

Dim rsbookID As ADODB.Recordset
Set rsbookID = New ADODB.Recordset
rsbookID.Open "select * from BOOKING order by BOOKING_ID desc", mcon, adOpenKeyset
If rsbookID.EOF = False Then
    Me.txtBookingId.Text = rsbookID("BOOKING_ID") + 1
Else
    Me.txtBookingId.Text = 10200
End If
End Sub

