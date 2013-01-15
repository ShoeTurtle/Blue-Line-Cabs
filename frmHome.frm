VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmHome 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15285
   FillColor       =   &H00C0FFFF&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmHome.frx":0000
   ScaleHeight     =   11190
   ScaleWidth      =   15285
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8400
      Width           =   1335
   End
   Begin VB.TextBox txtTripId 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      MaxLength       =   5
      TabIndex        =   17
      Top             =   7920
      Width           =   2655
   End
   Begin VB.CommandButton cmdDispatch 
      Caption         =   "&Dispatch"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7080
      Width           =   2415
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Trip &Complete"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7920
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   8520
      Top             =   3360
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13800
      TabIndex        =   5
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   10440
      TabIndex        =   4
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12120
      TabIndex        =   3
      Top             =   6840
      Width           =   1215
   End
   Begin VB.TextBox txtMsgBoard 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   10560
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   3600
      Width           =   4335
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2415
      Left            =   10200
      TabIndex        =   0
      Top             =   360
      Width           =   4935
      _Version        =   524288
      _ExtentX        =   8705
      _ExtentY        =   4260
      _StockProps     =   1
      BackColor       =   14737632
      Year            =   2011
      Month           =   9
      Day             =   23
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   0
      GridFontColor   =   4210816
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   16448
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid flxCabDriver 
      Height          =   2655
      Left            =   240
      TabIndex        =   10
      Top             =   1200
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4683
      _Version        =   393216
      Rows            =   4
      Cols            =   4
      BackColorFixed  =   -2147483628
      GridColor       =   64
      HighLight       =   0
      ScrollBars      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid flxTrip 
      Height          =   2655
      Left            =   240
      TabIndex        =   11
      Top             =   4320
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4683
      _Version        =   393216
      Rows            =   4
      Cols            =   4
      BackColorFixed  =   -2147483628
      GridColor       =   0
      ScrollBars      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdTripAbort 
      Caption         =   "&Trip Abort"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7920
      Width           =   1335
   End
   Begin VB.CommandButton cmdCommandCenter 
      Caption         =   "Co&mmand Center"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8400
      Width           =   2895
   End
   Begin VB.Label lblId 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   375
      Left            =   10440
      TabIndex        =   19
      Top             =   8160
      Width           =   4575
   End
   Begin VB.Image TheCab 
      Height          =   1095
      Left            =   7440
      Picture         =   "frmHome.frx":58FC
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   2655
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      X1              =   7320
      X2              =   10080
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      X1              =   7320
      X2              =   7320
      Y1              =   4200
      Y2              =   4920
   End
   Begin VB.Line Line11 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      X1              =   10080
      X2              =   10080
      Y1              =   4200
      Y2              =   4920
   End
   Begin VB.Line Line5 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      X1              =   7320
      X2              =   10080
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000C&
      X1              =   120
      X2              =   120
      Y1              =   720
      Y2              =   7560
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000C&
      X1              =   7080
      X2              =   7080
      Y1              =   720
      Y2              =   7560
   End
   Begin VB.Line Line10 
      BorderColor     =   &H8000000C&
      X1              =   120
      X2              =   7080
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000C&
      X1              =   120
      X2              =   7080
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Line Line9 
      BorderColor     =   &H8000000C&
      X1              =   7080
      X2              =   7080
      Y1              =   7680
      Y2              =   9000
   End
   Begin VB.Line Line8 
      BorderColor     =   &H8000000C&
      X1              =   120
      X2              =   7080
      Y1              =   9000
      Y2              =   9000
   End
   Begin VB.Line Line7 
      BorderColor     =   &H8000000C&
      X1              =   120
      X2              =   120
      Y1              =   7680
      Y2              =   9000
   End
   Begin VB.Line Line6 
      BorderColor     =   &H8000000C&
      X1              =   120
      X2              =   7080
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Driver - Cab Details"
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
      Left            =   240
      TabIndex        =   14
      Top             =   840
      Width           =   6255
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Trip Details"
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
      Left            =   240
      TabIndex        =   13
      Top             =   3960
      Width           =   6375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Trip ID :"
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
      Left            =   360
      TabIndex        =   12
      Top             =   7920
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   10440
      X2              =   15000
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Dispatch Control"
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
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   6975
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   6
      Top             =   4320
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Message Board"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   375
      Left            =   10560
      TabIndex        =   2
      Top             =   3120
      Width           =   4335
   End
End
Attribute VB_Name = "frmHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim keyDriver As String
Dim keyCab As String
Dim keyCust As String
Dim keyBook As String

Dim rsName As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim rsTrip As ADODB.Recordset

Dim BookCount As Integer
Dim AssignCount As Integer

Dim cn As ADODB.Connection

Private Sub cmdCommandCenter_Click(Index As Integer)

Call DisableCmd(10)
frmHome.Enabled = False
    
    With frmCommandCenter
        .Width = 6075
        .Height = 6060
        .Left = (MDIMain.ScaleWidth - .Width) / 2
        .Top = (MDIMain.ScaleHeight - .Height) / 2
        .Show
    End With

End Sub

Private Sub cmdDispatch_Click(Index As Integer)

If (keyCab = "") Or (keyCust = "") Or (keyBook = "") Then
    MsgBox "Select Trip and Assign Cab then Dispatch"
    Exit Sub
End If

SQLText = "SELECT * FROM cab_driver_assign where cab_id = " & "" & Val(keyCab) & ""
rs.Open SQLText, cn, adOpenKeyset

If (rs("ONHIRE")) Then
    MsgBox "Cab is On-Hire Select a different Cab"
    rs.Close
    Exit Sub
End If
rs.Close

SQLText = "SELECT * from booking where booking_id = " & "" & Val(keyBook) & ""
rs.Open SQLText, cn, adOpenKeyset

If (rs("BOOKING_STATUS") = "ONGO") Then
    MsgBox "Cab has already been dispatched for the trip"
    rs.Close
    Exit Sub
End If

Dim currDate As Date
Dim bookDate As Date

currDate = Format(Now, "mm-dd-yyyy hh:mm")
bookDate = Format(rs("booking_date_time"), "mm-dd-yyyy hh:mm")
rs.Close

'MsgBox "Today's : " & currDate & "  Booking : " & bookDate

diff = DateDiff("h", currDate, bookDate)

If (diff >= 24) Then
    MsgBox "Dispatch time is 24hrs Ahead"
    Exit Sub
End If


SQLText = "INSERT into TRIP values(" _
& "" & Val(keyBook) & ", " _
& "" & Val(keyCab) & ", " _
& "" & Val(keyDriver) & ", " _
& "'ONGO', " _
& "'False')"

'MsgBox SQLText

cn.Execute SQLText

SQLText = "UPDATE CAB_DRIVER_ASSIGN SET " _
& "ONHIRE = 'True' " _
& "WHERE CAB_ID = " & Val(keyCab) & " "

'MsgBox SQLText

cn.Execute SQLText

SQLText = "UPDATE BOOKING SET " _
& "BOOKING_STATUS = 'ONGO' " _
& "WHERE BOOKING_ID = " & "" & keyBook & ""

cn.Execute SQLText

'MsgBox SQLText


Unload frmHome
Call CenterMe(frmHome)
frmHome.Show

End Sub

Private Sub cmdDone_Click(Index As Integer)

SQLText = "SELECT * from trip WHERE trip_id = " & "" & Val(txtTripId.Text) & ""
rs.Open SQLText, cn, adOpenKeyset

If (rs.RecordCount = 0) Then
    MsgBox "Invalid Trip Id"
    rs.Close
    Exit Sub
ElseIf (rs("STATUS") <> "ONGO") Then
    MsgBox "Invalid Trip Id"
    rs.Close
    Exit Sub
End If
rs.Close


SQLText = "UPDATE trip SET status = 'DONE' WHERE trip_id = " & "" & Val(txtTripId.Text) & ""
'MsgBox SQLText
cn.Execute SQLText


SQLText = "SELECT cab_id from trip where trip_id = " & "" & Val(txtTripId.Text) & ""
'MsgBox SQLText
rs.Open SQLText, cn, adOpenKeyset


SQLText = "UPDATE cab_driver_assign SET onhire = 'FALSE' " _
& "WHERE cab_id = " & "" & Val(rs("CAB_ID")) & ""
'MsgBox SQLText
cn.Execute SQLText
rs.Close


SQLText = "UPDATE booking SET booking_status = 'DONE' " _
& "WHERE booking_id = " & "" & Val(txtTripId.Text) & ""
'MsgBox SQLText
cn.Execute SQLText

Unload frmHome
Call CenterMe(frmHome)
frmHome.Show

End Sub

Private Sub cmdRefresh_Click(Index As Integer)
keyCust = ""
keyBook = ""
keyCab = ""
keyDriver = ""
Me.txtTripId = ""

With flxTrip
    For j = 1 To (BookCount)
        .row = j
        For i = 0 To 5
            .Col = i
            .CellBackColor = vbWhite
        Next i
    Next j
End With

With flxCabDriver
    For j = 1 To (AssignCount)
        .row = j
        For i = 0 To 4
            .Col = i
            .CellBackColor = vbWhite
        Next i
    Next j
End With

Call Form_Load

End Sub

Private Sub cmdSave_Click()

If (MDIMain.cmdMain(0).Enabled = False) Then
    SQLText = "UPDATE user_table SET user_note = '" & Me.txtMsgBoard.Text & "'" _
    & " WHERE user_name = '" & Module1.currUserName & "'"
    
    MsgBox SQLText
    cn.Execute SQLText
    MsgBox "Notes Saved!!"
End If

End Sub

Private Sub cmdTripAbort_Click(Index As Integer)

Dim myTripId As Long

myTripId = Val(Me.txtTripId.Text)

If (myTripId = 0) Then
    MsgBox "Enter a Valid Trip Id"
    Exit Sub
End If

SQLText = "SELECT * FROM booking WHERE booking_id = " & "" & myTripId & ""
'MsgBox SQLText
rs.Open SQLText, cn, adOpenKeyset

If (rs.RecordCount = 0) Then
    MsgBox "Invalid trip Id"
    rs.Close
    Exit Sub
End If

If (rs("BOOKING_STATUS") <> "ONGO") Then
    MsgBox "Cab has not been dispatched delete from Booking Vew"
    rs.Close
    Exit Sub
End If
rs.Close


SQLText = "SELECT * from TRIP where TRIP_ID = " & "" & myTripId & ""
'MsgBox SQLText

rs.Open SQLText, cn, adOpenKeyset

myCabId = rs("CAB_ID")
rs.Close

Dim response As Integer
Dim message, mySQLText As String

message = "Are you sure you want to delete record with Booking Id : " & "" & myTripId & "" & " ?"
response = MsgBox(message, vbYesNo + vbQuestion, "Confirmation")

If response = 6 Then
    
    SQLText = "UPDATE cab_driver_assign SET onhire = 'FALSE' " _
    & "WHERE cab_id = " & "" & Val(myCabId) & ""
    'MsgBox SQLText
    cn.Execute SQLText
    
    
    SQLText = "DELETE from BOOKING where BOOKING_ID = " & myTripId
    'MsgBox SQLText
    cn.Execute SQLText
    
    'SQLText = "DELETE from TRIP where TRIP_ID = " & myTripId
    'MsgBox SQLText
    
    'SQLText = "UPDATE trip SET status = 'DONE' WHERE trip_id = " & "" & Val(myTripId) & ""
    'MsgBox SQLText
    'cn.Execute SQLText


    'SQLText = "SELECT cab_id from trip where trip_id = " & "" & Val(myTripId) & ""
    'MsgBox SQLText
    'rs.Open SQLText, cn, adOpenKeyset

    'SQLText = "UPDATE booking SET booking_status = 'DONE' " _
    '& "WHERE booking_id = " & "" & Val(txtTripId.Text) & ""
    'MsgBox SQLText
    'cn.Execute SQLText
    
Else
    Exit Sub
End If

Unload Me
Call CenterMe(frmHome)
frmHome.Show
End Sub

Private Sub flxCabDriver_Click()

If flxCabDriver.row = 0 Then
    Exit Sub
End If
    
'flex grid click event to capture the primary key
Me.flxCabDriver.Col = 0  'setting column to 0
keyDriver = flxCabDriver.Text  'row is automatically set
Me.flxCabDriver.Col = 2
keyCab = flxCabDriver.Text

'MsgBox "DRIVER ID: " & keyDriver
'MsgBox "CAB ID: " & keyCab
selectedRow = flxCabDriver.row

'Code to highlight the entire row which ever is clicked by the user

With flxCabDriver
    For j = 1 To (AssignCount)
        .row = j
        If (j = selectedRow) Then
            For i = 0 To 4
                .Col = i
                .CellBackColor = &H80FF80
            Next i
        Else
            For i = 0 To 4
            .Col = i
            .CellBackColor = vbWhite
            Next i
        End If
    Next j
End With
End Sub

Private Sub flxTrip_Click()

If flxTrip.row = 0 Then
    Exit Sub
End If
    
'flex grid click event to capture the primary key
Me.flxTrip.Col = 0  'setting column to 0
keyBook = flxTrip.Text  'row is automatically set
Me.flxTrip.Col = 3
keyCust = flxTrip.Text

'MsgBox "BOOK ID: " & keyBook
'MsgBox "CUST ID: " & keyCust
selectedRow = flxTrip.row

'Code to highlight the entire row which ever is clicked by the user

With flxTrip
    For j = 1 To (BookCount)
        .row = j
        If (j = selectedRow) Then
            For i = 0 To 5
                .Col = i
                .CellBackColor = &H80FF80
            Next i
        Else
            For i = 0 To 5
            .Col = i
            .CellBackColor = vbWhite
            Next i
        End If
    Next j
End With

End Sub

Private Sub Form_Load()

'*********************
'***FORM LOAD EVENT***
'*********************



key = ""
Me.Calendar1.Value = Now

Set cn = New ADODB.Connection
Set rsAssign = New ADODB.Recordset
Set rsName = New ADODB.Recordset
Set rsTrip = New ADODB.Recordset
Set rs = New ADODB.Recordset


cn.Open "Provider=msdasql.1; driver=Microsoft ODBC for Oracle; uid=scott; pwd=tiger;"


SQLText = "SELECT * from CAB_DRIVER_ASSIGN order by CAB_ID"
rsAssign.Open SQLText, cn, adOpenKeyset
AssignCount = rsAssign.RecordCount


With flxCabDriver
    .Rows = AssignCount + 1
    .RowHeight(0) = 300
    .Cols = 5
    j = 0
    For j = 0 To 4
        .row = 0
        .Col = j
        .CellFontBold = True
        .CellAlignment = 1
        .CellFontSize = 11
        .CellFontName = "Georgia"
        .ColWidth(.Col) = 1600
    Next j

    .ColWidth(0) = 1200
    .ColWidth(2) = 1200
    .ColWidth(1) = 2000
    .ColWidth(3) = 1200
    
    
    .row = 0
    
    .Col = 0
    .Text = "Driver Id"
    .Col = 1
    .Text = "Name"
    .Col = 2
    .Text = "Cab Id"
    .Col = 3
    .Text = "Status"
    .Col = 4
    .Text = "Trip"
    
    
    i = 1
    Do While rsAssign.EOF = False
        .row = i
    
        .Col = 0
        .CellAlignment = 1
        .Text = rsAssign("DRIVER_ID")
        .Col = 1
        
        SQLText = "SELECT DRIVER_FIRST_NAME, DRIVER_LAST_NAME from DRIVER where " _
        & "DRIVER_ID = " & rsAssign("DRIVER_ID") & ""
        rsName.Open SQLText, cn, adOpenKeyset
        .Text = rsName("DRIVER_FIRST_NAME") & "  " & rsName("DRIVER_LAST_NAME")
        rsName.Close
                
        .Col = 2
        .CellAlignment = 1
        .Text = rsAssign("CAB_ID")
        .Col = 3
        .CellAlignment = 1
        
        If (rsAssign("ONHIRE")) Then
            .Text = "On-Hire"
        Else
            .Text = "Stand-By"
        End If
               
        .Col = 4
        .CellAlignment = 1
        SQLText = "SELECT tt.trip_id, cc.driver_id FROM " _
        & "trip tt, cab_driver_assign cc " _
        & "WHERE tt.cab_id = cc.cab_id AND tt.status = 'ONGO' " _
        & "AND tt.driver_id = " & "" & rsAssign("DRIVER_ID") & ""
        
        rsName.Open SQLText, cn, adOpenKeyset
               
        'MsgBox SQLText
        
        If (rsName.EOF = False) Then
            .Text = rsName("TRIP_ID")
        Else
            .Text = "     -  "
        End If
        rsName.Close
        
        i = i + 1
        rsAssign.MoveNext
    Loop
    rsAssign.Close
End With
    
Dim currDate As Date
currDate = Format(Now, "MM-dd-yyyy")

'***Query to filter out missed booking****
'***i.e the date has already passed*******

SQLText = "SELECT booking_id, booking_date_time, booking_status, cust_id " _
& "FROM booking  WHERE " _
& "(booking_status = 'Pending' AND " _
& "booking_date_time > to_date('" & Format(Now, "dd-mm-yyyy hh-mm") & "', 'dd-mm-yyyy hh24-mi')) " _
& "OR (booking_status = 'ONGO') "


rsTrip.Open SQLText, cn, adOpenKeyset
   
BookCount = rsTrip.RecordCount
    
With flxTrip
    .Rows = BookCount + 1
    .RowHeight(0) = 300
    .Cols = 6
    j = 0
    For j = 0 To 5
        .row = 0
        .Col = j
        .CellFontBold = True
        .CellAlignment = 1
        .CellFontSize = 11
        .CellFontName = "Georgia"
        .ColWidth(.Col) = 800
    Next j

    .ColWidth(0) = 850
    .ColWidth(1) = 1080
    .ColWidth(3) = 990
    .ColWidth(4) = 2000
    .ColWidth(5) = 1000

    .row = 0
    
    .Col = 0
    .Text = "Trip"
    .Col = 1
    .Text = "Date"
    .Col = 2
    .Text = "Time"
    .Col = 3
    .Text = "Cust Id"
    .Col = 4
    .Text = "Name"
    .Col = 5
    .Text = "Status"
    
    
    i = 1
    Do While rsTrip.EOF = False
        .row = i
    
        .Col = 0
        .CellAlignment = 1
        .Text = rsTrip("BOOKING_ID")
        .Col = 1
        .CellAlignment = 1
        .Text = Format(rsTrip("BOOKING_DATE_TIME"), "dd-mm-yyyy")
        .Col = 2
        .CellAlignment = 1
        .Text = Format(rsTrip("BOOKING_DATE_TIME"), "hh : mm")
        .Col = 3
        .CellAlignment = 1
        .Text = rsTrip("CUST_ID")
        
        .Col = 4
        .CellAlignment = 1
        
        SQLText = "SELECT CUST_FIRST_NAME, CUST_LAST_NAME from CUSTOMER where " _
        & "CUST_ID = " & rsTrip("CUST_ID") & ""
        
        rsName.Open SQLText, cn, adOpenKeyset
        .Text = rsName("CUST_FIRST_NAME") & "  " & rsName("CUST_LAST_NAME")
        rsName.Close
        
        .Col = 5
        .CellAlignment = 1
        .Text = rsTrip("BOOKING_STATUS")
               
        i = i + 1
        rsTrip.MoveNext
    Loop
    rsTrip.Close
End With


If (Module1.AsAdmin) Then
    Me.lblId.Caption = "ADMINISTRATIVE-LOGIN"
Else
    Me.lblId.Caption = UCase(Module1.currUserName)
End If

Call TripToday

End Sub
 

Private Sub Timer1_Timer()
'lblTime.Caption = Format(Time, "hh:mm:ss")
lblTime.Caption = Time
'lblDate.Caption = Date
End Sub

Private Sub TripToday()

Dim i, j As Integer
Dim chkDate As Date
Dim currDate As Date
Dim bookDate As Date

Dim mytime As String
Dim mydate As String
Dim diff As Long
Dim tt As String

currDate = Format(Now, "MM-dd-yyyy hh:mm")

With flxTrip
    For i = 1 To .Rows - 1
        .Col = 1
        .row = i
        mydate = .Text
        .Col = 2
        mytime = .Text
        tt = mydate & " " & mytime
      
        chkDate = Format(tt, "dd-MM-yyyy hh:mm")
        
        diff = DateDiff("h", currDate, chkDate)
              
        If (diff < 24 And diff >= 1) Then
            .Col = 0
            .row = i
            .CellBackColor = &HC0C0FF
        End If
        
        diff = DateDiff("n", currDate, chkDate)
        
        If (diff < 60 And diff > 0) Then
            .Col = 0
            .row = i
            .CellBackColor = &HFF&
        End If
        
        
        'If (diff = 1) Then
        '    .Col = 0
        '    .Row = i
        '    .CellBackColor = &HFF&
        'End If
        
    Next i
       
    For i = 1 To .Rows - 1
        .Col = 5
        .row = i
       
        If (.Text = "ONGO") Then
                .Col = 0
                .row = i
                .CellBackColor = &HC0FFFF
        End If
    Next i
    
End With
End Sub



