VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBookingList 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   FillColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmBookingList.frx":0000
   ScaleHeight     =   11190
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Generate &Report"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12000
      TabIndex        =   9
      Top             =   8760
      Width           =   1455
   End
   Begin VB.CommandButton ccmdSearch 
      Caption         =   "&Search"
      Height          =   735
      Left            =   13800
      Picture         =   "frmBookingList.frx":58FC
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "S&ort"
      Height          =   735
      Left            =   13800
      Picture         =   "frmBookingList.frx":5D3E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdBookingListDelete 
      Caption         =   "&Delete"
      Height          =   735
      Left            =   13800
      Picture         =   "frmBookingList.frx":6608
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdBookingListAdd 
      Caption         =   "&Add"
      Height          =   735
      Left            =   13800
      Picture         =   "frmBookingList.frx":6B92
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdBookingListClose 
      Caption         =   "&Exit"
      Height          =   735
      Left            =   13800
      Picture         =   "frmBookingList.frx":711C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton cmdBookingListRefresh 
      Caption         =   "&Refresh"
      Height          =   735
      Left            =   13800
      Picture         =   "frmBookingList.frx":76A6
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdBookingListEdit 
      Caption         =   "E&dit"
      Height          =   735
      Left            =   13800
      Picture         =   "frmBookingList.frx":7C30
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid flxBookingList 
      Height          =   6975
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   12303
      _Version        =   393216
      BackColorFixed  =   -2147483628
      GridColor       =   0
      HighLight       =   0
      ScrollBars      =   2
      SelectionMode   =   1
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
   Begin VB.Line Line1 
      BorderColor     =   &H00004040&
      BorderStyle     =   4  'Dash-Dot
      X1              =   13680
      X2              =   15120
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00004040&
      BorderStyle     =   4  'Dash-Dot
      X1              =   13680
      X2              =   13680
      Y1              =   2040
      Y2              =   8040
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00004040&
      BorderStyle     =   4  'Dash-Dot
      X1              =   13680
      X2              =   15120
      Y1              =   8040
      Y2              =   8040
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00004040&
      BorderStyle     =   4  'Dash-Dot
      X1              =   15120
      X2              =   15120
      Y1              =   8040
      Y2              =   2040
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Booking List"
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
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   13095
   End
End
Attribute VB_Name = "frmBookingList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim keyBook As String
Dim keyCust As String
Dim BookCount As Integer

Dim rsFill As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim rsName As ADODB.Recordset
Dim cn As ADODB.Connection

Private Sub ccmdSearch_Click()
Dim SQLText As String
frmBookingList.Enabled = False

Call Form_Click

calledFrom = 1
calledbySearch = True

SQLText = "SELECT tblcust.CUST_ID, tblcust.CUST_FIRST_NAME, " _
& "tblcust.CUST_LAST_NAME, tblbook.BOOKING_ID " _
& "from CUSTOMER tblcust, BOOKING tblbook " _
& "where tblcust.CUST_ID = tblbook.CUST_ID"

rs.Open SQLText, cn, adOpenKeyset

FillCombo frmSearch.Combo1, rs, False
rs.Close

Call CenterFrmSearch
frmSearch.Show
End Sub

Private Sub cmdBookingListClose_Click()

Unload Me
Call EnableCmd

Call CenterMe(frmHome)
frmHome.Show

End Sub

Private Sub cmdBookingListAdd_Click()
Unload Me
Call CenterMe(frmBooking)
frmBooking.Show
End Sub

Private Sub cmdBookingListDelete_Click()

'*****************************
'***Deleting Booking Reords***
'*****************************


Dim response As Integer
Dim message, SQLText As String

If (keyBook = "") Then
    MsgBox "You must select a row to edit"
    Exit Sub
End If

SQLText = "SELECT booking_status FROM booking where booking_id = " & keyBook
rs.Open SQLText, cn, adOpenKeyset

If (rs("BOOKING_STATUS") = "ONGO") Then
    MsgBox "Cab already dispatched!!! Abort via 'Command Center'"
    rs.Close
    Exit Sub
End If
rs.Close

message = "Are you sure you want to delete record with Booking Id : " & "" & Val(keyBook) & "" & " ?"
response = MsgBox(message, vbYesNo + vbQuestion, "Confirmation")

If response = 6 Then
    SQLText = "DELETE from BOOKING where BOOKING_ID = " & keyBook
    'MsgBox SQLText
    cn.Execute SQLText
Else
    Exit Sub
End If

Unload Me
Call CenterMe(frmBookingList)
frmBookingList.Show

End Sub

Private Sub cmdBookingListEdit_Click()
'Put all the data into the respective textboxes
'Calling the FillDriverForm method
'Calling the Driver Registration Form

If (keyCust = "") Or (keyBook = "") Then
    MsgBox "You must select a row to edit"
    Exit Sub
End If

Unload Me
Call CenterMe(frmBooking)
frmBooking.Show
'Passing the primary key
Call frmBooking.FillBookingForm(keyCust, keyBook)

End Sub

Private Sub cmdSort_Click()
Dim SQLText As String
frmBookingList.Enabled = False

calledFrom = 1
SQLText = "SELECT BOOKING_ID, BOOKING_DATE_TIME, CUST_ID, BOOKING_STATUS FROM BOOKING"
rs.Open SQLText, cn, adOpenKeyset
FillCombo frmSort.Combo1, rs, True
rs.Close

Call CenterFrmSort
frmSort.Show
End Sub

Private Sub Command1_Click()
BookingReport.Show
End Sub

Private Sub flxBookingList_Click()

If flxBookingList.row = 0 Then
    Exit Sub
End If
    
'flex grid click event to capture the primary key
Me.flxBookingList.Col = 0
keyBook = flxBookingList.Text 'KeyB
Me.flxBookingList.Col = 1
keyCust = flxBookingList.Text 'KeyA

xx = flxBookingList.row

'Code to highlight the entire row which ever is clicked by the user

With flxBookingList
    For j = 1 To (BookCount)
        .row = j
        If (j = xx) Then
            For i = 0 To 7
                .Col = i
                .CellBackColor = &H80FF80
            Next i
        Else
            For i = 0 To 7
            .Col = i
            .CellBackColor = vbWhite
            Next i
        End If
    Next j
End With
End Sub

Public Sub Form_Click()
keyCust = ""
keyBook = ""
With flxBookingList
    For j = 1 To (BookCount)
        .row = j
        For i = 0 To 7
            .Col = i
            .CellBackColor = vbWhite
        Next i
    Next j
End With

If (calledbySearch = False) Then
    Call TripToday
End If

End Sub

Private Sub Form_Load()

'*********************
'***FORM LOAD EVENT***
'*********************

keyCust = ""
keyBook = ""

Dim mySQLText As String

Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset
Set rsName = New ADODB.Recordset
Set rsFill = New ADODB.Recordset

cn.Open "Provider=msdasql.1; driver=Microsoft ODBC for Oracle; uid=scott; pwd=tiger;"

mySQLText = "select * from BOOKING order by BOOKING_ID"

Call FillList(mySQLText)
Call TripToday

End Sub


Public Sub FillList(SQLText As String)

'*******************************
'***Fill the List Accordingly***
'*******************************

rsFill.Open SQLText, cn, adOpenKeyset
BookCount = rsFill.RecordCount

With flxBookingList
    .Rows = BookCount + 1
    .Cols = 8
    j = 0
        
    .RowHeight(0) = 400
    For j = 0 To 7
        .row = 0
        .Col = j
        .CellFontBold = True
        .CellAlignment = 1
        .CellFontSize = 12
        .CellFontName = "Georgia"
        .ColWidth(j) = 1300
        'MsgBox rsFill.Fields.Item(j).Name
    Next j
    .ColWidth(0) = 1200
    .ColWidth(1) = 1200
    .ColWidth(2) = 2000
    .ColWidth(5) = 2500
    .ColWidth(6) = 2500
  
    
    .row = 0
    
    .Col = 0
    .Text = "Book ID"
    .Col = 1
    .Text = "Cust ID"
    .Col = 2
    .Text = "Name"
    .Col = 3
    .Text = "Date"
    .Col = 4
    .Text = "Time"
    .Col = 5
    .Text = "Source"
    .Col = 6
    .Text = "Destination"
    .Col = 7
    .Text = "Status"
            
    i = 1
    Do While rsFill.EOF = False
        
        .row = i
        
        .Col = 0
        .CellAlignment = 1
        .Text = rsFill("BOOKING_ID")
        .Col = 1
        .CellAlignment = 1
        .Text = rsFill("CUST_ID")
        .Col = 2
        .CellAlignment = 1
           
        custKey = rsFill("CUST_ID")
        SQLText = "SELECT * from CUSTOMER where CUST_ID = " & custKey & ""
        
        rsName.Open SQLText, cn, adOpenKeyset
        fullName = rsName("CUST_FIRST_NAME") & " " & rsName("CUST_LAST_NAME")
        .Text = fullName
        rsName.Close
        
        .Col = 3
        .CellAlignment = 1
        .Text = Format(rsFill("BOOKING_DATE_TIME"), "dd-mm-yyyy")
        .Col = 4
        .CellAlignment = 1
        .Text = Format(rsFill("BOOKING_DATE_TIME"), "hh : mm")
        .Col = 5
        .CellAlignment = 1
        .Text = rsFill("PICKUP_AREA")
        .Col = 6
        .CellAlignment = 1
        .Text = rsFill("DROP_AREA")
        .Col = 7
        .Text = rsFill("BOOKING_STATUS")
        
        i = i + 1
        rsFill.MoveNext
    Loop
    rsFill.Close
End With

End Sub

Public Sub TripToday()


Dim i, j As Integer
Dim chkDate As Date
Dim currDate As Date
Dim bookDate As Date

Dim mytime As String
Dim mydate As String
Dim diff As Long
Dim tt As String

currDate = Format(Now, "MM-dd-yyyy hh:mm")

With Me.flxBookingList
    For i = 1 To .Rows - 1
        .Col = 3
        .row = i
        mydate = .Text
        .Col = 4
        mytime = .Text
        tt = mydate & " " & mytime
      
        chkDate = Format(tt, "dd-MM-yyyy hh:mm")
        
        diff = DateDiff("h", currDate, chkDate)
        
        txtStatus = .TextMatrix(i, 7)
        
        'MsgBox diff & "  diff" & .TextMatrix(i, 7).Text & " STATUS"
        
        If (diff <= 24 And diff >= 1 And (txtStatus <> "DONE")) Then
            .Col = 0
            .row = i
            .CellBackColor = &HC0C0FF
        End If
       'red = &HFF&
       
        diff = DateDiff("n", currDate, chkDate)
        
        If (diff < 60 And diff > 0 And (txtStatus <> "DONE")) Then
            .Col = 0
            .row = i
            .CellBackColor = &HFF&
        End If
        
        
        If ((diff <= 0) And (txtStatus = "Pending") And (txtStatus <> "DONE")) Then
           .Col = 3
            .row = i
            .CellBackColor = &HFFFFC0
            .CellFontStrikeThrough = True
        End If
          
    Next i
       
    For i = 1 To .Rows - 1
        .Col = 7
        .row = i
       
        If (.Text = "ONGO") Then
                .Col = 0
                .row = i
                .CellBackColor = &HC0FFFF
        End If
    Next i
    
   
End With
End Sub
