VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmInvoiceList 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15285
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmInvoiceList.frx":0000
   ScaleHeight     =   11190
   ScaleWidth      =   15285
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Generate &Report"
      Height          =   375
      Left            =   11040
      TabIndex        =   9
      Top             =   8880
      Width           =   1455
   End
   Begin VB.CommandButton cmdOperations6 
      Caption         =   "S&ort"
      Height          =   735
      Left            =   13440
      Picture         =   "frmInvoiceList.frx":58FC
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton ccmdSearch 
      Caption         =   "&Search"
      Height          =   735
      Left            =   13440
      Picture         =   "frmInvoiceList.frx":61C6
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   735
      Left            =   13440
      Picture         =   "frmInvoiceList.frx":6608
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   735
      Left            =   13440
      Picture         =   "frmInvoiceList.frx":6B92
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   735
      Left            =   13440
      Picture         =   "frmInvoiceList.frx":711C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   735
      Left            =   13440
      Picture         =   "frmInvoiceList.frx":76A6
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "E&dit"
      Height          =   735
      Left            =   13440
      Picture         =   "frmInvoiceList.frx":7C30
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid flxInvoiceList 
      Height          =   6975
      Left            =   840
      TabIndex        =   7
      Top             =   1800
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   12303
      _Version        =   393216
      BackColorFixed  =   -2147483628
      GridColor       =   0
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
   Begin VB.Line Line1 
      BorderColor     =   &H00004040&
      BorderStyle     =   4  'Dash-Dot
      X1              =   13320
      X2              =   14760
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00004040&
      BorderStyle     =   4  'Dash-Dot
      X1              =   13320
      X2              =   13320
      Y1              =   2280
      Y2              =   8280
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00004040&
      BorderStyle     =   4  'Dash-Dot
      X1              =   13320
      X2              =   14760
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00004040&
      BorderStyle     =   4  'Dash-Dot
      X1              =   14760
      X2              =   14760
      Y1              =   8280
      Y2              =   2280
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice List"
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
      Left            =   840
      TabIndex        =   8
      Top             =   960
      Width           =   11535
   End
End
Attribute VB_Name = "frmInvoiceList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsFill As ADODB.Recordset
Dim rsStatus As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim cn As ADODB.Connection
Dim Invoice As Integer

Dim xx As Integer
Dim key As String
Dim InvoiceCount As Integer

Private Sub ccmdSearch_Click()
Dim SQLText As String
Me.Enabled = False

Call Form_Click

calledFrom = 6
SQLText = "SELECT invoice_id, booking_id, fare_ref, invoice_date FROM invoice"
rs.Open SQLText, cn, adOpenKeyset
FillCombo frmSearch.Combo1, rs, False
rs.Close

Call CenterFrmSearch
frmSearch.Show
End Sub

Private Sub cmdAdd_Click()
Unload Me
Call CenterMe(frmInvoice)
frmInvoice.Show
End Sub

Private Sub cmdDelete_Click()

Dim response As Integer
Dim message, SQLText As String

If (key = "") Then
    MsgBox "You must select a row to Delete"
    Exit Sub
End If

message = "Are you sure you want to delete record with Booking Id : " & "" & Val(key) & "" & " ?"
response = MsgBox(message, vbYesNo + vbQuestion, "Confirmation")

If response = 6 Then
    SQLText = "Update trip set invoice = 'False' where trip_id = " & key
    'MsgBox SQLText
    cn.Execute SQLText

    SQLText = "Delete from invoice where booking_id = " & key
    'MsgBox SQLText
    cn.Execute SQLText
    
    SQLText = "Delete from charge_sheet where booking_id = " & key
    'MsgBox SQLText
    cn.Execute SQLText
    
Else
    Exit Sub
End If

Unload Me
Call CenterMe(frmInvoiceList)
frmInvoiceList.Show
End Sub

Private Sub cmdEdit_Click()
'Put all the data into the respective textboxes
'Calling the FillDriverForm method
'Calling the Driver Registration Form

If (key = "") Then
    MsgBox "You must select a row to edit"
    Exit Sub
End If

Call CenterMe(frmInvoice)
Call frmInvoice.FillInvoiceForm(Me.flxInvoiceList, xx)

Unload Me

End Sub

Private Sub cmdExit_Click()
Unload Me
Call EnableCmd
End Sub

Public Sub FillList(SQLText As String)

'*******************************
'***Fill the List Accordingly***
'*******************************

Dim myName As String
Dim myCustId As String


rsFill.Open SQLText, cn, adOpenKeyset

InvoiceCount = rsFill.RecordCount

With flxInvoiceList
    .Cols = 9
    .Rows = InvoiceCount + 1
    .RowHeight(0) = 400
    j = 0
    For j = 0 To 8
        .row = 0
        .Col = j
        .CellFontBold = True
        .CellAlignment = 1
        .CellFontSize = 12
        .CellFontName = "Georgia"
        .ColWidth(.Col) = 1000
    Next j

    .ColWidth(1) = 1500
    .ColWidth(2) = 2200
    .ColWidth(6) = 1200
    .ColWidth(3) = 1300
    .ColWidth(5) = 1300
        
    
    .row = 0
    
    .Col = 0
    .Text = "Inv No"
    .Col = 1
    .Text = "Book Id"
    .Col = 2
    .Text = "Name"
    .Col = 3
    .Text = "Date"
    .Col = 4
    .Text = "Rate"
    .Col = 5
    .Text = "Distance"
    .Col = 6
    .Text = "Wait"
    .Col = 7
    .Text = "Misc"
    .Col = 8
    .Text = "Total"
          
    
    i = 1
    Do While rsFill.EOF = False
        .row = i
    
       .Col = 0
        .CellAlignment = 1
        .Text = rsFill("INVOICE_ID")
        .Col = 1
        .CellAlignment = 1
        .Text = rsFill("BOOKING_ID")
        .Col = 2
        
        SQLText = "SELECT * from booking where booking_id = " & rsFill("BOOKING_ID")
        rs.Open SQLText, cn, adOpenKeyset
        myCustId = rs("CUST_ID")
        rs.Close

        SQLText = "SELECT cust_first_name, cust_last_name from customer where cust_id = " & myCustId
        rs.Open SQLText, cn, adOpenKeyset
        myName = rs("cust_first_name") & " " & rs("cust_last_name")
        rs.Close
        .Text = myName
        
        
        .Col = 3
        .CellAlignment = 1
        SQLText = "SELECT * FROM invoice WHERE invoice_id = " & rsFill("INVOICE_ID")
        rs.Open SQLText, cn, adOpenKeyset
        myInvoiceDate = Format(rs("INVOICE_DATE"), "dd-MMM-yyyy")
        rs.Close
        .Text = myInvoiceDate
        
        
        .Col = 4
        .CellAlignment = 1
        .Text = rsFill("FARE_REF")
        .Col = 5
        .CellAlignment = 1
        .Text = rsFill("DISTANCE")
        .Col = 6
        .CellAlignment = 1
        .Text = rsFill("WAIT_TIME")
        .Col = 7
        .CellAlignment = 1
        .Text = "misc"
        .Col = 8
        .CellAlignment = 1
        .Text = rsFill("TOTAL_AMOUNT")
         
        i = i + 1
        rsFill.MoveNext
        Loop
        rsFill.Close
End With

End Sub

Private Sub cmdOperations6_Click()
Dim SQLText As String
frmInvoiceList.Enabled = False

calledFrom = 6
SQLText = "SELECT invoice_id, booking_id, invoice_date, total_amount FROM invoice"
rs.Open SQLText, cn, adOpenKeyset
FillCombo frmSort.Combo1, rs, True
rs.Close

Call CenterFrmSort
frmSort.Show
End Sub

Private Sub Command1_Click()
InvoiceReport.Show
End Sub

Public Sub flxInvoiceList_Click()

If flxInvoiceList.row = 0 Then
    Exit Sub
End If
    
'flex grid click event to capture the primary key
Me.flxInvoiceList.Col = 1  'setting column to 0
key = flxInvoiceList.Text  'row is automatically set

xx = flxInvoiceList.row

'Code to highlight the entire row which ever is clicked by the user

With flxInvoiceList
    For j = 1 To (InvoiceCount)
        .row = j
        If (j = xx) Then
            For i = 0 To 8
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

Private Sub Form_Click()
key = ""
With flxInvoiceList
    For j = 1 To (InvoiceCount)
        .row = j
        For i = 0 To 8
            .Col = i
            .CellBackColor = vbWhite
        Next i
    Next j
End With
End Sub


Private Sub Form_Load()

key = ""

Dim mySQLText As String
Dim txtStat As String

Set cn = New ADODB.Connection
Set rsStatus = New ADODB.Recordset
Set rsFill = New ADODB.Recordset
Set rs = New ADODB.Recordset

cn.Open "Provider=msdasql.1; driver=Microsoft ODBC for Oracle; uid=scott; pwd=tiger;"

mySQLText = "SELECT * FROM invoice"
Call FillList(mySQLText)

End Sub

