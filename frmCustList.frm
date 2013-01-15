VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCustList 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15285
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmCustList.frx":0000
   ScaleHeight     =   11190
   ScaleWidth      =   15285
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
      Left            =   10560
      TabIndex        =   9
      Top             =   8880
      Width           =   1455
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "S&ort"
      Height          =   735
      Left            =   12960
      Picture         =   "frmCustList.frx":58FC
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   735
      Left            =   12960
      Picture         =   "frmCustList.frx":61C6
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6600
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid flxCustList 
      Height          =   6975
      Left            =   2400
      TabIndex        =   8
      Top             =   1800
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   12303
      _Version        =   393216
      Cols            =   5
      FixedRows       =   0
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
   Begin VB.CommandButton cmdCustListEdit 
      Caption         =   "E&dit"
      Height          =   735
      Left            =   12960
      Picture         =   "frmCustList.frx":6608
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdCustListEdit1 
      Caption         =   "&Refresh"
      Height          =   735
      Left            =   12960
      Picture         =   "frmCustList.frx":72D2
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdCustListClose 
      Caption         =   "&Exit"
      Height          =   735
      Left            =   12960
      Picture         =   "frmCustList.frx":785C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton cmdCustListAddNew 
      Caption         =   "&Add"
      Height          =   735
      Left            =   12960
      Picture         =   "frmCustList.frx":7DE6
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdCustListDelete 
      Caption         =   "&Delete"
      Height          =   735
      Left            =   12960
      Picture         =   "frmCustList.frx":8370
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00004040&
      BorderStyle     =   4  'Dash-Dot
      X1              =   14280
      X2              =   14280
      Y1              =   8280
      Y2              =   2280
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00004040&
      BorderStyle     =   4  'Dash-Dot
      X1              =   12840
      X2              =   14280
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00004040&
      BorderStyle     =   4  'Dash-Dot
      X1              =   12840
      X2              =   12840
      Y1              =   2280
      Y2              =   8280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004040&
      BorderStyle     =   4  'Dash-Dot
      X1              =   12840
      X2              =   14280
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Customer List"
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
      Left            =   2400
      TabIndex        =   7
      Top             =   960
      Width           =   9495
   End
End
Attribute VB_Name = "frmCustList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim key As String
Dim recLen As Integer
Dim i As Integer
Dim j As Integer

Dim CustCount As Integer

Dim cn As ADODB.Connection
Dim rsFill As ADODB.Recordset
Dim rs As ADODB.Recordset



Private Sub ccmdSearch_Click()
Dim SQLText As String
Me.Enabled = False

calledFrom = 2
SQLText = "SELECT CAB_ID, CAB_PLATE_NO, CAB_MAKE FROM CAB"
rs.Open SQLText, cn, adOpenKeyset
FillCombo frmSearch.Combo1, rs, False
rs.Close

Call CenterFrmSearch
frmSearch.Show
End Sub

Private Sub cmdCustListClose_Click()
Unload Me
Call EnableCmd
End Sub

Private Sub cmdCustListAddNew_Click()
Unload Me
Call CenterMe(frmCustRegistration)
frmCustRegistration.Show
End Sub

Private Sub cmdCustListDelete_Click()

Dim response As Integer
Dim message, SQLText As String


If (key = "") Then
    MsgBox "You must select a row to Delete!!!"
    Exit Sub
End If


SQLText = "SELECT booking_id, booking_status FROM booking WHERE cust_id = " & key
'MsgBox SQLText
rs.Open SQLText, cn, adOpenKeyset

Do While (rs.EOF = False)
    If (rs("BOOKING_STATUS") = "ONGO") Then
        MsgBox ("Currently Customer is in service Unable to Delete")
        rs.Close
        Exit Sub
    End If
    rs.MoveNext
Loop
rs.Close

'******Customer is not currently in-serive procede to delete**********'

message = "Are you sure you want to delete record with Customer Id : " & "" & Val(key) & "" & " ?"
response = MsgBox(message, vbYesNo + vbQuestion, "Confirmation")

If response = 6 Then
    SQLText = "DELETE from CUSTOMER where CUST_ID = " & key
    'MsgBox SQLText
    cn.Execute SQLText
Else
    Exit Sub
End If


Unload Me
Call CenterMe(frmCustList)
frmCustList.Show

End Sub

Private Sub cmdCustListEdit_Click()
'Put all the data into the respective textboxes
'Calling the FillDriverForm method
'Calling the Driver Registration Form

If (key = "") Then
    MsgBox "You must select a row to edit"
    Exit Sub
End If

Unload Me
Call CenterMe(frmCustRegistration)
frmCustRegistration.Show
'Passing the primary key
Call frmCustRegistration.FillCustForm(key)

End Sub

Private Sub cmdSearch_Click()
Dim SQLText As String
Me.Enabled = False

Call Form_Click

calledFrom = 3
SQLText = "SELECT CUST_ID, CUST_FIRST_NAME, CUST_LAST_NAME, CUST_EMAIL, CUST_PHONE FROM CUSTOMER"
rs.Open SQLText, cn, adOpenKeyset
FillCombo frmSearch.Combo1, rs, False
rs.Close

Call CenterFrmSearch
frmSearch.Show
End Sub

Private Sub cmdSort_Click()
Dim SQLText As String
frmCustList.Enabled = False

calledFrom = 3
SQLText = "SELECT CUST_ID, CUST_FIRST_NAME, CUST_LAST_NAME, CUST_EMAIL, CUST_PHONE FROM CUSTOMER"
rs.Open SQLText, cn, adOpenKeyset
FillCombo frmSort.Combo1, rs, True
rs.Close

Call CenterFrmSort
frmSort.Show
End Sub

Private Sub Command1_Click()
CustReport.Show
End Sub

Private Sub flxCustList_Click()

If flxCustList.row = 0 Then
    Exit Sub
End If
    
'flex grid click event to capture the primary key
Me.flxCustList.Col = 0  'setting column to 0
key = flxCustList.Text  'row is automatically set

xx = flxCustList.row

'Code to highlight the entire row which ever is clicked by the user

With flxCustList
    For j = 1 To (CustCount)
        .row = j
        If (j = xx) Then
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

Private Sub Form_Click()
key = ""
With flxCustList
    For j = 1 To (CustCount)
        .row = j
        For i = 0 To 4
            .Col = i
            .CellBackColor = vbWhite
        Next i
    Next j
End With
End Sub


Private Sub Form_Load()

'*********************
'***FORM LOAD EVENT***
'*********************

key = ""
Dim mySQLText As String

Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset
Set rsFill = New ADODB.Recordset

cn.Open "Provider=msdasql.1; driver=Microsoft ODBC for Oracle; uid=scott; pwd=tiger;"

mySQLText = "select * from customer order by cust_id"

Call FillList(mySQLText)

End Sub



Public Sub FillList(SQLText As String)

'*******************************
'***Fill the List Accordingly***
'*******************************

rsFill.Open SQLText, cn, adOpenKeyset
CustCount = rsFill.RecordCount

With flxCustList
    .Rows = CustCount + 1
    .RowHeight(0) = 400
    j = 0
    For j = 0 To 4
        .row = 0
        .Col = j
        .Text = "test"
        .CellFontBold = True
        .CellAlignment = 1
        .CellFontSize = 12
        .CellFontName = "Georgia"
        .ColWidth(.Col) = 2000
    Next j

    .ColWidth(0) = 1000
    .ColWidth(3) = 3000
    .ColWidth(4) = 1500
    
    
    .row = 0
    
    .Col = 0
    .Text = "Cust Id"
    .Col = 1
    .Text = "First Name"
    .Col = 2
    .Text = "Last Name"
    .Col = 3
    .Text = "Email"
    .Col = 4
    .Text = "Ph No"
    
    
    
    i = 1
    Do While rsFill.EOF = False
        .row = i
    
        .Col = 0
        .CellAlignment = 1
        .Text = rsFill("CUST_ID")
        .Col = 1
        .Text = rsFill("CUST_FIRST_NAME")
        .Col = 2
        .Text = rsFill("CUST_LAST_NAME")
        .Col = 3
        .Text = rsFill("CUST_EMAIL")
        .Col = 4
        .CellAlignment = 1
        .Text = rsFill("CUST_PHONE")
        
        i = i + 1
        rsFill.MoveNext
    Loop
    rsFill.Close
End With

End Sub
