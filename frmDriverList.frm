VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDriverList 
   BorderStyle     =   0  'None
   Caption         =   "frmDriverList"
   ClientHeight    =   11190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15285
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmDriverList.frx":0000
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
      Left            =   11280
      TabIndex        =   9
      Top             =   8640
      Width           =   1455
   End
   Begin VB.CommandButton ccmdSearch 
      Caption         =   "&Search"
      Height          =   735
      Left            =   13320
      Picture         =   "frmDriverList.frx":58FC
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "S&ort"
      Height          =   735
      Left            =   13320
      Picture         =   "frmDriverList.frx":5D3E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmdDriverListDelete 
      Caption         =   "&Delete"
      Height          =   735
      Left            =   13320
      Picture         =   "frmDriverList.frx":6608
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdDriverListAddNew 
      Caption         =   "&Add"
      Height          =   735
      Left            =   13320
      Picture         =   "frmDriverList.frx":6B92
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdDriverListClose 
      Caption         =   "&Exit"
      Height          =   735
      Left            =   13320
      Picture         =   "frmDriverList.frx":711C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton cmdDriverRefresh 
      Caption         =   "&Refresh"
      Height          =   735
      Left            =   13320
      Picture         =   "frmDriverList.frx":76A6
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdDriverListEdit 
      Caption         =   "E&dit"
      Height          =   735
      Left            =   13320
      Picture         =   "frmDriverList.frx":7C30
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid flxDriverList 
      Height          =   6735
      Left            =   1680
      TabIndex        =   5
      Top             =   1800
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   11880
      _Version        =   393216
      Cols            =   6
      BackColorFixed  =   16777215
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
      X1              =   13200
      X2              =   14640
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00004040&
      BorderStyle     =   4  'Dash-Dot
      X1              =   13200
      X2              =   13200
      Y1              =   2160
      Y2              =   8160
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00004040&
      BorderStyle     =   4  'Dash-Dot
      X1              =   13200
      X2              =   14640
      Y1              =   8160
      Y2              =   8160
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00004040&
      BorderStyle     =   4  'Dash-Dot
      X1              =   14640
      X2              =   14640
      Y1              =   8160
      Y2              =   2160
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Driver List"
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
      Left            =   1680
      TabIndex        =   6
      Top             =   960
      Width           =   10935
   End
End
Attribute VB_Name = "frmDriverList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim key As String
Dim DriverCount As Integer
Dim i As Integer
Dim j As Integer

Dim rsA As ADODB.Recordset
Dim rsFill As ADODB.Recordset
Dim rsStatus As ADODB.Recordset
Dim rs As ADODB.Recordset

Dim cn As ADODB.Connection


Private Sub ccmdSearch_Click()
Dim SQLText As String
Me.Enabled = False

Call Form_Click

calledFrom = 4
SQLText = "SELECT DRIVER_ID, DRIVER_FIRST_NAME, DRIVER_LAST_NAME, " _
& "DRIVER_EMAIL, DRIVER_PHONE, DRIVER_LICENCE_NO FROM DRIVER"

rs.Open SQLText, cn, adOpenKeyset
FillCombo frmSearch.Combo1, rs, False
rs.Close

Call CenterFrmSearch
frmSearch.Show
End Sub

Private Sub cmdDriverListAddNew_Click()
Unload Me
Call CenterMe(frmDriverRegistration)
frmDriverRegistration.Show
End Sub

Private Sub cmdDriverListClose_Click()
Unload Me
Call CenterMe(frmHome)
frmHome.Show
Call EnableCmd
End Sub

Private Sub cmdDriverListDelete_Click()

Dim response As Integer
Dim message, SQLText As String

If (key = "") Then
    MsgBox "You must select a row to Delete"
    Exit Sub
End If

SQLText = "SELECT * from DRIVER where DRIVER_ID = " & key
'MsgBox SQLText
rsStatus.Open SQLText, cn, adOpenKeyset

If (rsStatus("ASSIGNED")) Then
    MsgBox "DeAssign the Driver to Delete"
    rsStatus.Close
    Exit Sub
End If
    
message = "Are you sure you want to delete record with Driver Id : " & "" & rsStatus("DRIVER_ID") & "" & " ?"
response = MsgBox(message, vbYesNo + vbQuestion, "Confirmation")

If response = 6 Then
    SQLText = "delete from driver where driver_id = " & key
    'MsgBox SQLText
    cn.Execute SQLText
Else
    rsStatus.Close
    Exit Sub
End If

Unload Me
Call CenterMe(frmDriverList)
frmDriverList.Show

End Sub

Private Sub cmdDriverListEdit_Click()
'Put all the data into the respective textboxes
'Calling the FillDriverForm method
'Calling the Driver Registration Form

If (key = "") Then
    MsgBox "You must select a row to edit"
    Exit Sub
End If

Unload Me
Call CenterMe(frmDriverRegistration)
frmDriverRegistration.Show
'Passing the primary key
Call frmDriverRegistration.FillDriverForm(key)

End Sub

Private Sub cmdSort_Click()
Dim SQLText As String
Me.Enabled = False

calledFrom = 4
SQLText = "SELECT DRIVER_ID, DRIVER_FIRST_NAME, DRIVER_LAST_NAME, DRIVER_EMAIL, DRIVER_PHONE FROM DRIVER"
rs.Open SQLText, cn, adOpenKeyset
FillCombo frmSort.Combo1, rs, True
rs.Close

Call CenterFrmSort
frmSort.Show

End Sub

Private Sub Command1_Click()
DriverReport.Show
End Sub

Private Sub flxDriverList_Click()

If flxDriverList.row = 0 Then
    Exit Sub
End If

'flex grid click event to capture the primary key
flxDriverList.Col = 0  'setting column to 0
key = flxDriverList.Text  'row is automatically set

xx = flxDriverList.row

'Code to highlight the entire row which ever is clicked by the user

With flxDriverList
    For j = 1 To (DriverCount)
        .row = j
        If (j = xx) Then
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

Private Sub Form_Click()
key = ""
With flxDriverList
    For j = 1 To (DriverCount)
        .row = j
        For i = 0 To 5
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
Set rsFill = New ADODB.Recordset
Set rsStatus = New ADODB.Recordset
Set rs = New ADODB.Recordset

cn.Open "Provider=msdasql.1; driver=Microsoft ODBC for Oracle; uid=scott; pwd=tiger;"

mySQLText = "select * from driver order by driver_id"
Call FillList(mySQLText)

End Sub


Public Sub FillList(SQLText As String)

'*******************************
'***Fill the List Accordingly***
'*******************************

rsFill.Open SQLText, cn, adOpenKeyset


DriverCount = rsFill.RecordCount
With flxDriverList
    .Rows = DriverCount + 1
    .RowHeight(0) = 400
    j = 0
    For j = 0 To 5
        .row = 0
        .Col = j
        .CellFontBold = True
        .CellAlignment = 1
        .CellFontSize = 12
        .CellFontName = "Arial"
        .ColWidth(.Col) = 2000
    Next j

    .ColWidth(0) = 1500
    .ColWidth(3) = 2500
    .ColWidth(4) = 1500
    .ColWidth(5) = 1500
    
    
    .row = 0
    
    .Col = 0
    .Text = "Driver Id"
    .Col = 1
    .Text = "First Name"
    .Col = 2
    .Text = "Last Name"
    .Col = 3
    .Text = "Email"
    .Col = 4
    .Text = "Ph No"
    .Col = 5
    .Text = "Licence No"
    
    i = 1

    Do While rsFill.EOF = False
        .row = i
    
        .Col = 0
        .CellAlignment = 1
        .Text = rsFill("driver_id")
        .Col = 1
        .Text = rsFill("driver_first_name")
        .Col = 2
        .Text = rsFill("driver_last_name")
        .Col = 3
        .Text = rsFill("driver_email")
        .Col = 4
        .CellAlignment = 1
        .Text = rsFill("driver_phone")
        .Col = 5
        .CellAlignment = 1
        .Text = rsFill("driver_licence_no")
        
        i = i + 1
        rsFill.MoveNext
    Loop
    
    rsFill.Close
End With
End Sub

