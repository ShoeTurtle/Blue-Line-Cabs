VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmFareList 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15285
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmFareList.frx":0000
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
      Left            =   11880
      TabIndex        =   9
      Top             =   8760
      Width           =   1455
   End
   Begin VB.CommandButton ccmdSearch 
      Caption         =   "&Search"
      Height          =   735
      Left            =   13680
      Picture         =   "frmFareList.frx":58FC
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "S&ort"
      Height          =   735
      Left            =   13680
      Picture         =   "frmFareList.frx":5D3E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmdFareEdit 
      Caption         =   "E&dit"
      Height          =   735
      Left            =   13680
      Picture         =   "frmFareList.frx":6608
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdFareRefresh 
      Caption         =   "&Refresh"
      Height          =   735
      Left            =   13680
      Picture         =   "frmFareList.frx":72D2
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdFareClose 
      Caption         =   "&Exit"
      Height          =   735
      Left            =   13680
      Picture         =   "frmFareList.frx":785C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton cmdFareAdd 
      Caption         =   "&Add"
      Height          =   735
      Left            =   13680
      Picture         =   "frmFareList.frx":7DE6
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdFareDelete 
      Caption         =   "&Delete"
      Height          =   735
      Left            =   13680
      Picture         =   "frmFareList.frx":8370
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3960
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid flxFareList 
      Height          =   6855
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   12091
      _Version        =   393216
      BackColorFixed  =   -2147483628
      GridColor       =   0
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
   Begin VB.Line Line8 
      BorderColor     =   &H00004040&
      BorderStyle     =   4  'Dash-Dot
      X1              =   15000
      X2              =   15000
      Y1              =   8160
      Y2              =   2160
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00004040&
      BorderStyle     =   4  'Dash-Dot
      X1              =   13560
      X2              =   15000
      Y1              =   8160
      Y2              =   8160
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00004040&
      BorderStyle     =   4  'Dash-Dot
      X1              =   13560
      X2              =   13560
      Y1              =   2160
      Y2              =   8160
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00004040&
      BorderStyle     =   4  'Dash-Dot
      X1              =   13560
      X2              =   15000
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cab On-Hire Charge List"
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
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   12015
   End
End
Attribute VB_Name = "frmFareList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim key As String
Dim FareCount As Integer

Dim rs As ADODB.Recordset
Dim rsFill As ADODB.Recordset

Dim cn As ADODB.Connection

Private Sub ccmdSearch_Click()
Dim SQLText As String
Me.Enabled = False

Call Form_Click

calledFrom = 5
SQLText = "SELECT FARE_ID, MIN_CHARGE, PER_KM_CHARGE, ABOVE_HUNDRED_CHARGE, " _
& "AC_CHARGE, NIGHT_CHARGE, FARE_VALIDITY, UPTO_15, ABOVE_15 FROM FARE"

rs.Open SQLText, cn, adOpenKeyset
FillCombo frmSearch.Combo1, rs, False
rs.Close

Call CenterFrmSearch
frmSearch.Show
End Sub

Private Sub cmdFareAdd_Click()
Unload Me
Call CenterMe(frmFare)
frmFare.Show
End Sub

Private Sub cmdFareClose_Click()

Call CenterMe(frmHome)
frmHome.Show

Unload Me
Call EnableCmd

End Sub

Private Sub cmdFareDelete_Click()

'**************************
'***Deleting Fare Reords***
'**************************

If (key = "") Then
    MsgBox "You must select a row to edit"
    Exit Sub
End If

SQLText = "delete from FARE where FARE_ID = " & key
'MsgBox SQLText

cn.Execute SQLText
Unload Me
Call CenterMe(frmFareList)
frmFareList.Show

End Sub

Private Sub cmdFareEdit_Click()
'Put all the data into the respective textboxes
'Calling the FillDriverForm method
'Calling the Driver Registration Form

If (key = "") Then
    MsgBox "You must select a row to edit"
    Exit Sub
End If

Unload Me
Call CenterMe(frmFare)
frmFare.Show
'Passing the primary key
Call frmFare.FillFareForm(key)

End Sub

Private Sub cmdSort_Click()
Dim SQLText As String
Me.Enabled = False

calledFrom = 5
SQLText = "SELECT * FROM FARE"
rs.Open SQLText, cn, adOpenKeyset
FillCombo frmSort.Combo1, rs, True
rs.Close

Call CenterFrmSort
frmSort.Show

End Sub

Private Sub Command1_Click()
FareReport.Show
End Sub

Private Sub flxFareList_Click()

If flxFareList.row = 0 Then
    Exit Sub
End If
    
'flex grid click event to capture the primary key
Me.flxFareList.Col = 0  'setting column to 0
key = flxFareList.Text  'row is automatically set

xx = flxFareList.row

'Code to highlight the entire row which ever is clicked by the user

With flxFareList
    For j = 1 To (FareCount)
        .row = j
        If (j = xx) Then
            For i = 0 To 8
                .Col = i
                .CellBackColor = &H80FF80
            Next i
        Else
            For i = 0 To 8
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
Dim mySQLText As String

Set cn = New ADODB.Connection
Set rsFill = New ADODB.Recordset
Set rs = New ADODB.Recordset

cn.Open "Provider=msdasql.1; driver=Microsoft ODBC for Oracle; uid=scott; pwd=tiger;"

mySQLText = "select * from fare order by fare_id"

Call FillList(mySQLText)

End Sub

Private Sub Form_Click()
key = ""
With flxFareList
    For j = 1 To (FareCount)
        .row = j
        For i = 0 To 8
            .Col = i
            .CellBackColor = vbWhite
        Next i
    Next j
End With
End Sub


Public Sub FillList(SQLText As String)

'*******************************
'***Fill the List Accordingly***
'*******************************

rsFill.Open SQLText, cn, adOpenKeyset

FareCount = rsFill.RecordCount

flxFareList.Rows = recLen
With flxFareList
    .Rows = FareCount + 1
    .Cols = 9
    j = 0
    .RowHeight(0) = 400
    For j = 0 To 8
        .row = 0
        .Col = j
        '.Text = "test"
        .CellFontBold = True
        .CellAlignment = 1
        .CellFontSize = 12
        .CellFontName = "Georgia"
        .ColWidth(j) = 1500
    Next j

    .ColWidth(0) = 1300
    
    .row = 0
    
    .Col = 0
    .Text = "Fare Id"
    .Col = 1
    .Text = "Minimum"
    .Col = 2
    .Text = "Per Km"
    .Col = 3
    .Text = "> 100 Km"
    .Col = 4
    .Text = "A/C Xtra"
    .Col = 5
    .Text = "Night Xtra"
    .Col = 6
    .Text = "Validity"
    .Col = 7
    .Text = "Till 15 min"
    .Col = 8
    .Text = "> 15 min"
        
    i = 1
    Do While rsFill.EOF = False
        .row = i

        .Col = 0
        .CellAlignment = 1
        .Text = rsFill("FARE_ID")
        .Col = 1
        .CellAlignment = 1
        .Text = rsFill("MIN_CHARGE")
        .Col = 2
        .CellAlignment = 1
        .Text = rsFill("PER_KM_CHARGE")
        .Col = 3
        .CellAlignment = 1
        .Text = rsFill("ABOVE_HUNDRED_CHARGE")
        .Col = 4
        .CellAlignment = 1
        .Text = rsFill("AC_CHARGE")
        .Col = 5
        .CellAlignment = 1
        .Text = rsFill("NIGHT_CHARGE")
        .Col = 6
        .CellAlignment = 1
        .Text = rsFill("FARE_VALIDITY")
        .Col = 7
        .CellAlignment = 1
        .Text = rsFill("UPTO_15")
        .Col = 8
        .CellAlignment = 1
        .Text = rsFill("ABOVE_15")
        
        i = i + 1
        rsFill.MoveNext
    Loop
    rsFill.Close
End With

End Sub


