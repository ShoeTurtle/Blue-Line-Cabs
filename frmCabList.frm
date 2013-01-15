VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCabList 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15285
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmCabList.frx":0000
   ScaleHeight     =   11190
   ScaleWidth      =   15285
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
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
      Left            =   11400
      TabIndex        =   9
      Top             =   8880
      Width           =   1455
   End
   Begin VB.CommandButton ccmdSearch 
      Caption         =   "&Search"
      Height          =   735
      Left            =   13200
      Picture         =   "frmCabList.frx":58FC
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "S&ort"
      Height          =   735
      Left            =   13200
      Picture         =   "frmCabList.frx":5D3E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   735
      Left            =   13200
      Picture         =   "frmCabList.frx":6608
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "&Add"
      Height          =   735
      Left            =   13200
      Picture         =   "frmCabList.frx":6B92
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Exit"
      Height          =   735
      Left            =   13200
      Picture         =   "frmCabList.frx":711C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   735
      Left            =   13200
      Picture         =   "frmCabList.frx":76A6
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "E&dit"
      Height          =   735
      Left            =   13200
      Picture         =   "frmCabList.frx":7C30
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid flxCabList 
      Height          =   6975
      Left            =   600
      TabIndex        =   8
      Top             =   1800
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   12303
      _Version        =   393216
      Cols            =   6
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
   Begin VB.Line Line1 
      BorderColor     =   &H00004040&
      BorderStyle     =   4  'Dash-Dot
      X1              =   13080
      X2              =   14520
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00004040&
      BorderStyle     =   4  'Dash-Dot
      X1              =   13080
      X2              =   13080
      Y1              =   1920
      Y2              =   8040
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00004040&
      BorderStyle     =   4  'Dash-Dot
      X1              =   13080
      X2              =   14520
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00004040&
      BorderStyle     =   4  'Dash-Dot
      X1              =   14520
      X2              =   14520
      Y1              =   7920
      Y2              =   1920
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cab List"
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
      Left            =   600
      TabIndex        =   7
      Top             =   960
      Width           =   12255
   End
End
Attribute VB_Name = "frmCabList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection

Dim rsStatus As ADODB.Recordset
Dim rsFill As ADODB.Recordset
Dim rs As ADODB.Recordset

Dim CabCount As Integer
Dim key As String

Private Sub ccmdSearch_Click()
Dim SQLText As String
Me.Enabled = False

Call Form_Click

calledFrom = 2
SQLText = "SELECT CAB_ID, CAB_PLATE_NO, CAB_MAKE, CAB_MODEL FROM CAB"
rs.Open SQLText, cn, adOpenKeyset
FillCombo frmSearch.Combo1, rs, False
rs.Close

Call CenterFrmSearch
frmSearch.Show
End Sub

Private Sub cmdAddNew_Click()
Unload Me
Call CenterMe(frmCabRegistration)
frmCabRegistration.Show
End Sub

Private Sub cmdClose_Click()
Unload Me
Call CenterMe(frmHome)
frmHome.Show
Call EnableCmd
End Sub

Private Sub cmdDelete_Click()

Dim response As Integer
Dim message, SQLText As String

If (key = "") Then
    MsgBox "You must select a row to Delete"
    Exit Sub
End If

SQLText = "SELECT * from CAB_STATUS where CAB_ID = " & key
rsStatus.Open SQLText, cn, adOpenKeyset

If (rsStatus("ASSIGNED")) Then
    MsgBox "De-Assign Cab to delete"
    rsStatus.Close
    Exit Sub
End If
    
message = "Are you sure you want to delete cab with cab id : " & "" & rsStatus("CAB_ID") & "" & " ?"
response = MsgBox(message, vbYesNo + vbQuestion, "Confirmation")

If response = 6 Then
    SQLText = "delete from cab where cab_id = " & key
    'MsgBox SQLText
    cn.Execute SQLText
Else
    rsStatus.Close
    Exit Sub
End If

rsStatus.Close
Unload Me
Call CenterMe(frmCabList)
frmCabList.Show

End Sub

Private Sub cmdEdit_Click()
'Put all the data into the respective textboxes
'Calling the FillDriverForm method
'Calling the Cab Registration Form

If (key = "") Then
    MsgBox "You must select a row to edit"
    Exit Sub
End If

Unload Me
Call CenterMe(frmCabRegistration)
frmCabRegistration.Show
'Passing the primary key
Call frmCabRegistration.FillCabForm(key)

End Sub

Private Sub cmdSort_Click()
Dim SQLText As String
frmCabList.Enabled = False

calledFrom = 2
SQLText = "SELECT CAB_ID, CAB_PLATE_NO, CAB_MODEL, CAB_INSU_VALIDITY FROM CAB"
rs.Open SQLText, cn, adOpenKeyset
FillCombo frmSort.Combo1, rs, True
rs.Close

Call CenterFrmSort
frmSort.Show

End Sub

Private Sub Command1_Click()
DataReport1.Show
End Sub

Private Sub Command2_Click()
CabReport.Show
End Sub

Private Sub flxCabList_Click()
If flxCabList.row = 0 Then
    Exit Sub
End If

'flex grid click event to capture the primary key
Me.flxCabList.Col = 0  'setting column to 0
key = flxCabList.Text  'row is automatically set

xx = flxCabList.row

'Code to highlight the entire row which ever is clicked by the user

With flxCabList
    For j = 1 To (CabCount)
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
With flxCabList
    For j = 1 To (CabCount)
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
'***Form Load Event***
'*********************

key = ""

Dim mySQLText As String
Dim txtStat As String

Set cn = New ADODB.Connection
Set rsStatus = New ADODB.Recordset
Set rsFill = New ADODB.Recordset
Set rs = New ADODB.Recordset

cn.Open "Provider=msdasql.1; driver=Microsoft ODBC for Oracle; uid=scott; pwd=tiger;"

mySQLText = "SELECT * FROM cab"
Call FillList(mySQLText)

End Sub

Public Sub FillList(SQLText As String)

'*******************************
'***Fill the List Accordingly***
'*******************************
Dim txtStat As String

rsFill.Open SQLText, cn, adOpenKeyset

CabCount = rsFill.RecordCount
With flxCabList
    .Rows = CabCount + 1
    .RowHeight(0) = 400
    j = 0
    For j = 0 To 5
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
    .ColWidth(1) = 1300
    .ColWidth(2) = 2200
    .ColWidth(3) = 2200
    .ColWidth(4) = 1500
    .ColWidth(5) = 4300

  
    .row = 0
    
    .Col = 0
    .Text = "Cab Id"
    .Col = 1
    .Text = "Plate No"
    .Col = 2
    .Text = "Make"
    .Col = 3
    .Text = "Model"
    .Col = 4
    .Text = "Status"
    .Col = 5
    .Text = "Description"
    
                 
    i = 1
    Do While rsFill.EOF = False
        .row = i
    
        .Col = 0
        .CellAlignment = 1
        .Text = rsFill("CAB_ID")
        .Col = 1
        .CellAlignment = 1
        .Text = rsFill("CAB_PLATE_NO")
        .Col = 2
        .CellAlignment = 1
        .Text = rsFill("CAB_MAKE")
        .Col = 3
        .CellAlignment = 1
        .Text = rsFill("CAB_MODEL")
        .Col = 4
        .CellAlignment = 1
                
        currCabid = Val(rsFill("CAB_ID"))
        SQLText = "SELECT * from CAB_STATUS where (CAB_ID = " & currCabid & ")"
            
        rsStatus.Open SQLText, cn, adOpenKeyset
        
        If (rsStatus("MAINTENANCE")) Then
            txtStat = "Maintenance"
        ElseIf (rsStatus("AVAILABLE")) Then
            txtStat = "Available"
        End If
        rsStatus.Close
        
        SQLText = "SELECT ONHIRE from CAB_DRIVER_ASSIGN where (CAB_ID = " & currCabid & ")"
        rsStatus.Open SQLText, cn, adOpenKeyset
        If (rsStatus.RecordCount = 1) Then
            If (rsStatus("ONHIRE")) Then
                txtStat = "On-Hire"
            End If
        End If
        rsStatus.Close
        
        .Text = txtStat
        .Col = 5
        .Text = rsFill("CAB_DESC")
        
        i = i + 1
        rsFill.MoveNext
    Loop
    rsFill.Close
End With
End Sub
