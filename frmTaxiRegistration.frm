VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCabRegistration 
   BorderStyle     =   0  'None
   Caption         =   "Cab Registration - Blue Line "
   ClientHeight    =   11190
   ClientLeft      =   540
   ClientTop       =   -135
   ClientWidth     =   15285
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmTaxiRegistration.frx":0000
   ScaleHeight     =   11190
   ScaleWidth      =   15285
   ShowInTaskbar   =   0   'False
   Begin MSMask.MaskEdBox txtPlateNo 
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   2160
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   7
      Mask            =   ">??-####"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Go &Back"
      Height          =   735
      Left            =   6600
      Picture         =   "frmTaxiRegistration.frx":58FC
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7920
      Width           =   1215
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   735
      Left            =   2640
      Picture         =   "frmTaxiRegistration.frx":5A46
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7920
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Add"
      Height          =   735
      Left            =   4680
      Picture         =   "frmTaxiRegistration.frx":5FD0
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7920
      Width           =   1215
   End
   Begin VB.CommandButton cmdBrowse 
      Height          =   375
      Left            =   11640
      Picture         =   "frmTaxiRegistration.frx":6C9A
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4080
      Width           =   1815
   End
   Begin VB.TextBox txtCabId 
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
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   4800
      Width           =   2655
   End
   Begin VB.TextBox txtDesc 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6840
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox txtType 
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
      Left            =   2160
      TabIndex        =   8
      Top             =   5280
      Width           =   2655
   End
   Begin VB.TextBox txtColor 
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
      Left            =   6720
      TabIndex        =   7
      Top             =   4800
      Width           =   2655
   End
   Begin VB.OptionButton opAc 
      BackColor       =   &H80000014&
      Caption         =   "A/C"
      Height          =   375
      Left            =   7560
      TabIndex        =   12
      Top             =   6240
      Width           =   1095
   End
   Begin VB.OptionButton opNonAc 
      BackColor       =   &H80000014&
      Caption         =   "Non A/C"
      Height          =   255
      Left            =   7560
      TabIndex        =   13
      Top             =   6600
      Width           =   1095
   End
   Begin VB.ComboBox cmboNoSeat 
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
      ItemData        =   "frmTaxiRegistration.frx":7224
      Left            =   7080
      List            =   "frmTaxiRegistration.frx":7234
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   5280
      Width           =   1695
   End
   Begin VB.TextBox txtModel 
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
      Left            =   2280
      TabIndex        =   0
      Top             =   1680
      Width           =   2655
   End
   Begin VB.TextBox txtMake 
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
      Left            =   6840
      TabIndex        =   4
      Top             =   1680
      Width           =   2655
   End
   Begin VB.TextBox txtInsType 
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
      Left            =   2880
      TabIndex        =   2
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Frame frameStatus 
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Caption         =   "Status"
      Height          =   855
      Left            =   2280
      TabIndex        =   19
      Top             =   6000
      Width           =   1335
      Begin VB.OptionButton opMaintenance 
         BackColor       =   &H80000014&
         Caption         =   "Maintenance"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   1335
      End
      Begin VB.OptionButton opAvailable 
         BackColor       =   &H80000014&
         Caption         =   "Available"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   1335
      End
   End
   Begin MSComCtl2.DTPicker dtpickInsValidity 
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   3120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd/MMM/yyyy"
      Format          =   66060291
      CurrentDate     =   40836
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00004000&
      BorderStyle     =   4  'Dash-Dot
      X1              =   2520
      X2              =   7920
      Y1              =   7800
      Y2              =   7800
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00004000&
      BorderStyle     =   4  'Dash-Dot
      X1              =   2520
      X2              =   7920
      Y1              =   8760
      Y2              =   8760
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00004000&
      BorderStyle     =   4  'Dash-Dot
      X1              =   2520
      X2              =   2520
      Y1              =   7800
      Y2              =   8760
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00004000&
      BorderStyle     =   4  'Dash-Dot
      X1              =   7920
      X2              =   7920
      Y1              =   7800
      Y2              =   8760
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Air-Conditioning :"
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
      Left            =   5520
      TabIndex        =   31
      Top             =   6240
      Width           =   1815
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Cab Status :"
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
      Left            =   840
      TabIndex        =   30
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Image picCab 
      BorderStyle     =   1  'Fixed Single
      Height          =   2655
      Left            =   10080
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   4695
   End
   Begin VB.Line Line8 
      BorderColor     =   &H8000000C&
      X1              =   9720
      X2              =   9720
      Y1              =   4440
      Y2              =   7200
   End
   Begin VB.Line Line7 
      BorderColor     =   &H8000000C&
      X1              =   600
      X2              =   9720
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Line Line6 
      BorderColor     =   &H8000000C&
      X1              =   600
      X2              =   600
      Y1              =   4440
      Y2              =   7200
   End
   Begin VB.Line Line5 
      BorderColor     =   &H8000000C&
      X1              =   600
      X2              =   9720
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000C&
      X1              =   9720
      X2              =   9720
      Y1              =   1320
      Y2              =   3720
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000C&
      X1              =   600
      X2              =   9720
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000C&
      X1              =   600
      X2              =   600
      Y1              =   1320
      Y2              =   3720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      X1              =   600
      X2              =   9720
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Type :"
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
      Left            =   840
      TabIndex        =   29
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Color :"
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
      Left            =   5520
      TabIndex        =   28
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Desc :"
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
      Left            =   5760
      TabIndex        =   27
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cab ID :"
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
      Left            =   840
      TabIndex        =   26
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "No Of Seats :"
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
      Left            =   5520
      TabIndex        =   25
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Make :"
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
      Left            =   5760
      TabIndex        =   24
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Plate No :"
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
      Left            =   840
      TabIndex        =   23
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Model :"
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
      Left            =   840
      TabIndex        =   22
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Insurance Validity :"
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
      Left            =   840
      TabIndex        =   21
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Insurace Type :"
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
      Left            =   840
      TabIndex        =   20
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cab Registration"
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
      TabIndex        =   18
      Top             =   480
      Width           =   9135
   End
End
Attribute VB_Name = "frmCabRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mypicpath As String
Dim edit As Boolean
Dim myKey As String
Dim mcon As ADODB.Connection
Dim rs As ADODB.Recordset

Dim flgUploader As Boolean
Dim filepath As String
Dim filename As String
Dim picpath As String


Private Sub cmdBrowse_Click()

'**************************************
'***Calls the frmUploader form, sets***
'***the picture to the picture box*****
'**************************************

mypicpath = frmUploader.lblPickpath.Caption
If (mypicpath = "xxxx") Then
    flgUploader = False
    Unload frmDriverRegistration
End If

If (flgUploader = False) Then
    frmCabRegistration.Enabled = False
    
    Call frmUploader.setWhoCalled("cab")
    frmUploader.Top = 4510
    frmUploader.Left = 10440
    frmUploader.Show
    flgUploader = True
Else
    flgUploader = False
    'mypicpath = frmUploader.lblPickpath.Caption
    Unload frmUploader
    
    On Error GoTo skip
    If (mypicpath <> "") Then
        Me.picCab.Picture = LoadPicture(mypicpath)
        picpath = mypicpath
    Else
        GoTo skip
    End If
    Exit Sub
skip:
    MsgBox "Invalid Picture Format"
End If




'frameUpload.Visible = True
End Sub



Private Sub cmdClose_Click()
Unload Me
Call CenterMe(frmCabList)
frmCabList.Show
End Sub

Private Sub cmdSave_Click()

'****************************
'***SAVE / EDIT Cab Record***
'****************************

Dim myInsDate As Date
Dim currDate As Date
Dim diff As Long
Dim str1, str2 As String
Dim currPlateNo As String
Dim updatedPlateNo As String

If (edit = True) Then
    SQLText = "SELECT assigned FROM cab_status WHERE cab_id = " & myKey
    'MsgBox SQLText
    rs.Open SQLText, mcon, adOpenKeyset

    If (rs("ASSIGNED") And (Me.opMaintenance.Value)) Then
        MsgBox "De-Assign the Cab first to set Maintenance"
        rs.Close
        Exit Sub
    End If
    rs.Close
    
    SQLText = "SELECT cab_plate_no FROM cab where cab_id = " & myKey
    rs.Open SQLText, mcon, adOpenKeyset
    currPlateNo = LCase(rs("cab_plate_no"))
    updatedPlateNo = LCase(Me.txtPlateNo.Text)
    rs.Close
    
    SQLText = "SELECT cab_plate_no FROM cab"
    rs.Open SQLText, mcon, adOpenKeyset
    
    If (currPlateNo <> updatedPlateNo) Then
        Do While (rs.EOF = False)
            If (updatedPlateNo = LCase(rs("CAB_PLATE_NO"))) Then
                MsgBox "Licence Plate No Already Exists"
                rs.Close
                Exit Sub
            End If
            rs.MoveNext
        Loop
    End If
    rs.Close

End If


If (Len(Me.txtPlateNo.ClipText) < 6) Then
    MsgBox "Invalid Plate No"
    Me.txtPlateNo.SetFocus
    Exit Sub
End If

If (edit = False) Then

    SQLText = "SELECT * FROM cab"
    rs.Open SQLText, mcon, adOpenKeyset
    str1 = LCase$(Me.txtPlateNo.Text)

    Do While (rs.EOF = False)
        str2 = LCase$(rs("CAB_PLATE_NO"))
        If (str1 = str2) Then
            MsgBox "Licence Plate No already Exists"
            rs.Close
            Exit Sub
        End If
        rs.MoveNext
    Loop
rs.Close
End If

myInsDate = Format(Me.dtpickInsValidity.Value, "MM-dd-yyyy hh:mm")
currDate = Format(Now, "MM-dd-yyyy hh:mm")

'MsgBox myInsDate
'MsgBox currDate

diff = DateDiff("d", currDate, myInsDate)

If (diff < 1) Then
    MsgBox "Invalid Insurance Date"
    Exit Sub
End If


If (Me.txtColor = "" Or Me.txtDesc = "" Or Me.txtInsType = "" Or Me.txtMake = "" Or Me.txtModel = "" Or Me.txtPlateNo = "" Or Me.txtType = "") Then
    Status = MsgBox("Please fill out all the necessary information.", vbOKOnly, "Incomplete Cab Info")
    Me.txtModel.SetFocus
    Exit Sub
End If



If (Me.cmboNoSeat.Text = "") Then
    MsgBox "Select the No of Seats"
    Exit Sub
End If


'MsgBox "Picture Path: " & mypicpath
'MsgBox picpath

If (picpath = "" And edit = False) Then
    picpath = "C:\Documents and Settings\Binaya\Desktop\BlueLineTake6\Cabs\unavailable.jpeg"
End If

If (edit = False) Then
    SQLTextA = "insert into cab values(" _
    & "" & Val(Me.txtCabId.Text) & ", " _
    & "'" & Trim(Me.txtPlateNo.Text) & "', " _
    & "'" & Trim(Me.txtModel.Text) & "', " _
    & "'" & Trim(Me.txtMake.Text) & "', " _
    & "'" & Trim(Me.txtInsType.Text) & "', " _
    & "to_date('" & Format(Me.dtpickInsValidity.Value, "dd-mm-yyyy") & "', 'dd-mm-yyyy'), " _
    & "'" & Trim(Me.txtColor.Text) & "', " _
    & "'" & Trim(Me.cmboNoSeat.Text) & "', " _
    & "'" & Trim(Me.opAc.Value) & "', " _
    & "'" & Trim(Me.txtDesc.Text) & "', " _
    & "'" & Trim(Me.txtType.Text) & "', " _
    & "'" & picpath & "')"
    
    SQLTextB = "insert into cab_status values(" _
    & "" & Val(Me.txtCabId.Text) & ", " _
    & "'" & Trim(Me.opMaintenance.Value) & "', " _
    & "'" & Trim(Me.opAvailable.Value) & "', " _
    & "'False')"
    
    'MsgBox SQLTextA
    'MsgBox SQLTextB
    
    '& "'" & Trim(Me.opOnHire.Value) & "', "
    mcon.Execute SQLTextA
    mcon.Execute SQLTextB
    
Else
    
    SQLTextA = "UPDATE CAB SET " _
    & "CAB_ID = " & Val(Me.txtCabId.Text) & ", " _
    & "CAB_PLATE_NO = '" & Trim(Me.txtPlateNo.Text) & "', " _
    & "CAB_MODEL = '" & Trim(Me.txtModel.Text) & "', " _
    & "CAB_MAKE = '" & Trim(Me.txtMake.Text) & "', " _
    & "CAB_INSU_TYPE = '" & Trim(Me.txtInsType.Text) & "', " _
    & "CAB_INSU_VALIDITY = to_date('" & Format(Me.dtpickInsValidity.Value, "dd-mm-yyyy") & "', 'dd-mm-yyyy'), " _
    & "CAB_COLOR = '" & Trim(Me.txtColor.Text) & "', " _
    & "CAB_SEAT_NO = '" & Trim(Me.cmboNoSeat.Text) & "', " _
    & "CAB_AC= '" & Trim(Me.opAc.Value) & "', " _
    & "CAB_DESC = '" & Trim(Me.txtDesc.Text) & "', " _
    & "CAB_TYPE =  '" & Trim(Me.txtType.Text) & "', " _
    & "CAB_PIC = '" & picpath & "' " _
    & "WHERE CAB_ID = " & Val(myKey) & ""


    SQLTextB = "UPDATE CAB_STATUS SET " _
    & "MAINTENANCE = '" & Me.opMaintenance.Value & "', " _
    & "AVAILABLE = '" & Me.opAvailable.Value & "' " _
    & "WHERE CAB_ID = " & Me.txtCabId & " "
        
    '& "ON_HIRE = '" & Me.opOnHire.Value & "', " _

    'MsgBox SQLTextA
    'MsgBox SQLTextB
   
    mcon.Execute SQLTextA
    mcon.Execute SQLTextB
    
End If

Unload Me
mcon.Close

Call CenterMe(frmCabList)
frmCabList.Show

End Sub


Private Sub Form_Load()

'*********************
'***FORM LOAD EVENT***
'*********************

flgUploader = False
edit = False
Me.cmboNoSeat.listindex = -1
Me.opAvailable = True
Me.opNonAc = True
Me.dtpickInsValidity.Value = Now

'Me.opOnHire.Enabled = False
Set mcon = New ADODB.Connection
Set rs = New ADODB.Recordset

mcon.Open "provider=MSDASQL.1;DRIVER=MICROSOFT ODBC FOR ORACLE;UID=SCOTT;PWD=TIGER"
cabIdGenerate

Me.Visible = True
Me.txtModel.SetFocus
End Sub


Private Sub cabIdGenerate()

'**********************
'***Cab Id Generator***
'**********************

Dim rscabID As ADODB.Recordset
Set rscabID = New ADODB.Recordset

rscabID.Open "select * from cab order by cab_id desc", mcon, adOpenKeyset
If rscabID.EOF = False Then
    Me.txtCabId.Text = rscabID("cab_id") + 1
Else
    Me.txtCabId.Text = 10200
End If

rscabID.Close

End Sub


Public Sub FillCabForm(key As String)

'**************************
'***Cab Form Fill Method***
'**************************

Dim i As Integer
Dim rsA As ADODB.Recordset
Dim rsB As ADODB.Recordset

myKey = key
edit = True

Me.cmdSave.Caption = "Update"
Me.txtCabId.Locked = True
'Me.opOnHire.Enabled = True

SQLTextA = "Select * from cab where (cab_id = " & key & ")"
SQLTextB = "Select * from cab_status where (cab_id = " & key & ")"


Set rsA = New ADODB.Recordset
Set rsB = New ADODB.Recordset

rsA.Open SQLTextA, mcon, adOpenKeyset
rsB.Open SQLTextB, mcon, adOpenKeyset

Me.txtCabId.Text = rsA("CAB_ID")
Me.txtPlateNo.Text = rsA("CAB_PLATE_NO")
Me.txtModel.Text = rsA("CAB_MODEL")
Me.txtMake.Text = rsA("CAB_MAKE")
Me.txtInsType.Text = rsA("CAB_INSU_TYPE")
Me.dtpickInsValidity.Value = Format(rsA("CAB_INSU_VALIDITY"), "dd-mm-yyyy")
Me.txtColor.Text = rsA("CAB_COLOR")
Me.cmboNoSeat.Text = rsA("CAB_SEAT_NO")
    
If (rsA("CAB_AC")) Then
    Me.opAc.Value = True
Else
    Me.opNonAc.Value = True
End If


Me.txtDesc = rsA("CAB_DESC")
Me.txtType = rsA("CAB_TYPE")
    

On Error GoTo jumpError
picpath = rsA("CAB_PIC")
Me.picCab.Picture = LoadPicture(picpath)
jumpError:

'Me.opOnHire = rsB("ON_HIRE")
Me.opAvailable = rsB("AVAILABLE")
Me.opMaintenance = rsB("MAINTENANCE")

rsA.Close
rsB.Close

End Sub

