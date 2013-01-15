VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCustRegistration 
   BorderStyle     =   0  'None
   Caption         =   "Driver Registration - Blue Line"
   ClientHeight    =   9825
   ClientLeft      =   420
   ClientTop       =   -105
   ClientWidth     =   15285
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmCustomerRegistration.frx":0000
   ScaleHeight     =   9825
   ScaleWidth      =   15285
   ShowInTaskbar   =   0   'False
   Begin MSMask.MaskEdBox txtCustPinNo 
      Height          =   375
      Left            =   9240
      TabIndex        =   10
      Top             =   5760
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   6
      Mask            =   "######"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtCustPhone 
      Height          =   375
      Left            =   9240
      TabIndex        =   4
      Top             =   3120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##########"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtCustHouseNo 
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
      Left            =   4200
      TabIndex        =   5
      Top             =   4560
      Width           =   1575
   End
   Begin VB.TextBox txtCustState 
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
      Left            =   4200
      TabIndex        =   9
      Top             =   5760
      Width           =   2655
   End
   Begin VB.TextBox txtCustStreet 
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
      Left            =   9240
      TabIndex        =   6
      Top             =   4800
      Width           =   2655
   End
   Begin VB.TextBox txtCustCity 
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
      Left            =   9240
      TabIndex        =   8
      Top             =   5280
      Width           =   2655
   End
   Begin VB.TextBox txtCustBlock 
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
      Left            =   4200
      TabIndex        =   7
      Top             =   5280
      Width           =   2655
   End
   Begin VB.TextBox txtCustID 
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
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox txtCustLN 
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
      Left            =   9240
      TabIndex        =   2
      Top             =   2640
      Width           =   2775
   End
   Begin VB.TextBox txtCustEmail 
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
      Left            =   4200
      TabIndex        =   3
      Top             =   3120
      Width           =   2895
   End
   Begin VB.TextBox txtCustFN 
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
      Left            =   4200
      TabIndex        =   1
      Top             =   2640
      Width           =   2895
   End
   Begin VB.CommandButton cmdCusSave 
      Caption         =   "&Add"
      Height          =   735
      Left            =   6720
      Picture         =   "frmCustomerRegistration.frx":58FC
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton cmdCustListEdit1 
      Caption         =   "&Refresh"
      Height          =   735
      Left            =   4680
      Picture         =   "frmCustomerRegistration.frx":65C6
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton cmdCusClose 
      Caption         =   "Go &Back"
      Height          =   735
      Left            =   8640
      Picture         =   "frmCustomerRegistration.frx":6B50
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000C&
      X1              =   12360
      X2              =   12360
      Y1              =   4200
      Y2              =   6360
   End
   Begin VB.Line Line11 
      BorderColor     =   &H8000000C&
      X1              =   2160
      X2              =   2160
      Y1              =   6360
      Y2              =   4200
   End
   Begin VB.Line Line10 
      BorderColor     =   &H8000000C&
      X1              =   2160
      X2              =   12360
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line Line9 
      BorderColor     =   &H8000000C&
      X1              =   2160
      X2              =   12360
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line8 
      BorderColor     =   &H8000000C&
      X1              =   12360
      X2              =   12360
      Y1              =   1560
      Y2              =   3720
   End
   Begin VB.Line Line7 
      BorderColor     =   &H8000000C&
      X1              =   2160
      X2              =   2160
      Y1              =   3720
      Y2              =   1560
   End
   Begin VB.Line Line6 
      BorderColor     =   &H8000000C&
      X1              =   2160
      X2              =   12360
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line5 
      BorderColor     =   &H8000000C&
      X1              =   2160
      X2              =   12360
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "House No :"
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
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label12 
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
      Left            =   7680
      TabIndex        =   24
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "City :"
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
      Left            =   7680
      TabIndex        =   23
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Pin No :"
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
      Left            =   7680
      TabIndex        =   22
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label15 
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
      Left            =   2400
      TabIndex        =   21
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "State :"
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
      TabIndex        =   20
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer ID :"
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
      TabIndex        =   19
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name :"
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
      TabIndex        =   18
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name :"
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
      Left            =   7680
      TabIndex        =   17
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No :"
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
      Left            =   7680
      TabIndex        =   16
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Email :"
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
      TabIndex        =   15
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00004000&
      BorderStyle     =   4  'Dash-Dot
      X1              =   9960
      X2              =   9960
      Y1              =   7200
      Y2              =   8160
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00004000&
      BorderStyle     =   4  'Dash-Dot
      X1              =   4560
      X2              =   4560
      Y1              =   7200
      Y2              =   8160
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00004000&
      BorderStyle     =   4  'Dash-Dot
      X1              =   4560
      X2              =   9960
      Y1              =   8160
      Y2              =   8160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004000&
      BorderStyle     =   4  'Dash-Dot
      X1              =   4560
      X2              =   9960
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Registration"
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
      TabIndex        =   14
      Top             =   720
      Width           =   10215
   End
End
Attribute VB_Name = "frmCustRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim myKey As String
Dim edit As Boolean

Dim mcon As ADODB.Connection

Private Sub cmdCusClose_Click()

If (flgcmdGoCust) Then
    Unload Me
    Call CenterMe(frmBooking)
    frmBooking.Show
Else
    Unload Me
    Call CenterMe(frmCustList)
    frmCustList.Show
End If
End Sub

Private Sub cmdCusSave_Click()

'*******************************
'**SAVE / EDIT Customer Record**
'*******************************

If (Me.txtCustFN.Text = "" Or Me.txtCustLN = "") Then
    MsgBox "Customer Full Name Required"
    Exit Sub
End If

If (Me.txtCustEmail.Text = "") Then
    MsgBox "Enter the email address"
    Me.txtCustEmail.SetFocus
    Exit Sub
End If

If (InStr(1, Me.txtCustEmail.Text, "@", vbTextCompare) = 0) Then
    MsgBox "Invalid Email Address"
    Me.txtCustEmail.SetFocus
    Exit Sub
End If

If (Val(Me.txtCustPhone.Text) = 0) Then
    MsgBox "Enter Customer Phone No"
    Me.txtCustPhone.SetFocus
    Exit Sub
End If

If ((Me.txtCustBlock.Text = "") Or (Me.txtCustCity.Text = "") Or (Me.txtCustHouseNo.Text = "") Or (Val(Me.txtCustPinNo.Text) = 0) Or (Me.txtCustState.Text = "")) Then
    MsgBox "Few Address Items Missing, All Fields Are Mandatory"
    Me.txtCustHouseNo.SetFocus
    Exit Sub
End If

'***the leading 0 is truncated*****

testVal = Val(Me.txtCustPinNo.Text)
Dim testStr As String
testStr = testVal
'MsgBox testStr


If (Len(testStr) < 6) Then
    MsgBox "Invalid Pin No"
    Me.txtCustPinNo.SetFocus
    Exit Sub
End If

testVal = Val(Me.txtCustPhone.Text)
testStr = testVal

    
If (Len(testStr) < 10) Then
    MsgBox "Invalid Phone No"
    Me.txtCustPhone.SetFocus
    Exit Sub
End If

If IsNumeric(Me.txtCustHouseNo.Text) = False Then
    MsgBox ("Please enter a number for House No")
    txtCustHouseNo.Text = ""
    txtCustHouseNo.SetFocus
    Exit Sub
End If


If (edit = False) Then
    SQLTextA = "insert into customer values(" _
    & "" & Val(Me.txtCustID.Text) & ", " _
    & "'" & Trim(Me.txtCustFN.Text) & "', " _
    & "'" & Trim(Me.txtCustLN.Text) & "', " _
    & "'" & Trim(Me.txtCustEmail.Text) & "', " _
    & "" & Val(Me.txtCustPhone) & ", " _
    & "" & Val(Me.txtCustHouseNo.Text) & ", " _
    & "'" & Trim(Me.txtCustStreet.Text) & "', " _
    & "'" & Trim(Me.txtCustBlock.Text) & "', " _
    & "'" & Trim(Me.txtCustCity.Text) & "', " _
    & "'" & Trim(Me.txtCustState.Text) & "', " _
    & "" & Val(Me.txtCustPinNo.Text) & ")"
    
    mcon.Execute SQLTextA
    
Else

    SQLTextB = "update customer set " _
    & "CUST_ID = " & Val(Me.txtCustID.Text) & ", " _
    & "CUST_FIRST_NAME = '" & Trim(Me.txtCustFN.Text) & "', " _
    & "CUST_LAST_NAME = '" & Trim(Me.txtCustLN.Text) & "', " _
    & "CUST_EMAIL = '" & Trim(Me.txtCustEmail.Text) & "', " _
    & "CUST_PHONE = " & Val(Me.txtCustPhone) & ", " _
    & "CUST_HOUSE_NO = " & Val(Me.txtCustHouseNo.Text) & ", " _
    & "CUST_STREET = '" & Trim(Me.txtCustStreet.Text) & "', " _
    & "CUST_BLOCK = '" & Trim(Me.txtCustBlock.Text) & "', " _
    & "CUST_CITY= '" & Trim(Me.txtCustCity.Text) & "', " _
    & "CUST_STATE = '" & Trim(Me.txtCustState.Text) & "', " _
    & "CUST_PIN_NO = " & Val(Me.txtCustPinNo.Text) & " " _
    & "WHERE CUST_ID = " & Val(myKey) & ""

    mcon.Execute SQLTextB
    
End If

Unload Me

Call CenterMe(frmCustList)
frmCustList.Show
End Sub

Private Sub Form_Load()

'*******************
'**Form Load Event**
'*******************

edit = False

Set mcon = New ADODB.Connection
mcon.Open "provider=MSDASQL.1;DRIVER=MICROSOFT ODBC FOR ORACLE;UID=SCOTT;PWD=TIGER"
custIdGenerate

Me.Visible = True
Me.txtCustFN.SetFocus
End Sub


Private Sub custIdGenerate()
Dim rscustID As ADODB.Recordset
Set rscustID = New ADODB.Recordset
rscustID.Open "select * from customer order by cust_id desc", mcon, adOpenKeyset
If rscustID.EOF = False Then
    Me.txtCustID = rscustID("cust_id") + 1
Else
    Me.txtCustID.Text = 10200
End If
End Sub

Public Sub FillCustForm(key As String)

'*****************************
'**Customer Form Fill Method**
'*****************************

Dim i As Integer
Dim rs As ADODB.Recordset

myKey = key
edit = True
Me.cmdCusSave.Caption = "Update"
Me.txtCustID.Locked = True

SQLTextA = "Select * from customer where (cust_id = " & key & ")"


Set rs = New ADODB.Recordset

rs.Open SQLTextA, mcon, adOpenKeyset

With Me
    .txtCustID = rs("cust_id")
    .txtCustFN = rs("cust_first_name")
    .txtCustLN = rs("cust_last_name")
    .txtCustEmail = rs("cust_email")
    .txtCustPhone = rs("cust_phone")
    .txtCustHouseNo = rs("cust_house_no")
    .txtCustStreet = rs("cust_street")
    .txtCustBlock = rs("cust_block")
    .txtCustCity = rs("cust_city")
    .txtCustState = rs("cust_state")
    .txtCustPinNo = rs("cust_pin_no")
    
End With

rs.Close

End Sub


