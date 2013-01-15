VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   0  'None
   ClientHeight    =   2010
   ClientLeft      =   2790
   ClientTop       =   3045
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmLoginScreen.frx":0000
   ScaleHeight     =   134
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   301
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtUserName 
      Height          =   300
      Left            =   1650
      TabIndex        =   1
      Top             =   315
      Width           =   2505
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Log-On"
      Default         =   -1  'True
      Height          =   390
      Left            =   270
      TabIndex        =   3
      Top             =   1425
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   390
      Left            =   3180
      TabIndex        =   4
      Top             =   1425
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1650
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   795
      Width           =   2505
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   1095
      Left            =   270
      Top             =   180
      Width           =   4020
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      Height          =   1095
      Left            =   285
      Top             =   195
      Width           =   4020
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   270
      Index           =   0
      Left            =   465
      TabIndex        =   0
      Top             =   375
      Width           =   1050
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   270
      Index           =   1
      Left            =   465
      TabIndex        =   5
      Top             =   855
      Width           =   915
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Private mUsername As String
Private mPassword As String
Private mCancel As Boolean

Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdOK_Click()

mCancel = False
mUsername = Me.txtUserName.Text
mPassword = Me.txtPassword.Text
        
SQLText = "SELECT * from admin"
rs.Open SQLText, cn, adOpenKeyset
   
   
Do While (rs.EOF = False)
    If (rs("ADMIN_USERNAME") = mUsername And rs("ADMIN_PASSWORD") = mPassword) Then
        Module1.currUserName = mUsername
        AsAdmin = True
        For i = 0 To 9
            MDIMain.cmdMain(i).Enabled = True
        Next i
        
        
        Call CenterMe(frmHome)
        frmHome.Show
        frmHome.txtMsgBoard.Text = rs("ADMIN_NOTES")
        rs.Close
        Unload Me
        Exit Sub
    End If
    rs.MoveNext
Loop
rs.Close

SQLText = "SELECT * FROM user_table"
rs.Open SQLText, cn, adOpenKeyset
        
Do While (rs.EOF = False)
    If (rs("USER_NAME") = mUsername And rs("USER_PASSWORD") = mPassword) Then
            Module1.currUserName = mUsername
            AsAdmin = False
            For i = 1 To 9
                MDIMain.cmdMain(i).Enabled = True
            Next i
            
            Call CenterMe(frmHome)
            frmHome.Show
            frmHome.txtMsgBoard.Text = rs("USER_NOTE")
            rs.Close
            Unload Me
            Exit Sub
    End If
    rs.MoveNext
Loop
rs.Close
    

MsgBox "Invalid Username/Password", vbExclamation
Me.txtPassword = ""
Me.txtUserName = ""
Me.txtUserName.SetFocus
End Sub

'Public Function GetLogIn(ByRef UserName As String, ByRef Password As String, Owner As Object) As Boolean
'    Me.txtUserName.Text = UserName
'
'    Me.Show vbModal, Owner
'
'    UserName = mUsername
'    Password = mPassword
'
'    GetLogIn = Not mCancel
'End Function

'Private Sub Form_Activate()
'    If Len(Me.txtUserName.Text) > 0 Then Me.txtPassword.SetFocus
'End Sub


Private Sub Form_Load()
Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset

cn.Open "Provider=msdasql.1; driver=Microsoft ODBC for Oracle; uid=scott; pwd=tiger;"




End Sub

