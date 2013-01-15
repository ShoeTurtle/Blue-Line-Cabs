VERSION 5.00
Begin VB.Form frmManageUsers 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   3840
      TabIndex        =   12
      Top             =   2400
      Width           =   960
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   2880
      Width           =   960
   End
   Begin VB.ListBox lstUsers 
      Height          =   1425
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   2040
      Width           =   960
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   3840
      TabIndex        =   8
      Top             =   3240
      Width           =   960
   End
   Begin VB.TextBox txtUserName 
      Height          =   330
      Left            =   1245
      TabIndex        =   1
      Top             =   2160
      Width           =   2295
   End
   Begin VB.TextBox txtPassword1 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1245
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2880
      Width           =   2295
   End
   Begin VB.TextBox txtPassword2 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1245
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   3240
      Width           =   2295
   End
   Begin VB.TextBox txtOldPassword 
      Enabled         =   0   'False
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1245
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5040
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "User Name:"
      Height          =   315
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   1200
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Confirmation :"
      Height          =   300
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   1140
   End
   Begin VB.Label lblOldPassword 
      AutoSize        =   -1  'True
      Caption         =   "Old Password:"
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   1140
   End
End
Attribute VB_Name = "frmManageUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rsData As ADODB.Recordset

Private Sub cmdCancel_Click()
Unload Me
Call MeCenter(5010, 3705, frmManageUsers)
frmManageUsers.Show
End Sub

Private Sub cmdDelete_Click()

SQLText = "DELETE from USER_TABLE where user_name = '" & Me.txtUserName.Text & "'"
cn.Execute SQLText
cmdCancel_Click

End Sub

Private Sub cmdExit_Click()
For i = 0 To 9
    MDIMain.cmdMain(i).Enabled = True
Next i
Unload Me

Call CenterMe(frmHome)
frmHome.Show

End Sub

Private Sub Form_Load()

Set rsData = New ADODB.Recordset
Set cn = New ADODB.Connection

cn.Open "provider=MSDASQL.1;DRIVER=MICROSOFT ODBC FOR ORACLE;UID=SCOTT;PWD=TIGER"

ListUsers

End Sub
Private Sub ListUsers()
   
SQLText = "SELECT * from USER_TABLE"
rsData.Open SQLText, cn, adOpenKeyset
   
Me.lstUsers.Clear
Do While (rsData.EOF = flase)
    Me.lstUsers.AddItem rsData("USER_NAME")
    rsData.MoveNext
Loop
rsData.Close
End Sub
Private Sub cmdAdd_Click()

Dim Request As String, NewID As Long
Dim NewPassword As String, OldPassword As String
    
If Len(txtPassword1.Text) > 0 Or Len(txtPassword2.Text) > 0 Then
    If txtPassword1.Text <> txtPassword2.Text Then
        MsgBox "Confirm password must the same as Password field", vbExclamation
        Exit Sub
    End If
End If

NewPassword = Me.txtPassword1.Text
OldPassword = Me.txtOldPassword.Text
    

If cmdAdd.Caption = "&Add" Then
    
    Dim rsUserID As ADODB.Recordset
    Set rsUserID = New ADODB.Recordset
    
    SQLText = "select * from USER_TABLE order by USER_ID desc"
    
    'MsgBox SQLText
    rsUserID.Open SQLText, cn, adOpenKeyset
    If rsUserID.EOF = False Then
        NewID = rsUserID("USER_ID") + 1
    Else
        NewID = 1
    End If
    
    
    If (Me.txtUserName = "") Then
        MsgBox "Enter the User-Name"
        Exit Sub
    End If
    
    If (Me.txtPassword1 = "") Then
        MsgBox "Enter the Password and Confirm it!!!"
        Exit Sub
    End If
    
        
    SQLText = "INSERT INTO user_table VALUES(" & NewID & ", '" _
    & Format(Me.txtUserName.Text) & "', '" & NewPassword & "'," _
    & " 'Welcome to The Cab!!!') "
    'MsgBox SQLText
    cn.Execute SQLText
    
Else
    SQLText = "SELECT * FROM user_table WHERE user_name = '" & Me.lstUsers.Text & "'"
    'MsgBox SQLText
    rsData.Open SQLText, cn, adOpenKeyset
    
    If OldPassword <> rsData("USER_PASSWORD") Then
        MsgBox "Invalid old password." & vbNewLine & "You must enter the valid password for user selected.", vbInformation
        rsData.Close
        Exit Sub
    End If
    rsData.Close
    
    SQLText = "UPDATE user_table SET user_password = '" & NewPassword & "'"
    MsgBox SQLText
    cn.Execute SQLText
End If

cmdAdd.Caption = "&Add"
    
ListUsers
txtUserName.Text = ""
txtPassword1.Text = ""
txtPassword2.Text = ""
Me.txtOldPassword.Text = ""

txtOldPassword.Enabled = False
lblOldPassword.Enabled = False

lstUsers.Enabled = True
  
End Sub

Private Sub lstUsers_DblClick()

Me.txtUserName.Text = Me.lstUsers.Text
lstUsers.Enabled = False
cmdAdd.Caption = "&Update"
lblOldPassword.Enabled = True
txtOldPassword.Enabled = True
       
    'If LogInUserID <> Val(lstUsers.SelectedItem.Text) Then
        '  txtOldPassword.Enabled = True
        '  lblOldPassword.Enabled = True
    'End If
    
End Sub

