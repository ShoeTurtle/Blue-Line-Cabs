VERSION 5.00
Begin VB.Form frmCommandCenter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Command Center"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmCommandCenter.frx":0000
   ScaleHeight     =   5640
   ScaleWidth      =   5730
   Begin VB.ListBox lstAssigned 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      Left            =   3600
      TabIndex        =   6
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton cmdDAssign 
      Caption         =   "[-] D-Assign"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton cmdAssign 
      Caption         =   "[+] Assign "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   4320
      Width           =   1575
   End
   Begin VB.ListBox lstAvailDrivers 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      Left            =   2040
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.ListBox lstAvailCabs 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      Left            =   600
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Line Line10 
      BorderColor     =   &H8000000C&
      BorderWidth     =   2
      X1              =   3480
      X2              =   3480
      Y1              =   240
      Y2              =   4080
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cab/Driver View"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   375
      Left            =   3600
      TabIndex        =   9
      Top             =   360
      Width           =   1695
   End
   Begin VB.Line Line9 
      BorderColor     =   &H8000000C&
      X1              =   480
      X2              =   5400
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line8 
      BorderColor     =   &H8000000C&
      X1              =   480
      X2              =   5400
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line7 
      BorderColor     =   &H8000000C&
      X1              =   480
      X2              =   480
      Y1              =   240
      Y2              =   4080
   End
   Begin VB.Line Line6 
      BorderColor     =   &H8000000C&
      X1              =   5400
      X2              =   5400
      Y1              =   240
      Y2              =   4080
   End
   Begin VB.Line Line5 
      BorderColor     =   &H8000000C&
      X1              =   480
      X2              =   5400
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000C&
      X1              =   840
      X2              =   5040
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000C&
      X1              =   840
      X2              =   840
      Y1              =   4200
      Y2              =   5280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000C&
      X1              =   5040
      X2              =   5040
      Y1              =   4200
      Y2              =   5280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      X1              =   840
      X2              =   5040
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Avail-Cab"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Avail-Driver"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "frmCommandCenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim myCab As String
Dim myDriver As String
Dim myAssignedCab As String
Dim myAssignedDriver As String

Dim mcon As ADODB.Connection
Dim rs As ADODB.Recordset

Private Sub cmdAddDriver_Click()
Dim driverID As String
driverID = InputBox("Enter the Driver ID: ", "Add Driver")

lstAvailDrivers.AddItem driverID

End Sub

Private Sub cmdAssign_Click()

'*************************
'***Assign Cab / Driver***
'*************************


If (myCab = "") Then
    MsgBox "Select a Cab from the Avail-Cab List"
    Exit Sub
ElseIf (myDriver = "") Then
    MsgBox "Select a Driver form the Avail-Driver List"
    Exit Sub
End If

SQLText = "INSERT into CAB_DRIVER_ASSIGN values(" _
& "" & Val(myCab) & ", " _
& "" & Val(myDriver) & ", " _
& "'False')"

mcon.Execute SQLText

SQLText = "UPDATE CAB_STATUS SET " _
& "ASSIGNED = 'True' " _
& "WHERE CAB_ID = " & Val(myCab) & " "

mcon.Execute SQLText

SQLText = "UPDATE DRIVER SET " _
& "ASSIGNED = 'True' " _
& "WHERE DRIVER_ID = " & Val(myDriver) & " "
mcon.Execute SQLText


Unload frmHome
Call CenterMe(frmHome)
frmHome.Show

Unload Me
With frmCommandCenter
        .Width = 6075
        .Height = 6060
        .Left = (MDIMain.ScaleWidth - .Width) / 2
        .Top = (MDIMain.ScaleHeight - .Height) / 2
        .Show
End With

End Sub

Private Sub cmdClose_Click()
Unload Me
frmHome.Enabled = True
Call EnableCmd
End Sub

Private Sub cmdDAssign_Click()

'*****************
'***DeAssigning***
'*****************

If (myAssignedCab = "") Then
    MsgBox "Select from Cab-Driver List"
    Exit Sub
ElseIf (myAssignedDriver = "") Then
    MsgBox "Select from Cab-Driver List"
    Exit Sub
End If



SQLText = "SELECT * from cab_driver_assign where cab_id = " & "" & Val(myAssignedCab) & ""
'MsgBox SQLText

rs.Open SQLText, mcon, adOpenKeyset

If (rs("ONHIRE")) Then
    MsgBox "Unable to De-Assign, the cab/driver is already dispatched"
    rs.Close
    Exit Sub
End If
rs.Close

SQLText = "UPDATE DRIVER SET " _
& "ASSIGNED = 'False' " _
& "WHERE DRIVER_ID = " & Val(myAssignedDriver) & " "
mcon.Execute SQLText

SQLText = "UPDATE CAB_STATUS SET " _
& "ASSIGNED = 'False' " _
& "WHERE CAB_ID = " & Val(myAssignedCab) & " "
mcon.Execute SQLText

SQLText = "DELETE from CAB_DRIVER_ASSIGN " _
& "WHERE DRIVER_ID = " & Val(myAssignedDriver) & " "
mcon.Execute SQLText


Unload frmHome
Call CenterMe(frmHome)
frmHome.Show

Unload Me
With frmCommandCenter
        .Width = 6075
        .Height = 6060
        .Left = (MDIMain.ScaleWidth - .Width) / 2
        .Top = (MDIMain.ScaleHeight - .Height) / 2
        .Show
End With

End Sub

Private Sub cmdUpdate_Click()

'*************
'***REFRESH***
'*************

Unload Me

With frmCommandCenter
        .Width = 6075
        .Height = 6060
        .Left = (MDIMain.ScaleWidth - .Width) / 2
        .Top = (MDIMain.ScaleHeight - .Height) / 2
        .Show
    End With


End Sub

Private Sub Form_Load()

myCab = ""
myDriver = ""
myAssignedCab = ""
myAssignedDriver = ""

'*********************
'***FORM LOAD EVENT***
'*********************

Set rs = New ADODB.Recordset
Set mcon = New ADODB.Connection

mcon.Open "provider=MSDASQL.1;DRIVER=MICROSOFT ODBC FOR ORACLE;UID=SCOTT;PWD=TIGER"

SQLText = "SELECT DRIVER_ID from DRIVER where (ASSIGNED = 'False')"

rs.Open SQLText, mcon, adOpenKeyset
'rs.MoveFirst
Do Until rs.EOF
    lstAvailDrivers.AddItem rs("DRIVER_ID")
    rs.MoveNext
Loop
rs.Close

SQLText = "SELECT CAB_ID FROM CAB_STATUS where (AVAILABLE = 'True' AND ASSIGNED = 'False')"
rs.Open SQLText, mcon, adOpenKeyset
'rs.MoveFirst
Do Until rs.EOF
    lstAvailCabs.AddItem rs("CAB_ID")
    rs.MoveNext
Loop
rs.Close

SQLText = "SELECT CAB_ID, DRIVER_ID from CAB_DRIVER_ASSIGN "
rs.Open SQLText, mcon, adOpenKeyset
'rs.MoveFirst
Do Until rs.EOF
    lstAssigned.AddItem rs("CAB_ID") & "     " & rs("DRIVER_ID")
    rs.MoveNext
Loop
rs.Close

End Sub

Private Sub lstAssigned_Click()

Me.lstAvailCabs.listindex = -1
Me.lstAvailDrivers.listindex = -1

myAssignedCabDriver = lstAssigned.Text

myAssignedCab = Trim(Mid$(lstAssigned.Text, 1, 5))
myAssignedDriver = Trim(Mid$(lstAssigned.Text, 9, 13))

'MsgBox "CAB-ID: " & myAssignedCab
'MsgBox "DRIVER-ID: " & myAssignedDriver


End Sub

Private Sub lstAvailCabs_Click()
Me.lstAssigned.listindex = -1
myCab = lstAvailCabs.Text
End Sub

Private Sub lstAvailDrivers_Click()
Me.lstAssigned.listindex = -1
myDriver = lstAvailDrivers.Text
End Sub
