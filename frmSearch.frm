VERSION 5.00
Begin VB.Form frmSearch 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmSearch.frx":0000
   ScaleHeight     =   5205
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdLookup 
      BackColor       =   &H8000000D&
      Caption         =   "&LookUp"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H0000C0C0&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   4680
      Width           =   1095
   End
   Begin VB.ListBox lstWords 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2460
      Left            =   360
      TabIndex        =   3
      Top             =   2040
      Width           =   4335
   End
   Begin VB.TextBox txtWord 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   4335
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   315
      Left            =   360
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   735
      Width           =   4335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the Search String"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000E&
      X1              =   4800
      X2              =   4800
      Y1              =   240
      Y2              =   1200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000E&
      X1              =   240
      X2              =   240
      Y1              =   240
      Y2              =   1200
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search records by:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000E&
      X1              =   240
      X2              =   4800
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      X1              =   240
      X2              =   4800
      Y1              =   1200
      Y2              =   1200
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQLText As String
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset


Private Sub cmdCancel_Click()
Call EnableAfterSort(calledFrom)
Unload Me
End Sub

Private Sub cmdLookup_Click()
If (lstWords.Text = "") Then
    MsgBox "Select Item From the List to Look-Up"
    Exit Sub
End If

Select Case (calledFrom)

Case 1:
    Call LockItem(lstWords.Text, frmBookingList.flxBookingList)
    
Case 2:
    Call LockItem(lstWords.Text, frmCabList.flxCabList)

Case 3:
    Call LockItem(lstWords.Text, frmCustList.flxCustList)
    
Case 4:
    Call LockItem(lstWords.Text, frmDriverList.flxDriverList)
    
Case 5:
    Call LockItem(lstWords.Text, frmFareList.flxFareList)

Case 6:
    Call LockItem(lstWords.Text, frmInvoiceList.flxInvoiceList)

End Select


Call EnableAfterSort(calledFrom)
Unload Me
End Sub

Private Sub Combo1_Click()
lstWords.Clear
End Sub

Private Sub Form_Load()
Set rs = New ADODB.Recordset
Set cn = New ADODB.Connection

cn.Open "provider=MSDASQL.1;DRIVER=MICROSOFT ODBC FOR ORACLE;UID=SCOTT;PWD=TIGER"
End Sub

Private Sub txtWord_Change()
lstWords.Clear
Select Case (calledFrom)
    
    Case 0:
        MsgBox "Case 0"
                
    Case 1:
          
        SQLTextA = "SELECT tblcust.CUST_ID, tblcust.CUST_FIRST_NAME, " _
        & "tblcust.CUST_LAST_NAME, tblbook.BOOKING_ID " _
        & "from CUSTOMER tblcust, BOOKING tblbook " _
        & "where tblcust.CUST_ID = tblbook.CUST_ID"
        
        SQLText = "SELECT * FROM(" & SQLTextA & ") WHERE " _
        & "" & Combo1.Text & "" & " LIKE " & "'" & Me.txtWord.Text & "%'"
        rs.Open SQLText, cn, adOpenKeyset
        
        lstWords.Clear
        
        Do Until rs.EOF
            lstWords.AddItem rs(Combo1.Text)
            rs.MoveNext
        Loop
        
        If (lstWords.ListCount = 0) Then
            lstWords.AddItem "No Item was found"
        End If
               
        rs.Close
          
    Case 2:
             
        SQLText = "SELECT * " _
        & "FROM CAB WHERE " & "" & Combo1.Text & "" & " LIKE " & "'" & txtWord.Text & "%'"
        
        rs.Open SQLText, cn, adOpenKeyset
        lstWords.Clear
        
        Do Until rs.EOF
            lstWords.AddItem rs(Combo1.Text)
            rs.MoveNext
        Loop
        
        If (lstWords.ListCount = 0) Then
            lstWords.AddItem "No Item was found"
        End If
               
        rs.Close
        
    Case 3:
        SQLText = "SELECT CUST_ID, CUST_FIRST_NAME, CUST_LAST_NAME, CUST_EMAIL, CUST_PHONE " _
        & "FROM CUSTOMER WHERE " & "" & Combo1.Text & "" & " LIKE " & "'" & txtWord.Text & "%'"
        
        rs.Open SQLText, cn, adOpenKeyset
        lstWords.Clear
        
        Do Until rs.EOF
            lstWords.AddItem rs(Combo1.Text)
            rs.MoveNext
        Loop
        
        If (lstWords.ListCount = 0) Then
            lstWords.AddItem "No Item was found"
        End If
               
        rs.Close
        
    Case 4:
        SQLText = "SELECT DRIVER_ID, DRIVER_FIRST_NAME, DRIVER_LAST_NAME, " _
        & "DRIVER_EMAIL, DRIVER_PHONE, DRIVER_LICENCE_NO " _
        & "FROM DRIVER WHERE " & "" & Combo1.Text & "" & " LIKE " & "'" & txtWord.Text & "%'"
        
        rs.Open SQLText, cn, adOpenKeyset
        lstWords.Clear
        
        Do Until rs.EOF
            lstWords.AddItem rs(Combo1.Text)
            rs.MoveNext
        Loop
        
        If (lstWords.ListCount = 0) Then
            lstWords.AddItem "No Item was found"
        End If
               
        rs.Close
       
        
    Case 5:
        SQLText = "SELECT FARE_ID, MIN_CHARGE, PER_KM_CHARGE, ABOVE_HUNDRED_CHARGE, " _
        & "AC_CHARGE, NIGHT_CHARGE, FARE_VALIDITY, UPTO_15, ABOVE_15 " _
        & "FROM FARE WHERE " & "" & Combo1.Text & "" & " LIKE " & "'" & txtWord.Text & "%'"
        
        rs.Open SQLText, cn, adOpenKeyset
        lstWords.Clear
        
        Do Until rs.EOF
            lstWords.AddItem rs(Combo1.Text)
            rs.MoveNext
        Loop
        
        If (lstWords.ListCount = 0) Then
            lstWords.AddItem "No Item was found"
        End If
               
        rs.Close
       
        
    Case 6:
        SQLText = "SELECT * " _
        & "FROM INVOICE WHERE " & "" & Combo1.Text & "" & " LIKE " & "'" & txtWord.Text & "%'"
        
        rs.Open SQLText, cn, adOpenKeyset
        lstWords.Clear
        
        Do Until rs.EOF
            lstWords.AddItem rs(Combo1.Text)
            rs.MoveNext
        Loop
        
        If (lstWords.ListCount = 0) Then
            lstWords.AddItem "No Item was found"
        End If
               
        rs.Close
    
       
    End Select



End Sub
