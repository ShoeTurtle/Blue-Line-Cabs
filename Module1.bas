Attribute VB_Name = "Module1"
Option Explicit
Global currUserName As String
Global AsAdmin As Boolean
Global UsernameAdmin, PasswordAdmin As String
Global login As Boolean
Global calledbySearch As Boolean
Global calledFrom As Integer


Public Sub CenterMe(frmTest As Form)
    With frmTest
        .Left = 0
        .Top = 0
    End With
End Sub
    
Public Sub DisableCmd(Index As Integer)
    Dim i As Integer
    
    For i = 0 To 9
        If (i <> Index) Then
            MDIMain.cmdMain(i).Enabled = False
        End If
    Next
End Sub


Public Sub EnableCmd()
    Dim i As Integer
    
    If (AsAdmin) Then
        For i = 0 To 9 Step 1
            MDIMain.cmdMain(i).Enabled = True
        Next
    Else
        For i = 1 To 9
            MDIMain.cmdMain(i).Enabled = True
        Next i
    End If
            
End Sub

Public Sub CenterFrmSort()
    With frmSort
        .Width = 5595
        .Height = 2595
        .Left = (MDIMain.ScaleWidth - .Width) / 2
        .Top = (MDIMain.ScaleHeight - .Height) / 2
        .Show
    End With
End Sub

Public Sub CenterFrmSearch()
    With frmSearch
        .Width = 5040
        .Height = 5205
        .Left = (MDIMain.ScaleWidth - .Width) / 2
        .Top = (MDIMain.ScaleHeight - .Height) / 2
        .Show
    End With
End Sub



Public Sub FillCombo(ByRef sCombo As ComboBox, ByVal sRS As Recordset, Sort As Boolean)

'This procedure fills a combo box with field name from a given recordset
'used in the combo boxes for Searching/Filtering/Sorting records

Dim X As Long

With sCombo
    For X = 0 To sRS.Fields.Count - 1
        If Sort Then
            .AddItem "" & sRS.Fields.Item(X).Name & "" & " Asc"
            .AddItem "" & sRS.Fields.Item(X).Name & "" & " Desc"
        Else 'NOT SORT...
            .AddItem sRS.Fields.Item(X).Name
        End If
    Next X
End With

sCombo.listindex = 0
End Sub

Public Sub SortNow(listno As Integer, SortBy As String)

Dim mytable, SQLText As String
Dim k, listindex As Integer

Select Case listno
Case 1:
        
        'MsgBox "BookingTable"
        mytable = "Booking"
                        
        If SortBy = "" Then
            SQLText = "SELECT * FROM " & mytable
        Else
            SQLText = "SELECT * FROM " & mytable & " Order By " & SortBy & ""
        End If
        
        'MsgBox SQLText
        
        Call frmBookingList.FillList(SQLText)
        
         
Case 2:
        
        'MsgBox "Cab Table"
        mytable = "Cab"
        If SortBy = "" Then
            SQLText = "SELECT * FROM " & mytable
        Else
            SQLText = "SELECT * FROM " & mytable & " Order By " & SortBy & ""
        End If
        
        'MsgBox SQLText
        
        Call frmCabList.FillList(SQLText)
        
            
           

Case 3:
        'MsgBox "Customer Table"
        mytable = "CUSTOMER"
        If SortBy = "" Then
            SQLText = "SELECT * FROM " & mytable
        Else
            SQLText = "SELECT * FROM " & mytable & " Order By " & SortBy & ""
        End If
        
        'MsgBox SQLText
        
        Call frmCustList.FillList(SQLText)

Case 4:
        'MsgBox "Driver Table"
        mytable = "DRIVER"
        If SortBy = "" Then
            SQLText = "SELECT * FROM " & mytable
        Else
            SQLText = "SELECT * FROM " & mytable & " Order By " & SortBy & ""
        End If
        
        'MsgBox SQLText
        
        Call frmDriverList.FillList(SQLText)
        
        
        
Case 5:
        'MsgBox "Fare Table"
        mytable = "FARE"
        If SortBy = "" Then
            SQLText = "SELECT * FROM " & mytable
        Else
            SQLText = "SELECT * FROM " & mytable & " Order By " & SortBy & ""
        End If
        
        'MsgBox SQLText
        
        Call frmFareList.FillList(SQLText)
        
        
        
Case 6:
        'MsgBox "Invoice Table"
        mytable = "INVOICE"
        
        If SortBy = "" Then
            SQLText = "SELECT * FROM " & mytable
        Else
            SQLText = "SELECT * FROM " & mytable & " Order By " & SortBy & ""
        End If
        
        'MsgBox SQLText
        
        Call frmInvoiceList.FillList(SQLText)
        
Case Default

End Select

Set mytable = Nothing
'calledFrom = 0
End Sub

Public Sub EnableAfterSort(X As Integer)
    

    Select Case (X)
    
    Case 0:
        'MsgBox "called From = 0"
        
    Case 1:
        frmBookingList.Enabled = True
        If (calledbySearch = False) Then
            Call frmBookingList.Form_Click
        End If
                  
    Case 2:
        frmCabList.Enabled = True
        
    Case 3:
        frmCustList.Enabled = True
        
    Case 4:
        frmDriverList.Enabled = True
        
    Case 5:
        frmFareList.Enabled = True
        
    Case 6:
        frmInvoiceList.Enabled = True
    End Select
    calledFrom = 0
    calledbySearch = False
    
End Sub

Public Sub LockItem(SearchKey As String, myFlx As MSFlexGrid)

'****************************************
'***String Matching with the grid text***
'****************************************

Dim i, j As Integer
j = myFlx.Cols
i = myFlx.Rows

If (calledbySearch = True) Then
    Call frmBookingList.Form_Click
End If


With myFlx
    For i = 0 To .Rows - 1
        .row = i
        For j = 0 To .Cols - 1
            .Col = j
            If (InStr(1, myFlx.Text, SearchKey, vbTextCompare)) Then
               .CellBackColor = &HFFFF&
            End If
        Next j
    Next i
End With

End Sub

Public Sub MeCenter(xWidth As Integer, yHeight As Integer, myFrm As Form)
    With myFrm
        .Width = xWidth
        .Height = yHeight
        .Left = (MDIMain.ScaleWidth - .Width) / 2
        .Top = (MDIMain.ScaleHeight - .Height) / 2
        .Show
    End With
End Sub
