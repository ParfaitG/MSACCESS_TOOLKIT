Private Sub EmailCustomer_Click()
On Error GoTo ErrHandle:

    ' ACCESS OBJECTS
    Dim db As DAO.Database
    Dim customerrst As Recordset

    ' OUTLOOK OBJECTS
    Dim objOL As Object, objEmail As Object
    Const olMailItem As Long = 0

    Dim strSQL As String, strMessage As String
    Dim i As Integer

    ' INITIALIZING OBJECTS
    Set objOL = CreateObject("Outlook.Application")
    Set objEmail = objOL.CreateItem(olMailItem)

    strSQL = "SELECT * FROM Customers WHERE CustomerID = " & Me.CustomerCbo

    Set db = CurrentDb
    Set customerrst = db.OpenRecordset(strSQL, dbOpenDynaset)

    customerrst.MoveLast
    customerrst.MoveFirst

    strMessage = "Dear " & Me.FirstName & " " & Me.LastName & ":" & vbNewLine & vbNewLine 
    strMessage = strMessage & " Please find attached the related materials concerning your orders."  & vbNewLine & vbNewLine     
    strMessage = strMessage & " Let us know if you have any questions." & vbNewLine & vbNewLine 
    strMessage = strMessage & " Regards," & vbNewLine & "Company Staff"
 
    With objEmail
        .Subject = "Customer Order"
        .Recipients.Add Me.Email
        .Body = strMessage            
        '.HTMLBody = strMessage       ' FOR HTML MESSAGE  
        .Display                      ' TO OPEN TO SCREEN
        '.Save                        ' TO DRAFTS FOLDER
        '.Send                        ' TO SEND TO RECIPIENT
	.Attachments.Add "CustomerOrderReport.pdf"
    End With
    
    ' UNINITIALIZING OBJECTS
    Set objEmail = Nothing
    Set objOL = Nothing

    customerrst.Close
    Set customerrst = Nothing
    Set db = Nothing
    
    Exit Sub

ErrHandle:
    MsgBox Err.Number & " - " & Err.Description
    Exit Sub
    
End Sub