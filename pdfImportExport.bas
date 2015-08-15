Private Sub GeneratePDF_Click()
On Error GoTo ErrHandle
    Dim db As DAO.Database
    Dim prefillcontactrst As DAO.Recordset
    Dim adobeObj As Object, srcDoc As Object, AvDoc As Object, pdfDoc As Object, jso As Object
    Dim strPath As String
    Dim i As Integer, j As Integer
    Dim processResult As Boolean
        
    ' CREATES FOLDER IF DOES NOT EXIST FOR OUTPUTTED PDF FORMS
    If Dir(Application.CurrentProject.Path & "\ClientForms", vbDirectory) = vbNullString Then
        MkDir (Application.CurrentProject.Path & "\ClientForms")
    End If
    
    strPath = "C\Path\To\PDFFormTemplate.pdf"
    
    ' RETRIEVING CLIENT INFO
    Set db = CurrentDb
    db.Execute "UPDATE OneTable SET PlaceHolder = PlaceHolder + 1", dbFailOnError
    Set prefillcontactrst = db.OpenRecordset("SELECT * FROM Contractors WHERE ID = '" & Me.ClientCbo & "'", dbOpenDynaset)
    
    prefillcontactrst.MoveLast
    prefillcontactrst.MoveFirst
      
    ' INITIALIZING PDF OBJECTS
    Set adobeObj = CreateObject("AcroExch.App")
    Set srcDoc = CreateObject("AcroExch.PDDoc")
    Set AvDoc = CreateObject("AcroExch.AVDoc")
    srcDoc.Open strPath
   
    ' POPULATING PDF FILE WITH CLIENT INFO
    If AvDoc.Open(strPath, "") = True Then
        Set pdfDoc = AvDoc.GetPDDoc()
        Set jso = pdfDoc.GetJSObject()

            ' COMPANY INFO
            jso.GetField("CompanyName").Value = "Database, Incorporated"
            jso.GetField("CompanyStreet").Value = "101 Main Street"
            jso.GetField("CompanyCityStateZip").Value = "Cityville, ST 99999"
            jso.GetField("CompanyPhone").Value = "(999) 999-9999"
            jso.GetField("CompanyEmail").Value = "email@example.com"
            jso.GetField("CompanyEmail").Value = "email@example.com"
            jso.GetField("SalesPerson").Value = "Jane Doe"
            jso.GetField("PaymentTerms").Value = "Due upon receipt"
                        
          
            ' CLIENT INFO
            jso.GetField("ContractorName").Value = CStr(prefillcontactrst![ContactorName])          
            jso.GetField("ContractorStreet").Value = CStr(prefillcontactrst![ContractorStreet])
            jso.GetField("ContractorCityStateZip").Value = CStr(prefillcontactrst![ContractorCityStateZip])
            jso.GetField("ContractorPhone").Value = CStr(prefillcontactrst![ContractorPhone])
            jso.GetField("ContractorEmail").Value = CStr(prefillcontactrst![ContractorEmail])
  

            jso.GetField("CustomerID").Value = CStr(prefillcontactrst![ContractorID])
            jso.GetField("InvoiceNo").Value = DMax("InvoiceID", "Invoices") + 1
            jso.GetField("Job").Value = CStr(prefillcontactrst![JobDescription])

    End If
    
    ' SAVING PDF FILE
    pdfDoc.Save &H1, Application.CurrentProject.Path & "\ClientForms\" & CStr(prefillcontactrst![ContractorName]) & "_Invoice" & GenerateFormID & ".pdf"
    pdfDoc.Close
            
    ' UNINTIALIZING PDF OBJECTS
    Set adobeObj = Nothing
    Set srcDoc = Nothing
    Set AvDoc = Nothing
    Set pdfDoc = Nothing
    Set jso = Nothing
    
    prefillcontactrst.Close
    Set prefillcontactrst = Nothing
    Set db = Nothing

    DoCmd.Close acForm, "Loading"
    
    MsgBox "Document editing complete!", vbInformation
    DoCmd.SetWarnings True
    Exit Sub
    
ErrHandle:
    MsgBox Err.Number & " - " & Err.Description, vbCritical
    DoCmd.Close acForm, "Loading"
    DoCmd.SetWarnings True
    
    Set adobeObj = Nothing
    Set srcDoc = Nothing
    Set AvDoc = Nothing
    Set pdfDoc = Nothing
    Set jso = Nothing

    prefillcontactrst.Close
    Set prefillcontactrst = Nothing
    Set db = Nothing
    Exit Sub
End Sub


Private Sub ImportPDF_Click()
On Error GoTo ErrHandle
    Dim db As Database
    Dim strFile As String, ClientID As String, insertSQL As String, cleanDate As String, facilityNoCheck As String
    Dim adobeObj As Object, srcDoc As Object, AvDoc As Object, pdfDoc As Object, jso As Object
    Dim processResult As Boolean
    Dim i As Integer, f As Integer
    Dim intResult As Variant
        
    Set db = CurrentDb
        
    strFile = FilePathFind()
       
    If Len(strFile) = 0 Then
        Exit Sub
    End If
    
    db.Execute "Delete * FROM tmpInvoices", dbFailOnError
    db.Execute "Delete * FROM tmpInvoiceItems", dbFailOnError
    
    ' INTIALIZING PDF OBJECTS
    Set adobeObj = CreateObject("AcroExch.App")
    Set srcDoc = CreateObject("AcroExch.PDDoc")
    Set AvDoc = CreateObject("AcroExch.AVDoc")
    srcDoc.Open strFile
 
    If AvDoc.Open(strFile, "") = True Then
        Set pdfDoc = AvDoc.GetPDDoc()
        Set jso = pdfDoc.GetJSObject()
                    
        ' IMPORTING CLIENT LEVEL DATA               
        insertSQL = "INSERT INTO tmpInvoices ([ContractorID], [InvoiceDate], [InvoiceNo], [DueDate])"
        insertSQL = insertSQL & " VALUES ("
        insertSQL = insertSQL & " '" & jso.GetField("ContractorID") & "', "
        insertSQL = insertSQL & " #" & jso.GetField("InvoiceDate").Value & "#, "
        insertSQL = insertSQL & " '" & jso.GetField("InvoiceNo").Value & "', "
        insertSQL = insertSQL & " '" & jso.GetField("DueDate").Value & "');"
              
        db.Execute insertSQL, dbFailOnError
        
        ' IMPORTING MULTIPLE SAMPLE DATA
        For i = 1 To 20 
	    invoicenum = jso.GetField("InvoiceNo").Value        
            quantity = jso.GetField("QUANTITYRow" & i).Value
            description = jso.GetField("DESCRIPTIONRow" & i).Value
            unitprice = jso.GetField("UNIT PRICERow" & i).Value
            lineprice = jso.GetField("LINE TOTALRow" & i).Value
            
            insertSQL = "INSERT INTO tmpInvoiceItems (InvoiceID, Quantity, Description, UnitPrice, LinePrice)
            insertSQL = insertSQL & " VALUES ('" & invoicenum & "', "
            insertSQL = insertSQL & " " & quantity & ", "
            insertSQL = insertSQL & " '" & description & "', "
            insertSQL = insertSQL & " " & unitprice & ", "
            insertSQL = insertSQL & " " & lineprice & ");"                        
            db.Execute insertSQL, dbFailOnError
        Next i
        
    End If
    
    AvDoc.Close True
    srcDoc.Close

    ' CLEANING UP MISSINGS
    db.Execute "DELETE * FROM tmpInvoiceItems WHERE Len([Quantity]) = 0 ", dbFailOnError

        
    ' UNINTIALIZING PDF OBJECTS
    Set adobeObj = Nothing
    Set srcDoc = Nothing
    Set AvDoc = Nothing
    Set pdfDoc = Nothing
    Set jso = Nothing
    Set db = Nothing
    Exit Sub
    
ErrHandle:
    MsgBox Err.Number & " - " & Err.Description, vbCritical
    DoCmd.Close acForm, "Loading"
    DoCmd.SetWarnings True
    
    AvDoc.Close True
    srcDoc.Close
    Set adobeObj = Nothing
    Set srcDoc = Nothing
    Set AvDoc = Nothing
    Set pdfDoc = Nothing
    Set jso = Nothing
    Set db = Nothing
    Exit Sub
End Sub


Public Function FilePathFind() As String
On Error GoTo ErrHandle
    Dim fd As Object
    Const msoFileDialogFilePicker = 3
    Dim strInputFileName As Variant
    Dim strFilePath, folder As String
    Dim vrtSelectedItem As Variant
        
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    DoCmd.SetWarnings False

    With fd
        .Title = "Browse for Service Invoice Forms"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "PDF files", "*.pdf"
        .FilterIndex = 1
        .InitialFileName = Application.CurrentProject.Path
        If .Show = -1 Then
                strFilePath = .SelectedItems(1)
        Else
            'The user pressed Cancel.
            MsgBox "No file Selected", vbExclamation
            strFilePath = Null
            Set fd = Nothing
            Exit Function
        End If
    End With
        
    Set fd = Nothing
    
    FilePathFind = strFilePath
    Exit Function
    
ErrHandle:
    MsgBox Err.Number & " - " & Err.Description
    Set fd = Nothing
    Exit Function
    
End Function


Public Function GenerateFormID() As String
    Select Case Len(DLookup("PlaceHolder", "OneTable"))
    
        Case 1
        GenerateFormID = "0000" & DLookup("PlaceHolder", "OneTable")
    
        Case 2
        GenerateFormID = "000" & DLookup("PlaceHolder", "OneTable")
    
        Case 3
        GenerateFormID = "00" & DLookup("PlaceHolder", "OneTable")
    
        Case 4
        GenerateFormID = "0" & DLookup("PlaceHolder", "OneTable")
    
        Case 5
        GenerateFormID = DLookup("PlaceHolder", "OneTable")
        
        Case Is > 5
        GenerateFormID = DLookup("PlaceHolder", "OneTable")
        
    End Select
    
End Function
