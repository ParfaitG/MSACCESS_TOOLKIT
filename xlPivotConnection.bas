Private Sub ExportDashboard_Click()
On Error GoTo ErrHandle
    ' ACCESS OBJECTS
    Dim db As DAO.Database
    Dim datarst As DAO.Recordset

    ' EXCEL OBJECTS
    Dim xlObj As Object, xlWkb As Object, tblWks As Object, dataWks As Object, pivotWks As Object
    Dim xlSheet As Object, xlPivot As Object
    Dim lastrow As Long, i As Integer
    Const xlToLeft = -4159, xlUp = -4162, xlContinuous = 1, xlThick = 4, xlEdgeTop = 8
    Const xlDatabase = 1
    

    ' INITIALIZE OBJECTS
    Set db = CurrentDb
    Set xlObj = CreateObject("Excel.Application")
    Set xlWkb = xlObj.Workbooks.Open(Application.CurrentProject.Path & "\DashboardTemplate.xlsx")
    Set tblWks = xlWkb.WorkSheets("TABLE")
    Set dataWks = xlWkb.WorkSheets("REVENUES")
    Set pivotWks = xlWkb.WorkSheets("DEPT PIVOT")
    
    
    ' DATA WORKSHEET
    Set datarst = db.OpenRecordset("PivotData")
    lastrow = dataWks.Cells(dataWks.Rows.Count, "A").End(xlUp).Row
    dataWks.Range("A2:H" & lastrow + 5).Delete xlToLeft
        
    ' OUTPUTTING RECORDSET
    dataWks.Range("A2").CopyFromRecordset datarst

    ' FORMATTING SPECIFIC COLUMNS
    dataWks.Range("D2:G" & datarst.RecordCount + 1).NumberFormat = "$#,##0.00"
    
    ' CREATE EXPORT FILE PATH IF DOES NOT EXIST
    If Dir(Application.CurrentProject.Path & "\Exports", vbDirectory) = vbNullString Then
        MkDir (Application.CurrentProject.Path & "\Exports")
    End If

    ' DELETE PREVIOUS SAME DAY EXPORT FILE IF EXISTS
    If Not Dir(Application.CurrentProject.Path & "\Exports\DashboardAnalytics_" & Format(Date, "MMDDYYYY") & ".xlsx") = vbNullString Then
        Kill Application.CurrentProject.Path & "\Exports\DashboardAnalytics_" & Format(Date, "MMDDYYYY") & ".xlsx"
    End If
        

    ' PIVOT WORKSHEET
    pivotWks.PivotTables("RevenuesPivot").ChangePivotCache _
        xlWkb.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataWks.Range("A1:G" & datarst.RecordCount + 1))
    
    pivotWks.PivotTables("RevenuesPivot").RefreshTable
    datarst.Close
    
    xlWkb.SaveAs Application.CurrentProject.Path & "\Exports\PivotAnalytics_" & Format(Date, "MMDDYYYY") & ".xlsx"
    xlObj.Visible = True
        
    ' REFESH ALL PIVOT TABLES
    For Each xlSheet In xlWkb.WorkSheets
        For Each xlPivot In xlSheet.PivotTables
            xlPivot.RefreshTable
        Next
    Next
       
    MsgBox "Dashboard workbook successfully created!", vbInformation
    
    Set datarst = Nothing
    Set db = Nothing
    
    Set xlObj = Nothing
    Set xlWkb = Nothing
    Set tblWks = Nothing
    Set dataWks = Nothing
    Set pivotWks = Nothing
    Exit Sub

ErrHandle:
    MsgBox Err.Number & " - " & Err.Description, vbCritical
    Set datarst = Nothing
    Set db = Nothing
    
    Set xlObj = Nothing
    Set xlWkb = Nothing
    Set tblWks = Nothing
    Set dataWks = Nothing
    Set pivotWks = Nothing
    Exit Sub
End Sub