Option Compare Database

' BUILDS CUSTOM RIBBON XML 
Public Function GetRibbonDefinition() As String
On Error GoTo ErrHandle
    Dim txtXML As String
    Dim nTabID As Integer
    Dim nGroupID As Integer

    nTabID = 0
    nGroupID = 0

    ' Add the XML header info
    txtXML = "<customUI xmlns=" & Chr(34) & _
    "http://schemas.microsoft.com/office/2009/07/customui" & Chr(34) & ">" & vbCrLf & _
    "  <ribbon startFromScratch=" & Chr(34) & "false" & Chr(34) & ">" & vbCrLf & _
    "    <tabs>" & vbCrLf

    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("mnuCommands")
    
    Do Until rs.EOF
        If (rs.Fields("GroupID") <> nGroupID And nGroupID <> 0) Then
            ' Close the previous group
            txtXML = txtXML & "        </group>" & vbCrLf
        End If
            
        If (rs.Fields("TabID") <> nTabID And nTabID <> 0) Then
            ' Close the previous tab
            txtXML = txtXML & "      </tab>" & vbCrLf
        End If
            
        If (rs.Fields("TabID") <> nTabID) Then
            ' Save the current tab ID
            nTabID = rs.Fields("TabID")
            
            ' Open the next tab
            txtXML = txtXML & "      <tab id=" & Chr(34) & "tab" & nTabID & _
                     Chr(34) & " label=" & Chr(34) & rs.Fields("TabLabel") & _
                     Chr(34) & ">" & vbCrLf
        End If
        
        If (rs.Fields("GroupID") <> nGroupID) Then
            ' Save the current group ID
            nGroupID = rs.Fields("GroupID")
            
            ' Open the next group
            txtXML = txtXML & "        <group id=" & Chr(34) & "group" & _
                     nGroupID & Chr(34) & " label=" & Chr(34) & _
                     rs.Fields("GroupLabel") & Chr(34) & ">" & vbCrLf
            If (Len(rs.Fields("Description")) > 1) Then
                txtXML = txtXML & "          <labelControl id=" & Chr(34) & _
                         "label" & nGroupID & Chr(34) & " label=" & Chr(34) & _
                         rs.Fields("Description") & Chr(34) & " />" & vbCrLf
            End If
            
        End If
        
        ' Add a command button
        txtXML = txtXML & _
           "          <button " & _
                             "id=" & Chr(34) & "button" & _
                                  rs.Fields("CommandID") & Chr(34) & vbCrLf & _
           "                  size=" & Chr(34) & rs.Fields("GraphicSize") & _
                                  Chr(34) & vbCrLf & _
           "                  label=" & Chr(34) & rs.Fields("Label") & _
                                  Chr(34) & vbCrLf & _
           "                  getImage=" & Chr(34) & "GetImage" & _
                                  Chr(34) & vbCrLf & _
           "                  screentip=" & Chr(34) & rs.Fields("ScreenTip") & _
                                  Chr(34) & vbCrLf & _
           "                  supertip=" & Chr(34) & rs.Fields("SuperTip") & _
                                  Chr(34) & vbCrLf & _
           "                  tag=" & Chr(34) & rs.Fields("Graphic") & _
                                  Chr(34) & vbCrLf & _
           "                  onAction=" & Chr(34) & "OnMenuAction" & _
                                  Chr(34) & " />" & vbCrLf
        
        rs.MoveNext
    Loop
    
    txtXML = txtXML & "        </group>" & vbCrLf
    txtXML = txtXML & "      </tab>" & vbCrLf
    txtXML = txtXML & "    </tabs>" & vbCrLf
    txtXML = txtXML & "  </ribbon>" & vbCrLf
    txtXML = txtXML & "</customUI>"
    
    GetRibbonDefinition = txtXML
   
    rs.Close
    Set rs = Nothing
    Exit Function
    
ErrHandle:
    MsgBox Err.Number & " - " & Err.Description, vbCritical, "Runtime Error"
    Set rs = Nothing
    Exit Function
End Function


' LOADS RIBBON, CALLED ON OPENING FORM
' REQUIREMENT: THIS EVENT MUST OCCUR IMMEDIATELY ON DATABSE OPEN
'              SINCE DB CANNOT ADD A RIBBON TAB AFTER LAUNCH
Public Function LoadRibbon() As Integer
On Error GoTo ErrHandle
    Application.LoadCustomUI "Library", GetRibbonDefinition
    Exit Function
    
ErrHandle:
    MsgBox Err.Number & " - " & Err.Description, vbCritical, "Runtime Error"
    Exit Function
End Function


' RETRIEVES MS OFFICE ICON IMAGE
Public Sub GetImage(ByVal control As Office.IRibbonControl, ByRef image)
On Error GoTo ErrHandle
    image = control.Tag
    Exit Sub
    
ErrHandle:
    MsgBox Err.Number & " - " & Err.Description, vbCritical, "Runtime Error"
    Exit Sub
End Sub


' ACTIVATES MENU BUTTONS
Public Sub OnMenuAction(ByVal control As Office.IRibbonControl)
On Error GoTo ErrHandle
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("mnuCommand", dbOpenDynaset)
    
    rs.FindFirst "[CommandID]=" & Mid(control.ID, 7)
    If rs.NoMatch Then
        MsgBox "You've clicked the button " & control.ID & " on the Ribbon"
    Else
        Select Case rs.Fields("TargetType")
            Case "Form"
                DoCmd.OpenForm rs.Fields("Target"), acNormal
            Case "Report"
                DoCmd.OpenReport rs.Fields("Target"), acViewPreview
            Case "Macro"
                DoCmd.RunMacro rs.Fields("Target")
            Case Else
                MsgBox "You've clicked the button " & control.ID & " on the Ribbon"
        End Select
    End If
    Exit Sub
    
ErrHandle:
    MsgBox Err.Number & " - " & Err.Description, vbCritical, "Runtime Error"
    Exit Sub
End Sub
