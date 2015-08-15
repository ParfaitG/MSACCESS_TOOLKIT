Public Sub XMLTransformation()
    Dim xmlfile As Object, xslfile As Object, newxmlfile As Object
    Dim xmlstr As String, xslstr As String, newxmlstr As String
    
    ' LATE-BINDING OBJECT
    Set xmlfile = CreateObject("MSXML2.DOMDocument")
    Set xslfile = CreateObject("MSXML2.DOMDocument")
    Set newxmlfile = CreateObject("MSXML2.DOMDocument")
    
    ' SETTING DOCUMENT PATHS
    xmlstr = "C:\Path\To\Original_XML_File.xml"
    xslstr = "C:\Path\To\XSL_Stylesheet.xsl"
    newxmlstr = "C:\Path\To\New_XML_File.xml"

    ' LOADING XML & XSLT FILES
    xmlfile.async = False
    xmlfile.Load xmlstr
        
    xslfile.async = False
    xslfile.Load xslstr

    ' TRANSFORMING XML FILE USING XLST INTO NEW FILE
    xmlfile.transformNodeToObject xslfile, newxmlfile
    newxmlfile.Save newxmlstr
    
    ' UNINITIALIZING OBJECTS
    Set xmlfile = Nothing
    Set xslfile = Nothing
    Set newxmlfile = Nothing
    
    MsgBox "XML File successfully transformed!", vbInformation, "XML Transform Successful"
End Sub