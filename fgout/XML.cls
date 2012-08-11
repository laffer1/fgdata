VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'****************************************************
' Foolish Games Output
' XML translation component
' Version 1.0, 09/15/2001
' Written by Lucas Holt
' Takes XML input from an ADO data source and
' translates it to another format via XSLT.
' Changes:
' 1.01 9/2001  Added Error Checking
'****************************************************

Option Explicit

Private ofOutputType As OutputFormat    ' 1 or multi files?
Private strOutputPath As String         ' write to dir...
Private strFileName As String           ' suffix of file name
Private strXSLInputPath As String       ' xsl file used for trans.
Private strConnectString As String      ' ADO Data source
Private strSQLString As String          ' SQL command
Private blnEnableIdentity As Boolean    ' use an id field for file naming?
Private blnCancelProcessing As Boolean  ' used to cancel processing if something goes wrong.

' Used to build XML document in memory.
Private Const XMLStandardHeader = "<?xml version=""1.0""?>" & vbNewLine & vbNewLine
Private Const XMLStandardRootOpen = "<records>" & vbNewLine
Private Const XMLStandardRootClose = vbNewLine & "</records>" & vbNewLine

' Determines if Output is written to one file or multiple
' files.
Public Enum OutputFormat
  ofMultiFile = 0
  ofSingleFile = 1
End Enum

Private Sub Class_Initialize()
    ' Defaults
    ofOutputType = ofSingleFile    ' 1 or multi files?
    strOutputPath = "C:\temp\"       ' write to dir...
    blnEnableIdentity = False   ' use an id field for file naming?
    blnCancelProcessing = False  ' used to cancel processing if something goes wrong.
End Sub

Public Property Get OutputPath() As String
  OutputPath = strOutputPath
End Property

Public Property Let OutputPath(ByVal sNewValue As String)

    ' Add slash if it is not in place.
    If Not Right(sNewValue, 1) = "\" Then
        sNewValue = sNewValue & "\"
    End If
    
    'Set
    strOutputPath = sNewValue
End Property

Public Property Get XSLInputPath() As String
  XSLInputPath = strXSLInputPath
End Property

Public Property Let XSLInputPath(ByVal sNewValue As String)
  strXSLInputPath = sNewValue
End Property

Public Property Get IdentityNaming() As Boolean
    IdentityNaming = blnEnableIdentity
End Property

Public Property Let IdentityNaming(bNewValue As Boolean)
    blnEnableIdentity = bNewValue
End Property

Public Property Get SQLString() As String
  SQLString = strSQLString
End Property

Public Property Let SQLString(ByVal sNewValue As String)
  strSQLString = sNewValue
End Property

Public Property Let ConnectString(ByVal sNewValue As String)
  strConnectString = sNewValue
End Property

Public Property Get OutputType() As OutputFormat
  OutputType = ofOutputType
End Property

Public Property Let OutputType(ByVal ofMyType As OutputFormat)
    ofOutputType = ofMyType
End Property

Public Property Get FileName() As String
  FileName = strFileName
End Property

Public Property Let FileName(ByVal sNewValue As String)
  strFileName = sNewValue
End Property

Public Function WriteDocument()

    On Error GoTo HandleError
    
    Err.Clear
    
    Dim objFGdb As fgaccess.db
    Dim strXMLData As String
    Dim objFGfs As fgaccess.fs
    Dim strTemp As String
    Dim objXML As MSXML2.FreeThreadedDOMDocument30
    Dim blnXMLLoaded As Boolean
    
    Set objFGdb = New fgaccess.db
    Set objFGfs = New fgaccess.fs
    Set objXML = New MSXML2.FreeThreadedDOMDocument30
    
    With objXML
        .async = False
        .preserveWhiteSpace = False
        .validateOnParse = False
    End With
  
    ' Get XML data, i.e. read from the database
    objFGdb.ConnectString = strConnectString
    strXMLData = objFGdb.GetRecordsToXML(strSQLString)
  
    If Err.Number <> 0 Then
        Err.Raise (vbObjectError + 1), "FGOUT.XML.WriteDocument", "Error occured reading records from DB and converting them to XML."
    End If
    
    Select Case ofOutputType
        Case 0 ' Multifile Format
            ' Use XML object.. if true its loaded the xml.
            blnXMLLoaded = objXML.loadXML(strXMLData)
            
            If Not blnXMLLoaded Then
                Err.Raise (vbObjectError + 3), "FGOUT", "Error Loading XML Document"
            End If
            
            Dim rootEle As IXMLDOMNodeList
            Set rootEle = objXML.documentElement.childNodes
                
            If Err.Number <> 0 Then
                Err.Raise (vbObjectError + 4), "FGOUT.XML.WriteDocument", "Error Referencing Root child nodes."
            End If
                
            Dim intNumNodes As Integer
            Dim strXMLTemp As String
            Dim x As Integer
            Dim strFinalFileName As String
            
           ' content of <name>frank</name> i.e. just frank
           ' firstName = rootEle.Item(0).childNodes.Item(1).Text
            
            intNumNodes = CInt(rootEle.length)
            For x = 0 To (intNumNodes - 1)
                ' copy contents of Record to buffer
                ' keep format of XML (structure)
                ' Make buffer match that
                strXMLTemp = XMLStandardHeader & XMLStandardRootOpen & _
                    rootEle.Item(x).XML & XMLStandardRootClose
                ' Translate buffer
                strTemp = ConvertXML(strXMLTemp, strXSLInputPath)
                
                If blnEnableIdentity Then
                    strFinalFileName = rootEle.Item(x).childNodes.Item(0).Text & strFileName
                Else
                    strFinalFileName = CStr(x + 1) & strFileName
                End If
                
                ' Write buffer to disk
                Call objFGfs.PutFileStr(strTemp, strFinalFileName, strOutputPath)
            Next ' repeat process. (loop) until no more Record in Records
            
            Set rootEle = Nothing
        
        Case 1 ' Single File Format
            strTemp = ConvertXML(strXMLData, strXSLInputPath)
            Call objFGfs.PutFileStr(strTemp, strFileName, strOutputPath)
    End Select
    
    ' Clean Up
    Set objFGdb = Nothing
    Set objFGfs = Nothing
    Set objXML = Nothing
  
  Exit Function
  
HandleError:
  If Err.Number <> 0 Then
    Err.Raise Err.Number, Err.Source, Err.Description
  End If

End Function

Private Function ConvertXML(strXML, strFilePath) As String

    On Error GoTo HandleError

    Dim xmlDoc As MSXML2.FreeThreadedDOMDocument30   ' XML Document
    Dim xmlDoc2 As MSXML2.FreeThreadedDOMDocument30  ' XSL Document
    Dim strResult As String
    
    Set xmlDoc = New MSXML2.FreeThreadedDOMDocument30
    Set xmlDoc2 = New MSXML2.FreeThreadedDOMDocument30
    
    ' Set the Required Properties
    With xmlDoc
        .async = False
        .preserveWhiteSpace = False
        .validateOnParse = False
    End With
    
    ' Set the required properties for XSL document
    With xmlDoc2
        .async = False
        .preserveWhiteSpace = False
        .validateOnParse = False
    End With
    
    xmlDoc.loadXML (strXML)
    xmlDoc2.Load (strFilePath)
    
    ' Perform the translation
    strResult = xmlDoc.transformNode(xmlDoc2)
    
    ' Kill the XML objects
    Set xmlDoc = Nothing
    Set xmlDoc2 = Nothing
    
    ' return
    ConvertXML = strResult
    
    Exit Function
    
HandleError:
  If Err.Number <> 0 Then
    Err.Raise Err.Number, Err.Source, Err.Description
  End If

End Function
