VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "db"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'****************************************************
' Foolish Games Data Access
' Database Connectivity Component
' Version 1.2, 01/17/2002
' Written by Lucas Holt
' Returns data to an application
' from a database as a string, array, or XML document.
' Changes:
' 1.01 4/2001  Added Error Checking
' 1.02 7/2001  Fixed bugs with error handling,
'   streamlined a few of the functions.  Also made
'   all of the functions parameters ByVal instead of
'   the default ByRef.  This should speed up calls
'   made out of process by the webserver even though
'   the string is longer than the optimal for ByVal.
' 1.03 7/2001  Changed GetRecords Function to
'   GetRecordsAsArray Function.  New function is
'   general and accepts any SQL String instead of
'   hard coded table.
' 1.1 8/2001  Changed GetRecordsToXML in the
'   following ways: simplified string concat on the
'   first portion of the xml document.  Specified
'   that arrContent is an array to the compiler.
' 1.11 8/2001  Added CleanXMLData private function
'   the function replaces characters that are known
'   to conflict with the XML output.
' 1.12 8/2001  Changed the settings to allow the dll
'   to stay in memory as opposed to reloading each
'   time.
' 1.13 8/2001 Changed XML function to use the Clean
'   xml data routine as well as removed unique code
'   to access databases and replaced it with
'   GetRecordset from this class.  This made up
'   for some of the speed loss.
' 1.2 1/2002 Cleaned up creation of object instances.
'    Now using Dim var As type then a set var = New type
' 1.2.1 1/2002 New error handling code routines in place
'    now the errors are logged in the application event
'    log in NT.
'****************************************************

Option Explicit

' Private storage for property value.
' ODBC Connect String Property
Private strConnectString As String

Public Property Get ConnectString() As String

   ConnectString = strConnectString

End Property

Public Property Let ConnectString(ByVal NewValue As String)
   
    If Len(NewValue) > 5 Then
        strConnectString = NewValue
    Else
        Err.Raise vbObjectError + 33, "fgaccess::ConnectString", "Invalid Connect String"
    End If
    
End Property

' **************************************************
' GetRecords Queries a database with a SQL Statement (CMD) and
' returns an array of records with the GetRows Method.
' ***************************************************

Public Function GetRecordsAsArray(ByVal strSQL _
  As String) As Variant
 
  Dim objRS As ADODB.Recordset ' Database Recordset
  Dim arrMyName As Variant
  
  On Error GoTo HandleError
  
  Set objRS = GetRecordset(strSQL)
  
  ' Data requested only if data exists in the Recordset
  If Not objRS.EOF Then
    ' Call to the GetRows Method
    arrMyName = objRS.GetRows
  End If
   
  ' Kill Recordset
  Set objRS = Nothing
    
  ' Return
  GetRecordsAsArray = arrMyName
  
  Exit Function
  
HandleError:

    App.LogEvent Err.Description & " in " & _
        "fgaccess::GetRecordsAsArray", _
        vbLogEventTypeError

    If Err.Number < 0 Then
        Err.Raise Err.Number, "fgaccess::GetRecordsAsArray", _
            Err.Description
    Else
        Err.Raise vbObjectError + Err.Number, _
            "fgaccess::GetRecordsAsArray", Err.Description
    End If

End Function

' ****************************************************
' The function below will call a database with a stored procedure
' passed and return an xml formated document (string)
'
' Current issue is that data must be altered when
' characters contain / < or >.
' ****************************************************

Public Function GetRecordsToXML(ByVal strStoredProc _
    As String) As String
  
  On Error GoTo HandleError
  
  Dim objRS As ADODB.Recordset
  Dim strResult As String
  Dim arrElements() As String
  Dim intCount As Integer ' Stores the number of elements (fields)
  Dim i As Integer ' Used as loop counter
  Dim j As Integer ' Used as loop counter
  Dim arrContent() As Variant  ' Stores recordset as array
  
  ' Retrive the Data
  Set objRS = GetRecordset(strStoredProc)
  
  intCount = objRS.Fields.Count - 1 ' Get Number of Fields and subtract one to match zero length array
  ReDim arrElements(intCount)
  
  For i = 0 To intCount
    If Not IsNull(objRS.Fields(i).Name) Then
      arrElements(i) = CleanXMLData(objRS.Fields(i).Name)
    End If
  Next
   
  arrContent = objRS.GetRows
  
  strResult = "<?xml version=""1.0""?>" & vbNewLine & _
    "<records>" & vbNewLine
  
  For i = 0 To UBound(arrContent, 2)
    strResult = strResult & "<record>" & vbNewLine
    For j = 0 To intCount
      strResult = strResult & "<" & arrElements(j) & ">"
      If Not IsNull(arrContent(j, i)) Then
        strResult = strResult & CleanXMLData(arrContent(j, i))
      End If
      strResult = strResult & "</" & arrElements(j) & ">" & vbNewLine
    Next
    strResult = strResult & "</record>" & vbNewLine
  Next
  
  strResult = strResult & "</records>" & vbNewLine
   
  ' Disconnect and Kill Recordset
    If IsObject(objRS) Then
      objRS.Close
      Set objRS = Nothing
    End If
    
  GetRecordsToXML = strResult
  
  Exit Function
  
HandleError:

    App.LogEvent Err.Description & " in " & _
        "fgaccess::GetRecordsToXML", _
        vbLogEventTypeError

    If Err.Number < 0 Then
        Err.Raise Err.Number, "fgaccess::GetRecordsToXML", _
            Err.Description
    Else
        Err.Raise vbObjectError + Err.Number, _
            "fgaccess::GetRecordsToXML", Err.Description
    End If
    
End Function

'*************************************************
' GetRecordset(sqlstmt) As Recordset
' Function returns a Disconnected
' recordset object to the calling application.
'*************************************************

Public Function GetRecordset(ByVal sqlstmt _
  As String) As ADODB.Recordset
  
  ' Added Error Handling
  On Error GoTo HandleError
  
    Dim objConn As ADODB.Connection
    Dim objRS As ADODB.Recordset

    Set objConn = New ADODB.Connection
    Set objRS = New ADODB.Recordset
    
  ' Open Connection with Property value
  objConn.Open strConnectString
  ' For Disconnected Recordset
  objRS.CursorLocation = adUseClient
  objRS.Open sqlstmt, objConn, adOpenForwardOnly, adLockReadOnly
  
  Set objRS.ActiveConnection = Nothing
  
  If IsObject(objConn) Then
      objConn.Close
      Set objConn = Nothing
  End If
  
  Set GetRecordset = objRS

    Exit Function

' Return any errors from the object.
HandleError:

    App.LogEvent Err.Description & " in " & _
        "fgaccess::GetRecordset", _
        vbLogEventTypeError

    If Err.Number < 0 Then
        Err.Raise Err.Number, "fgaccess::GetRecordset", _
            Err.Description
    Else
        Err.Raise vbObjectError + Err.Number, _
            "fgaccess::GetRecordset", Err.Description
    End If
    
End Function

'******************************************************
' CleanXMLData converts characters that cause problems
' with proper XSL translation and display of XML docs.
' The new characters are escaped similar to HTML code.
'
' NOTE: Do NOT pass Null to this as it will DIE.
' Passing ByVal is required because pointers to
' pointers are used.  You can not pass NULL to ByVal
' string.
'******************************************************

Private Function CleanXMLData(ByVal strInput As String) _
    As String

  On Error GoTo HandleError

  Dim strTemp As String  ' Used to hold modified input
  
  strTemp = Replace(strInput, "&", "&amp;", , , vbTextCompare)
  
  CleanXMLData = strTemp ' Return corrected string
  
  Exit Function
  
HandleError:

    App.LogEvent Err.Description & " in " & _
        "fgaccess::CleanXMLData", _
        vbLogEventTypeError

    If Err.Number < 0 Then
        Err.Raise Err.Number, "fgaccess::CleanXMLData", _
            Err.Description
    Else
        Err.Raise vbObjectError + Err.Number, _
            "fgaccess::CleanXMLData", Err.Description
    End If
    
End Function
