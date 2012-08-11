VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "fs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Accesses file system to read and write xml data"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'****************************************************
' Foolish Games Data Access
' File System I/0 Component
' Version 1.2, 01/19/2002
' Written by Lucas Holt
' Reads/Writes ASCII content from/to disk.
' Changes:
' 1.01 4/2001  Added Error Checking
' 1.2.2 1/2002 Changed the Error checking code to
' log to the windows nt event log (application).
'****************************************************

Public Function PutFileStr(StorageData As String, _
    FileName As String, FilePath As String)
  
  On Error GoTo HandleError
  
  Dim fso As FileSystemObject
  Dim ts As TextStream
  
  Set fso = New FileSystemObject
  
  Set ts = fso.CreateTextFile(FilePath + FileName, True, False)
  
  ts.Write (StorageData)
  ts.Close
  
  Set fso = Nothing
  Set ts = Nothing
  
  Exit Function
  
HandleError:

    App.LogEvent Err.Description & " in " & _
        "fgaccess::PutFileStr", _
        vbLogEventTypeError

    If Err.Number < 0 Then
        Err.Raise Err.Number, "fgaccess::PutFileStr", _
            Err.Description
    Else
        Err.Raise vbObjectError + Err.Number, _
            "fgaccess::PutFileStr", Err.Description
    End If
  
End Function

Public Function GetFileStr(ByVal path As String) As String
Attribute GetFileStr.VB_Description = "Reads ASCII File from disk and returns contents as a string.  Common uses include reading template files from disk."

  On Error GoTo HandleError

  Dim fso As FileSystemObject
  Dim ts As TextStream
  Dim strContent As String
  
  Set fso = New FileSystemObject
  
  'Open file.
  Set ts = fso.OpenTextFile(path)
  
  'Loop while not at the end of the file.
  Do While Not ts.AtEndOfStream
    strContent = strContent + ts.ReadLine + vbCrLf
  Loop
  
  'Close the file.
  ts.Close
  
  GetFileStr = strContent
  
  strContent = ""
  Set fso = Nothing
  Set ts = Nothing
  
  Exit Function
  
HandleError:

    App.LogEvent Err.Description & " in " & _
        "fgaccess::GetFileStr", _
        vbLogEventTypeError

    If Err.Number < 0 Then
        Err.Raise Err.Number, "fgaccess::GetFileStr", _
            Err.Description
    Else
        Err.Raise vbObjectError + Err.Number, _
            "fgaccess::GetFileStr", Err.Description
    End If
  
End Function
