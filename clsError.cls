VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'===============================================================
' Class:    clsError
'
' Project:
'
' Purpose:  Application error handler.  Provides user with message
'           box to relay any error that has occurred in addition to
'           generating (and appending to) a error log file located
'           in the same directory as the application.
'
' Author:   Michael J. Nugent
'
' Date:     2/11/2000
'
'===============================================================
Option Explicit

' name of application error logfile
Private m_strLogFileName As String
' do we want to display error to user
Private m_DisplayErrMsg As Boolean
' how many errors do we have (read only)
Private m_ErrCount As Long

Public Property Get ErrCount() As Long
    ErrCount = m_ErrCount
End Property

Private Property Let ErrCount(vData As Long)
    m_ErrCount = vData
End Property

Public Property Get DisplayErrMsg() As Boolean
    DisplayErrMsg = m_DisplayErrMsg
End Property

Public Property Let DisplayErrMsg(vData As Boolean)
    m_DisplayErrMsg = vData
End Property

Public Property Let LogFileName(strFileName As String)
    ' add the path to the logfile name
    m_strLogFileName = GetCorrectPath & strFileName
End Property

Public Property Get LogFileName() As String
    LogFileName = m_strLogFileName
End Property

Public Sub UpdateLogFile(ByVal strModName As String, ByVal strProcName As String, ByVal lngErrCode As Long, _
    ByVal strErrText As String, ByVal strErrSource As String)
    ' logs the passed-in error number and description.
    '
    ' Parameter(s):
    '       strModName (String)     -    Name of form or module where error occured.
    '       strProcName (String)    -    Name of procedure where error occured.
    '       lngErrCode (Long)       -    Err.Number.
    '       strErrText (String)     -    Err.Description.
    '       strErrSource (String)   -    Err.Source.
    '
    ' Returns:
    '       nothing.
    '
    Dim intLogFileNumber As Integer
    Dim strLogFile As String
    Dim strpath As String
  
    On Error Resume Next
    
    intLogFileNumber = FreeFile
    
    ' see if user named the error log file
    If Len(m_strLogFileName) = 0 Then
        ' find correct network path
        strpath = GetCorrectPath
        ' default name
        strLogFile = strpath & "ApplicationErrorLog.txt"
    Else
        strLogFile = m_strLogFileName
    End If
    
    Open strLogFile For Append As #intLogFileNumber
    
    Print #intLogFileNumber, "***** Error event record *****"
    Print #intLogFileNumber, "Event occured at: " & Now
    Print #intLogFileNumber, "User: " & Environ$("USERNAME")
    Print #intLogFileNumber, "Computer: " & Environ$("COMPUTERNAME")
    Print #intLogFileNumber, "Module Name : " & strModName
    Print #intLogFileNumber, "Procedure   : " & strProcName
    Print #intLogFileNumber, "Error number: " & CStr(lngErrCode)
    Print #intLogFileNumber, "Error descr : " & strErrText
    Print #intLogFileNumber, "Error Source: " & strErrSource
    Print #intLogFileNumber, String$(40, "=")
    Print #intLogFileNumber,
    
    Close intLogFileNumber
    
    ErrCount = m_ErrCount + 1
    
    ' some errors we will not display to user via the message box
    If m_DisplayErrMsg Then
        ' change mousepointer to default
        If Screen.MousePointer = vbHourglass Then Screen.MousePointer = vbDefault
        ' display error message to user
        Call MessageToUser(strModName, strProcName, lngErrCode, strErrText, strErrSource)
    End If
    
End Sub

Private Sub MessageToUser(ByVal strModName As String, ByVal strProcName As String, ByVal lngErrCode As Long, _
    ByVal strErrText As String, ByVal strErrSource As String)
    ' displays error message to user
    '
    ' Parameter(s):
    '       strModName (String)     -    Name of form or module where error occured.
    '       strProcName (String)    -    Name of procedure where error occured.
    '       lngErrCode (Long)       -    Err.Number.
    '       strErrText (String)     -    Err.Description.
    '       strErrSource (String)   -    Err.Source.
    '
    ' Returns:
    '       Nothing
    '
    Dim strMessage As String
    
    ' make sure screen mousepointer is set to default
    SetMousePointer
    
    ' build error message string
    strMessage = "Program Error" & vbCrLf
    strMessage = strMessage & "Module Name : " & strModName & vbCrLf
    strMessage = strMessage & "Procedure   : " & strProcName & vbCrLf
    strMessage = strMessage & "Error Source: " & strErrSource & vbCrLf
    strMessage = strMessage & "Error Number: " & CStr(lngErrCode) & vbCrLf
    strMessage = strMessage & "Error Desc  : " & strErrText & vbCrLf
    
    ' display error to user
    MsgBox strMessage, vbExclamation, App.Title
    
End Sub

Private Sub SetMousePointer()
    Dim frm As Form
    
    On Error Resume Next
    
    Screen.MousePointer = vbDefault
        
    For Each frm In Forms
        frm.MousePointer = vbDefault
    Next
    
End Sub
