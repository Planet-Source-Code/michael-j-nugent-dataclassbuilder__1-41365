VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'===============================================================
' Class:    clsDataAccess
'
' Project:
'
' Purpose:  Application data access handler.  Provides database
'           connection (Access, Oracle, Sybase, SQL Server) plus
'           generic data retrieval and modification.
'
' Author:   Michael J. Nugent
'
' Date:     2/11/2000
'
'===============================================================
Option Explicit

' internal connection object (stays open for the life of the
' class object - will try to reconnect if the connection
' somehow breaks while the class object is still active).
Private WithEvents m_connData As ADODB.Connection
Attribute m_connData.VB_VarHelpID = -1
' internal generic command object
Private m_cmdData As ADODB.Command
' trap ADO-specific errors
Private m_objError As ADODB.Error
' do we have a valid connection
Private m_bolIsConnected As Boolean

' initial server name, user id, and password
Private m_strServer As String
Private m_strUserid As String
Private m_strPassword As String
' database name for SQLServer and Access
Private m_strDatabase As String

' secrity db and password (for Access db)
Private m_strSecurityDB As String
Private m_strSecurityPwd As String

' type of server we are connected to
Private Enum ServerConnection
    Access = 0
    Sybase = 1
    Oracle = 2
    SQLServer = 3
End Enum

Private m_ServerType As ServerConnection

' connection failure error
Private Const ERROR_CONNECTION_FAILED = vbObjectError + 1050
Private Const ERROR_CONNECTION_FAILED_DESC = "Connection to Database failed."

' connection not set error
Private Const ERROR_NO_CONNECTION = vbObjectError + 1051
Private Const ERROR_NO_CONNECTION_DESC = "There is no connection set - OpenConnection method must be used first."

Public Property Get IsConnected() As Boolean
    ' external check for valid connection
    If m_connData Is Nothing Then
       IsConnected = False
    Else
       IsConnected = True
    End If
    
End Property

Public Property Get cnConnection() As ADODB.Connection
    ' retrieve valid connection for client application
    If Not m_connData Is Nothing Then
        If Not m_bolIsConnected Or m_connData.State = adStateClosed Then
            CheckValidConnection
        End If
        Set cnConnection = m_connData
    Else
        Err.Raise ERROR_NO_CONNECTION, Err.Source & _
            ": clsDataAccess -> Property Get cnConnection", ERROR_NO_CONNECTION_DESC
    End If
    
End Property

Public Sub BeginTransaction()
    If Not m_connData Is Nothing Then
        m_connData.BeginTrans
    End If
End Sub

Public Sub CommitTransaction()
    If Not m_connData Is Nothing Then
        m_connData.CommitTrans
    End If
End Sub

Public Sub RollbackTransaction()
    If Not m_connData Is Nothing Then
        m_connData.RollbackTrans
    End If
End Sub

Public Function OpenConnectionAccess(strDatabase As String, _
         Optional strUserID As String, _
         Optional strUserPwd As String, _
         Optional strSysDBPath As String, _
         Optional strSysDBPwd As String) As Boolean
    ' open connection to Access database
    '
    ' Returns:
    '       True  - successful connection made
    '       False - failed to connect/error occurred
    '
    Dim strConnect As String

    On Error GoTo err_OpenConnectionAccess
    
    OpenConnectionAccess = False
    Set m_connData = New ADODB.Connection
           
    ' save to module mem vars if needed to reconnect
    m_strDatabase = strDatabase
    m_strUserid = strUserID
    m_strPassword = strUserPwd
    m_strSecurityDB = strSysDBPath
    m_strSecurityPwd = strSysDBPwd
       
    If Len(strUserID) = 0 And Len(strUserPwd) = 0 Then
        ' make connection using DSN-Less OLEDB connection w/o user id and pwd
        strConnect = "Data Source=" & strDatabase & ";User ID=;Password=;"
    Else
        ' make connection using DSN-Less OLEDB connection w/user id and pwd
        strConnect = "DataSource=" & strDatabase & ";User ID=" & strUserID & _
            ";Password=" & strUserPwd & ";"
    End If
    
    ' connection properties
    With m_connData
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .ConnectionString = strConnect
        .ConnectionTimeout = 30
        .CursorLocation = adUseClient
        If Len(strSysDBPwd) > 0 Then
            ' if we are using a secure database (.mdw),
            ' this is the password
            .Properties("Jet OLEDB:Database Password") = strSysDBPwd
        End If
        If Len(strSysDBPath) > 0 Then
            ' if a .mdw security file is used
            .Properties("Jet OLEDB:System Database") = strSysDBPath
        End If
        .Open
    End With
    
    ' connection failed
    If Not m_connData Is Nothing Then
        ' set command object to this connection
        m_cmdData.ActiveConnection = m_connData
        m_ServerType = Access
        OpenConnectionAccess = True
    Else
        Err.Raise ERROR_CONNECTION_FAILED, "Function: OpenConnectionAccess", ERROR_CONNECTION_FAILED_DESC
    End If
    
    m_bolIsConnected = OpenConnectionAccess
    Exit Function
    
err_OpenConnectionAccess:
    With m_connData
        ' ADO-specific errors
        If .Errors.Count > 0 Then
            For Each m_objError In .Errors
                With m_objError
                    Err.Raise .Number, .Source & _
                        ": clsDataAccess -> OpenConnectionAccess", .Description
                End With
            Next
            .Errors.Clear
        Else
            ' other errors
            With Err
                .Raise .Number, .Source & _
                    ": clsDataAccess -> OpenConnectionAccess", .Description
            End With
        End If
    End With
    Exit Function
    
End Function

Public Function OpenConnectionOracle(strServer As String, _
    strUserID As String, strPassword As String) As Boolean
    ' open connection to oracle database
    '
    ' Returns:
    '       True  - successful connection made
    '       False - failed to connect/error occurred
    '
    Dim strConnect As String
    Dim strProcess As String

    On Error GoTo err_OpenConnection
    
    OpenConnectionOracle = False
    strProcess = "Set Function = False..."
    Set m_connData = New ADODB.Connection
               
    strProcess = strProcess & vbCrLf & "Created Connection Object..."
    ' save to module mem vars if needed to reconnect
    m_strServer = strServer
    m_strUserid = strUserID
    m_strPassword = strPassword
    strProcess = strProcess & vbCrLf & "Set internal vars (server, id, pw)..."
    
    ' make connection using DSN-Less OLEDB connection
    strConnect = "Provider=MSDAORA.1;Password=" & strPassword & _
        ";User ID=" & strUserID & ";Data Source=" & strServer & _
        ";Persist Security Info=True"
    
    strProcess = strProcess & vbCrLf & "Created the Connection String..."
    ' make connection using DSN-Less ODBC connection
    'strConnect = "Server=" & strServer & ";Driver={Microsoft ODBC for Oracle};" & _
        "UID= " & strUserId & "; Password= " & strPassword
        
    ' connection properties
    strProcess = strProcess & vbCrLf & "Creating Connection properties..."
    With m_connData
        .ConnectionString = strConnect
        .ConnectionTimeout = 30
        .CursorLocation = adUseClient
        .Open
    End With
    strProcess = strProcess & vbCrLf & "Connection Open..."
    
    ' connection failed
    If m_connData Is Nothing Then
        Err.Raise ERROR_CONNECTION_FAILED, "Function: OpenConnectionOracle", ERROR_CONNECTION_FAILED_DESC
    Else
        strProcess = strProcess & vbCrLf & "Set Command ActiveConnection property..."
        ' set command object to this connection
        m_cmdData.ActiveConnection = m_connData
        strProcess = strProcess & vbCrLf & "Command ActiveConnection property is set..."
        m_ServerType = Oracle
        strProcess = strProcess & vbCrLf & "Server Type is set..."
        OpenConnectionOracle = True
    End If
    strProcess = strProcess & vbCrLf & "Setting class variable m_bolIsConnected..."
    m_bolIsConnected = OpenConnectionOracle
    strProcess = strProcess & vbCrLf & "Exiting function..."
    Exit Function
    
err_OpenConnection:
    With m_connData
        ' ADO-specific errors
        If .Errors.Count > 0 Then
            For Each m_objError In .Errors
                With m_objError
                    Err.Raise .Number, .Source & _
                        ": clsDataAccess -> OpenConnectionOracle", .Description & vbCrLf & strProcess
                End With
            Next
            .Errors.Clear
        Else
            ' other errors
            With Err
                .Raise .Number, .Source & _
                    ": clsDataAccess -> OpenConnectionOracle", .Description & vbCrLf & strProcess
            End With
        End If
    End With
    Exit Function
    
End Function

Public Function OpenConnectionSQLServer(strServer As String, _
    strDatabase As String, strUserID As String, strPassword As String) As Boolean
    ' open connection to SQLServer database
    '
    ' Returns:
    '       True  - successful connection made
    '       False - failed to connect/error occurred
    '
    Dim strConnect As String

    On Error GoTo err_OpenConnection
    
    OpenConnectionSQLServer = False
    Set m_connData = New ADODB.Connection
           
    ' save to module mem vars if needed to reconnect
    m_strServer = strServer
    m_strDatabase = strDatabase
    m_strUserid = strUserID
    m_strPassword = strPassword
    
    ' make connection using DSN-Less OLEDB connection
    strConnect = "Driver={SQL Server};Server=" & strServer & _
                   ";Database=" & strDatabase & _
                   ";Uid=" & strUserID & _
                   ";Pwd=" & strPassword
        
    ' connection properties
    With m_connData
        .ConnectionString = strConnect
        .ConnectionTimeout = 30
        .CursorLocation = adUseClient
        .Open
    End With
    
    ' connection failed
    If m_connData Is Nothing Then
        Err.Raise ERROR_CONNECTION_FAILED, "Function: OpenConnectionSQLServer", ERROR_CONNECTION_FAILED_DESC
    Else
        ' set command object to this connection
        m_cmdData.ActiveConnection = m_connData
        m_ServerType = SQLServer
        OpenConnectionSQLServer = True
    End If
    
    m_bolIsConnected = OpenConnectionSQLServer
    Exit Function
    
err_OpenConnection:
    With m_connData
        ' ADO-specific errors
        If .Errors.Count > 0 Then
            For Each m_objError In .Errors
                With m_objError
                    Err.Raise .Number, .Source & _
                        ": clsDataAccess -> OpenConnectionSQLServer", .Description
                End With
            Next
            .Errors.Clear
        Else
            ' other errors
            With Err
                .Raise .Number, .Source & _
                    ": clsDataAccess -> OpenConnectionSQLServer", .Description
            End With
        End If
    End With
    Exit Function
    
End Function

Public Function OpenConnectionSybase(strServer As String, strUserID As String, strPassword As String) As Boolean
    ' open connection to Sybase 11 database
    '
    ' Returns:
    '       True  - successful connection made
    '       False - failed to connect/error occurred
    '
    Dim strConnect As String

    On Error GoTo err_OpenConnectionSybase
    
    OpenConnectionSybase = False
    Set m_connData = New ADODB.Connection
           
    ' save to module mem vars if needed to reconnect
    m_strServer = strServer
    m_strUserid = strUserID
    m_strPassword = strPassword
    
    ' make connection
    
    'Using Sybase ASE ODBC Driver
    strConnect = "Provider=MSDASQL;Driver={SYBASE ASE ODBC Driver};" & _
             "Srvr=" & strServer & _
             ";Uid=" & strUserID & _
             ";Pwd=" & strPassword & ";"
    
    ' connection properties
    With m_connData
        .ConnectionString = strConnect
        .ConnectionTimeout = 30
        .CursorLocation = adUseClient
        .Open
    End With
    
    ' connection failed?
    If Not m_connData Is Nothing Then
        ' set command object to this connection
        m_cmdData.ActiveConnection = m_connData
        m_ServerType = Sybase
        OpenConnectionSybase = True
    Else
        Err.Raise ERROR_CONNECTION_FAILED, "Function: OpenConnectionSybase", ERROR_CONNECTION_FAILED_DESC
    End If
    
    m_bolIsConnected = OpenConnectionSybase
    Exit Function
    
err_OpenConnectionSybase:
    With m_connData
        ' ADO-specific errors
        If .Errors.Count > 0 Then
            For Each m_objError In .Errors
                With m_objError
                    Err.Raise .Number, .Source & _
                        ": clsDataAccess -> OpenConnectionSybase", .Description
                End With
            Next
            .Errors.Clear
        Else
            ' other errors
            With Err
                .Raise .Number, .Source & _
                    ": clsDataAccess -> OpenConnectionSybase", .Description
            End With
        End If
    End With
    Exit Function
    
End Function

Public Function SQLSelectAllArray(strSQL As String) As Variant
    ' Returns recordset in the form of a variant array
    ' based upon strSql query argument
    '
    ' Returns:
    '       Variant array on success
    '       Raise error on failure
    '
    Dim rsData As ADODB.Recordset
    
    On Error GoTo err_SQLSelectArray
    
    ' check for valid connection
    CheckValidConnection
    
    With m_cmdData
        .ActiveConnection = m_connData
        .CommandText = strSQL
        Set rsData = .Execute
    End With
    
    If Not rsData Is Nothing Then
        With rsData
            If Not .BOF And Not .EOF Then
                ' return all records into variant array
                SQLSelectAllArray = .GetRows
            End If
        End With
    End If
    
exit_SQLSelectArray:
    On Error Resume Next
    rsData.Close
    Set rsData = Nothing
    Exit Function
    
err_SQLSelectArray:
    With m_cmdData.ActiveConnection
        ' ADO-specific errors
        If .Errors.Count > 0 Then
            For Each m_objError In .Errors
                With m_objError
                    Err.Raise .Number, .Source & _
                        ": clsDataAccess -> SQLSelectArray", .Description
                End With
            Next
            .Errors.Clear
        Else
            ' other errors
            With Err
                .Raise .Number, .Source & _
                    ": clsDataAccess -> SQLSelectArray", .Description
            End With
        End If
    End With
    Resume exit_SQLSelectArray
    
End Function

Public Function SQLSelectArrayTabString(strSQL As String) As String
    ' Returns recordset in the form of a tab delimited string array
    ' based upon strSql query argument
    '
    ' Returns:
    '       Variant array on success
    '       Raise error on failure
    '
    Dim rsData As ADODB.Recordset
    
    On Error GoTo err_SQLSelectArrayTabString
    
    ' check for valid connection
    CheckValidConnection
    
    With m_cmdData
        .ActiveConnection = m_connData
        .CommandText = strSQL
        Set rsData = .Execute
    End With
    
    If Not rsData Is Nothing Then
        With rsData
            If Not .BOF And Not .EOF Then
                ' return all records into tab delimited string array
                SQLSelectArrayTabString = .GetString(adClipString)
            End If
        End With
    End If
    
exit_SQLSelectArrayTabString:
    On Error Resume Next
    rsData.Close
    Set rsData = Nothing
    Exit Function
    
err_SQLSelectArrayTabString:
    With m_cmdData.ActiveConnection
        ' ADO-specific errors
        If .Errors.Count > 0 Then
            For Each m_objError In .Errors
                With m_objError
                    Err.Raise .Number, .Source & _
                        ": clsDataAccess -> SQLSelectArrayTabString", .Description
                End With
            Next
            .Errors.Clear
        Else
            ' other errors
            With Err
                .Raise .Number, .Source & _
                    ": clsDataAccess -> SQLSelectArraytabStrstring", .Description
            End With
        End If
    End With
    Resume exit_SQLSelectArrayTabString
    
End Function

Public Function SQLSelect_ReadOnly(strSQL As String) As ADODB.Recordset
    ' Returns read-only/non-updatable ADO recordset
    ' based upon strSql query argument
    '
    ' Returns:
    '       ADO recordset on success
    '       Raise error on failure
    '
    Dim rsData As ADODB.Recordset
    
    On Error GoTo err_SQLSelect_ReadOnly
    
    ' check for valid connection
    CheckValidConnection
    
    Set rsData = New ADODB.Recordset
    With rsData
        .ActiveConnection = m_connData
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .Open strSQL
    End With
    
    If Not rsData Is Nothing Then
        ' disconnect recordset from database
        rsData.ActiveConnection = Nothing
    End If
    
    ' set return recordset
    Set SQLSelect_ReadOnly = rsData
    
exit_SQLSelect_ReadOnly:
    On Error Resume Next
    Set rsData = Nothing
    Exit Function
    
err_SQLSelect_ReadOnly:
    With m_cmdData.ActiveConnection
        ' ADO-specific errors
        If .Errors.Count > 0 Then
            For Each m_objError In .Errors
                With m_objError
                    Err.Raise .Number, .Source & _
                        ": clsDataAccess -> SQLSelect_ReadOnly", .Description
                End With
            Next
            .Errors.Clear
        Else
            ' other errors
            With Err
                .Raise .Number, .Source & _
                    ": clsDataAccess -> SQLSelect_ReadOnly", .Description
            End With
        End If
    End With
    Resume exit_SQLSelect_ReadOnly
    
End Function

Public Function SQLSelect_Updateable(strSQL As String) As ADODB.Recordset
    ' Returns updatable ADO recordset
    ' based upon strSql query argument
    '
    ' Returns:
    '       ADO recordset on success
    '       Raise error on failure
    '
    Dim rsData As ADODB.Recordset
    
    On Error GoTo err_SQLSelect_Updateable
    
    ' check for valid connection
    CheckValidConnection
    
    Set rsData = New ADODB.Recordset
    With rsData
        .ActiveConnection = m_connData
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open strSQL
    End With
    
    If Not rsData Is Nothing Then
        ' disconnect recordset from database
        rsData.ActiveConnection = Nothing
    End If
    
    ' set return recordset
    Set SQLSelect_Updateable = rsData
    
exit_SQLSelect_Updateable:
    On Error Resume Next
    Set rsData = Nothing
    Exit Function
    
err_SQLSelect_Updateable:
    With m_cmdData.ActiveConnection
        ' ADO-specific errors
        If .Errors.Count > 0 Then
            For Each m_objError In .Errors
                With m_objError
                    Err.Raise .Number, .Source & _
                        ": clsDataAccess -> SQLSelect_Updateable", .Description
                End With
            Next
            .Errors.Clear
        Else
            ' other errors
            With Err
                .Raise .Number, .Source & _
                    ": clsDataAccess -> SQLSelect_Updateable", .Description
            End With
        End If
    End With
    Resume exit_SQLSelect_Updateable
    
End Function

Public Function SQLSelect_BatchUpdate(strSQL As String) As ADODB.Recordset
    ' Returns batch-updatable ADO recordset based upon strSql query argument
    '
    ' Returns:
    '       ADO recordset on success
    '       Raise error on failure
    '
    Dim rsData As ADODB.Recordset
    
    On Error GoTo err_SQLSelect_BatchUpdate
    
    ' check for valid connection
    CheckValidConnection
    
    Set rsData = New ADODB.Recordset
    With rsData
        .ActiveConnection = m_connData
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open strSQL
    End With
    
    If Not rsData Is Nothing Then
        ' disconnect recordset from database
        rsData.ActiveConnection = Nothing
    End If
    
    ' set return recordset
    Set SQLSelect_BatchUpdate = rsData
    
exit_SQLSelect_BatchUpdate:
    On Error Resume Next
    Set rsData = Nothing
    Exit Function
    
err_SQLSelect_BatchUpdate:
    With m_cmdData.ActiveConnection
        ' ADO-specific errors
        If .Errors.Count > 0 Then
            For Each m_objError In .Errors
                With m_objError
                    Err.Raise .Number, .Source & _
                        ": clsDataAccess -> SQLSelectArray", .Description
                End With
            Next
            .Errors.Clear
        Else
            ' other errors
            With Err
                .Raise .Number, .Source & _
                    ": clsDataAccess -> SQLSelectArray", .Description
            End With
        End If
    End With
    Resume exit_SQLSelect_BatchUpdate
    
End Function

Public Function ReconnectRecordsetForUpdate(rsData As ADODB.Recordset) As Boolean
    ' Reconnects a disconnected recordset to the active database
    ' for a batch update.
    '
    ' Returns:
    '       True  - successful batch update made
    '       False - failed to update/error occurred
    '
    On Error GoTo err_ReconnectRecordsetForUpdate
    
    ' check for valid connection
    CheckValidConnection
    
    m_cmdData.ActiveConnection.Errors.Clear
    
    ReconnectRecordsetForUpdate = False
    
    With rsData
        ' reset connection and update the table
        .ActiveConnection = m_connData
        .UpdateBatch
    End With
    
    ' successful batch update
    ReconnectRecordsetForUpdate = True
    Exit Function
    
err_ReconnectRecordsetForUpdate:
    ReconnectRecordsetForUpdate = False
    With m_cmdData.ActiveConnection
        ' ADO-specific errors
        If .Errors.Count > 0 Then
            For Each m_objError In .Errors
                With m_objError
                    Err.Raise .Number, .Source & _
                        ": clsDataAccess -> ReconnectRecordsetForUpdate", _
                            .Description
                End With
            Next
            .Errors.Clear
        Else
            ' other errors
            With Err
                .Raise .Number, .Source & _
                    ": clsDataAccess -> ReconnectRecordsetForUpdate", _
                        .Description
            End With
        End If
    End With
    Exit Function
    
End Function

Public Function SQLExecute(strSQL As String) As Boolean
    
    On Error GoTo err_SQLExecute

    SQLExecute = False
    
    With m_cmdData
        .CommandType = adCmdText
        .CommandText = strSQL
        .Execute strSQL, , adExecuteNoRecords
    End With

    SQLExecute = True
    
    Exit Function
    
err_SQLExecute:
    With m_cmdData.ActiveConnection
        ' ADO-specific errors
        If .Errors.Count > 0 Then
            For Each m_objError In .Errors
                With m_objError
                    Err.Raise .Number, .Source & _
                        ": clsDataAccess -> err_SQLExecute", .Description
                End With
            Next
            .Errors.Clear
        Else
            ' other errors
            With Err
                .Raise .Number, .Source & _
                    ": clsDataAccess -> SQLExecute", .Description
            End With
        End If
    End With
    
End Function

Public Function CheckValidRecordset(rsData As ADODB.Recordset) As Boolean
    
    CheckValidRecordset = False
    
    If Not rsData Is Nothing Then
        With rsData
            If Not .BOF And Not .EOF Then
                CheckValidRecordset = True
            End If
        End With
    End If
    
End Function

Private Sub CheckValidConnection()
    ' checks for a valid connection to the database.  Will try to reset
    ' connection if there is no connection.
    '
    ' Returns:
    '       Nothing - error will be raised in the OpenConnection function
    '       in the event of a connection problem.
    '
    
    ' need to check Is Nothing and State separately because if Is Nothing
    ' is True then State will error out
    If m_connData Is Nothing Then
        ' find out what type of connection we had,
        ' then reconnect
        ServerTypeConnection
        ' see if the connection is closed
    ElseIf m_connData.State = adStateClosed Then
        ' clear this instance of the connection object
        Set m_connData = Nothing
        ' find out what type of connection we had,
        ' then reconnect
        ServerTypeConnection
    End If
        
End Sub

Private Sub ServerTypeConnection()
    
    Select Case m_ServerType
        Case Access
            ReOpenAccessConnection
        
        Case Oracle
            ReOpenOracleConnection
        
        Case Sybase
            ReOpenSybaseConnection
            
        Case SQLServer
            ReOpenSQLServerConnection
        
        Case Else
            ' do nothing
    End Select
        
End Sub

Private Sub ReOpenOracleConnection()
    OpenConnectionOracle m_strServer, m_strUserid, m_strPassword
End Sub

Private Sub ReOpenSybaseConnection()
    OpenConnectionSybase m_strServer, m_strUserid, m_strPassword
End Sub

Private Sub ReOpenAccessConnection()
    OpenConnectionAccess m_strServer, m_strUserid, m_strPassword, m_strSecurityDB, m_strSecurityPwd
End Sub

Private Sub ReOpenSQLServerConnection()
    OpenConnectionSQLServer m_strServer, m_strDatabase, m_strUserid, m_strPassword
End Sub

Private Sub Class_Initialize()
    On Error Resume Next
    
    Set m_cmdData = New ADODB.Command
    
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    
    m_connData.Close
    
    Set m_cmdData = Nothing
    Set m_connData = Nothing
    Set m_objError = Nothing
    
End Sub
