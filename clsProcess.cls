VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Module:   clsProcess
'
' Project:  DataClassBuilder
'
' Purpose:  Connects to selected database and allows user
'           to create data class files for each selected
'           table from the database (using any/all selected
'           fields as class properties).  Creates a SQL
'           select (single record), update, insert and delete
'           function in each class.
'
'           NOTE: this class AND the created data classses
'           use clsDataAccess as an ADO data access wrapper.
'
' Author:   Michael J. Nugent
'
' Date:
'
'===============================================================
Option Explicit

Private Type FieldInfo
    FieldName As String
    FieldDataType As String
    FieldMaxLength As Long
    FieldVarName As String
    FieldIsPrimaryKey As Boolean
    FieldIsSelected As Boolean
End Type

Private m_typeFieldInfo() As FieldInfo

Private Type TableFieldInfoArray
    TableName As String
    TableFieldInfo() As FieldInfo
End Type

' collection of the m_typeFieldInfo() array
Private m_arrTableFieldData() As TableFieldInfoArray
' this will tell us which element of
' m_arrTableFieldData() is active
Private m_intActiveTableIndex As Integer

Private m_objTableList As clsComboList
Private m_objProcessedList As clsComboList
Private m_objFieldList As clsComboList
Private m_objClassList As clsComboList
Private m_objData As clsDataAccess
Private m_objError As clsError
Private m_objStatus As clsStatus

Private m_rsTable As ADODB.Recordset

' *********************************************************
' constants used in scripting to create class modules
' Property constants
Private Const cnP_LET = "Public Property Let "
Private Const cnP_SET = "Public Property Set "
Private Const cnP_GET = "Public Property Get "

Private Const cnP_ARG_BOOL = "(ByVal vData As Boolean)"
Private Const cnP_ARG_BYTE = "(ByVal vData As Byte)"
Private Const cnP_ARG_COL = "(ByVal vData As Collection)"
Private Const cnP_ARG_CUR = "(ByVal vData As Currency)"
Private Const cnP_ARG_DATE = "(ByVal vData As Date)"
Private Const cnP_ARG_DBL = "(ByVal vData As Double)"
Private Const cnP_ARG_INT = "(ByVal vData As Integer)"
Private Const cnP_ARG_LNG = "(ByVal vData As Long)"
Private Const cnP_ARG_OBJ = "(ByVal vData As Object)"
Private Const cnP_ARG_SINGLE = "(ByVal vData As Single)"
Private Const cnP_ARG_STR = "(ByVal vData As String)"
Private Const cnP_ARG_VAR = "(ByVal vData As Variant)"

Private Const cnP_VAR_BOOL = "m_bol"
Private Const cnP_VAR_BYTE = "m_byt"
Private Const cnP_VAR_COL = "m_col"
Private Const cnP_VAR_CUR = "m_cur"
Private Const cnP_VAR_DATE = "m_dat"
Private Const cnP_VAR_DBL = "m_dbl"
Private Const cnP_VAR_INT = "m_int"
Private Const cnP_VAR_LNG = "m_lng"
Private Const cnP_VAR_OBJ = "m_obj"
Private Const cnP_VAR_SINGLE = "m_s"
Private Const cnP_VAR_STR = "m_str"
Private Const cnP_VAR_VAR = "m_var"

' do we want Integer data types set as Longs in the class module
' some db Integers have greater values than the VB Integer data type.
Private m_bolConvertIntToLong As Boolean
' *********************************************************

Public Property Set ObjTableListbox(vData As clsComboList)
    ' listbox that contains of all tables in
    ' the selected database
    Set m_objTableList = vData
End Property

Public Property Set ObjProcessedList(vData As clsComboList)
    ' listbox that contains the tables selected for
    ' data class creation.
    Set m_objProcessedList = vData
End Property

Public Property Set ObjFieldList(vData As clsComboList)
    ' all the fields (properties) of a selected table
    Set m_objFieldList = vData
End Property

Public Property Set ObjClassList(vData As clsComboList)
    ' all the class modules created in this session
    Set m_objClassList = vData
End Property

Public Property Set ObjDataAccess(vData As clsDataAccess)
    Set m_objData = vData
End Property

Public Property Set ObjError(vData As clsError)
    Set m_objError = vData
End Property

Public Property Set ObjStatus(vData As clsStatus)
    ' lets us display (label on the active form)
    ' what process we are working
    Set m_objStatus = vData
End Property

Public Property Let ConvertIntegerToLong(vData As Boolean)
    m_bolConvertIntToLong = vData
End Property

Public Function MakeDataBaseConnection(strDatabaseType As String, _
    strServer As String, strUser As String, strPassword As String, _
        Optional strDatabase As String, Optional strSysDBPath As String, _
        Optional strSysDBPwd As String) As Boolean
    
    On Error GoTo err_MakeDataBaseConnection
    
    MakeDataBaseConnection = False
                
    Select Case strDatabaseType
    
            Case "Access"
                ' here strServer means the db and path
                MakeDataBaseConnection = m_objData.OpenConnectionAccess(strDatabase, _
                        strUser, strPassword, strSysDBPath, strSysDBPwd)
                        
            Case "Oracle"
                MakeDataBaseConnection = m_objData.OpenConnectionOracle(strServer, _
                        strUser, strPassword)
                        
            Case "SQL Server"
                MakeDataBaseConnection = m_objData.OpenConnectionSQLServer(strServer, strDatabase, _
                        strUser, strPassword)
                
            Case "Sybase"
                MakeDataBaseConnection = m_objData.OpenConnectionSybase(strServer, _
                        strUser, strPassword)
                                
            Case Else
                ' should never happen
                Err.Raise vbObjectError + 1099, , "Invalid database type selected"
        End Select
                        
    Exit Function
    
err_MakeDataBaseConnection:
    With Err
        .Raise .Number, .Source & vbCrLf & "MakeDatabaseConnection", .Description
    End With

End Function
       
Public Sub OpenTableSchema()
    
    On Error GoTo err_OpenTableSchema
    
    ' return back the db table schema into a recordset
    Set m_rsTable = m_objData.cnConnection.OpenSchema(adSchemaTables, _
            Array(Empty, Empty, Empty, "TABLE"))
    
    ' make sure a valid recordset was returned
    If m_objData.CheckValidRecordset(m_rsTable) Then
        ' fill the form listbox with tables
        m_objTableList.LoadListBoxADORecordset m_rsTable, 2, -1, False
    End If
    
    Exit Sub
    
err_OpenTableSchema:
    With Err
        .Raise .Number, .Source & vbCrLf & "OpenTableSchema", .Description
    End With
    
End Sub

Public Sub SetPrimaryField(lngIndex As Long)
    Dim lngCount As Long
    Dim i As Integer
    
    lngCount = UBound(m_arrTableFieldData(m_intActiveTableIndex).TableFieldInfo)
        
    If lngCount > 0 Then
        For i = 0 To lngCount - 1
            If i = CInt(lngIndex) Then
                m_arrTableFieldData(m_intActiveTableIndex).TableFieldInfo(lngIndex).FieldIsPrimaryKey = True
            Else
                m_arrTableFieldData(m_intActiveTableIndex).TableFieldInfo(i).FieldIsPrimaryKey = False
            End If
        Next
    End If
    
End Sub

Public Sub SetFieldForPropertyCreation(lngIndex As Long, bolSelected As Boolean)
    ' select or deselect field for class property creation
    m_arrTableFieldData(m_intActiveTableIndex).TableFieldInfo(lngIndex).FieldIsSelected = bolSelected
End Sub

Public Sub CreateTableClassProperties(lngTableCount As Long)
    Dim rsColumn As ADODB.Recordset
    Dim arrSelectedList() As String
    Dim strTableName As String
    Dim bolIsDuplicate As Boolean
    Dim bolNoPrimaryKey As Boolean
    Dim bolContinue As Boolean
    Dim lngRowCount As Long
    Dim intKeyId As Integer
    Dim i As Integer
    Dim j As Integer
    
    If lngTableCount = 0 Then Exit Sub
    
    On Error GoTo err_CreateTableClassProperties
    
    arrSelectedList = m_objTableList.ReturnListBoxSelectedText
    
    If UBound(arrSelectedList) > 0 Then bolContinue = True
    
    If lngTableCount > 0 Then
        ReDim Preserve m_arrTableFieldData(lngTableCount - 1) As TableFieldInfoArray
        
        For i = 0 To lngTableCount - 1
            ' default value
            bolIsDuplicate = False
            ' skip table if it's already in the listbox / only if any
            ' items in array (bolContinue = True)
            If bolContinue Then
                bolIsDuplicate = m_objProcessedList.FindListboxItem(arrSelectedList(i))
            End If
            If Not bolIsDuplicate Then
                ' get selected table schema information (column data)
                Set rsColumn = m_objData.cnConnection.OpenSchema(adSchemaColumns, _
                Array(Empty, Empty, arrSelectedList(i)))
                If m_objData.CheckValidRecordset(rsColumn) Then
                    With rsColumn
                        lngRowCount = .RecordCount
                        ReDim m_typeFieldInfo(lngRowCount) As FieldInfo
                        ' get table name
                        strTableName = .Fields("TABLE_NAME").Value
                        ' display process status
                        m_objStatus.PostStatus " Creating field properties for table " & strTableName & "..."
                        ' get field data
                        For j = 0 To lngRowCount - 1
                            m_typeFieldInfo(j).FieldName = Trim$(.Fields("COLUMN_NAME").Value)
                            m_typeFieldInfo(j).FieldDataType = GetFieldType(.Fields("DATA_TYPE").Value)
                            If Not IsNull(.Fields("CHARACTER_MAXIMUM_LENGTH").Value) Then
                                m_typeFieldInfo(j).FieldMaxLength = .Fields("CHARACTER_MAXIMUM_LENGTH").Value
                            Else
                                m_typeFieldInfo(j).FieldMaxLength = 0
                            End If
                            ' default is for all fields to be selected
                            m_typeFieldInfo(j).FieldIsSelected = True
                            ' if the field name contains spaces, add brackets to surround field name
                            If InStr(m_typeFieldInfo(j).FieldName, " ") > 0 Then
                                m_typeFieldInfo(j).FieldName = "[" & m_typeFieldInfo(j).FieldName & "]"
                            End If
                            .MoveNext
                        Next
                        ' add selected table to to listbox
                        m_objProcessedList.AddSingleItemToList strTableName, CLng(i), False
                    End With
                    m_arrTableFieldData(i).TableName = strTableName
                    m_arrTableFieldData(i).TableFieldInfo = m_typeFieldInfo
                    ' now we need to check for the primary key, if it does not exist,
                    ' we will "guess" and select the 1st long data type we find (or if none,
                    ' the 1st string data type we find, else the 1st field in the table).
                    ' The user can manually select another field as the primary key.
                    m_objStatus.PostStatus " Discovering primary key for table " & strTableName & "..."
                    Set rsColumn = m_objData.cnConnection.OpenSchema(adSchemaPrimaryKeys, _
                        Array(Empty, Empty, arrSelectedList(i)))
                    bolNoPrimaryKey = True
                    If m_objData.CheckValidRecordset(rsColumn) Then
                        With m_arrTableFieldData(i)
                            For j = 0 To lngRowCount - 1
                                If .TableFieldInfo(j).FieldName = rsColumn.Fields("COLUMN_NAME").Value Then
                                    .TableFieldInfo(j).FieldIsPrimaryKey = True
                                   ' intKeyId = j
                                    bolNoPrimaryKey = False
                                    Exit For
                                End If
                            Next j
                        End With
                    End If
                    
                    ' we have not found a primary key
                    If bolNoPrimaryKey Then
                        ' create our own primary key (return back the index of the
                        ' (arrSelectedList(i)).TableFieldInfo array) selected as
                        ' the primary key
                        m_objStatus.PostStatus " Primary key NOT found - application will automatically select key from field list..."
                        'intKeyId =
                        CreatePrimaryKey i, lngRowCount
                    End If
                End If ' m_objData.CheckValidRecordset(rsColumn)
            End If ' Not bolIsDuplicate
        Next ' i
    End If ' lngTableCount > 0
    Set rsColumn = Nothing
    
    Exit Sub
    
err_CreateTableClassProperties:
    With Err
        m_objError.UpdateLogFile "clsProcess", "CreateTableClassProperties", .Number, .Description, .Source
    End With
    
End Sub

Private Function CreatePrimaryKey(intIndex As Integer, lngCount As Long) As Integer
    ' select a primary key on the fly.  The 1st long we find, we will use as the key,
    ' if not a long, then an integer, if no integer, then string, if no string, 1st field
    ' we grab.
    ' ****** This is a RECURSIVE routine ******
    '
    ' Returns: index (j) from m_arrTableFieldData(intIndex)TableFieldInfo(j) to calling sub so we
    ' know the element index of the primary key for the TableFieldInfo array.
    Dim bolIsPrimaryKey As Boolean
    Dim arrDatatype(3) As String
    Dim i As Integer
    Static j As Integer
    
    On Error GoTo err_CreatePrimaryKey
    
    bolIsPrimaryKey = False
    
    arrDatatype(0) = "Long"
    arrDatatype(1) = "Integer"
    arrDatatype(2) = "String"
    
    With m_arrTableFieldData(intIndex)
        For i = 0 To lngCount
            ' if this is the 4th time through (0 to 2 = 3 times), grab the
            ' 1st field as the primary key and reset the static j variable
            If j <> 3 Then
                If .TableFieldInfo(i).FieldDataType = arrDatatype(j) Then
                    .TableFieldInfo(i).FieldIsPrimaryKey = True
                    bolIsPrimaryKey = True
                    Exit For
                Else
                    j = j + 1
                End If
            Else
                .TableFieldInfo(0).FieldIsPrimaryKey = True
                bolIsPrimaryKey = True
                Exit For
            End If
        Next i
    End With
    
    ' we created our primary key - return index value
    If bolIsPrimaryKey Then
        CreatePrimaryKey = j
    Else
        ' haven't created key yet, recurse back to function
        CreatePrimaryKey intIndex, lngCount
    End If
    
    Exit Function
    
err_CreatePrimaryKey:
    With Err
        m_objError.UpdateLogFile "clsProcess", "CreatePrimaryKey", .Number, .Description, .Source
    End With
    
End Function

Private Function GetFieldType(lngFieldType As Long) As String
    Dim strFieldType As String
      
    Select Case lngFieldType

        ' An 8-byte signed integer (DBTYPE_I8).
        Case adBigInt, adNumeric, 139 ' An exact numeric value with a fixed precision and scale (DBTYPE_NUMERIC).
            strFieldType = "Long"
        
        Case adBinary ' A binary value (DBTYPE_BYTES).
            strFieldType = "Byte"
        
        Case adBoolean ' A Boolean value (DBTYPE_BOOL).
            strFieldType = "Boolean"
            
        ' A String value (DBTYPE_STR).
        Case adChar, adVarChar, 130 ' A String value (Parameter object only).
            strFieldType = "String"
            
        Case adCurrency 'A currency value (DBTYPE_CY). Currency is a fixed-point number with four digits to the right of the decimal point. It is stored in an 8-byte signed integer scaled by 10,000.
            strFieldType = "Currency"
            
        ' A Date value (DBTYPE_DATE). A date is stored as a Double, the whole part of which is the number of days since December 30, 1899, and the fractional part of which is the fraction of a day.
        ' A date value (yyyymmdd) (DBTYPE_DBDATE).
        ' A time value (hhmmss) (DBTYPE_DBTIME).
        Case adDate, adDBDate, adDBTime, adDBTimeStamp ' A date-time stamp (yyyymmddhhmmss plus a fraction in billionths) (DBTYPE_DBTIMESTAMP).
            strFieldType = "Date"
                
        ' An exact numeric value with a fixed precision and scale (DBTYPE_DECIMAL).
        Case adDecimal, adDouble ' A double-precision floating point value (DBTYPE_R8).
            strFieldType = "Double"
        
        Case adInteger ' A 4-byte signed integer (DBTYPE_I4).
            strFieldType = "Integer"
            
        Case adSingle ' A single-precision floating point value (DBTYPE_R4).
            strFieldType = "Single"
        
        Case adVariant ' An Automation Variant (DBTYPE_VARIANT).
            strFieldType = "Object"
            
        Case Else
            strFieldType = "Unknown"
    End Select
        
    GetFieldType = strFieldType
    
End Function

Public Sub DisplayFieldProperties(strTable As String)
    ' display field properties for the table selected
    Dim lngCount As Long
    Dim i As Integer
    Dim j As Integer
    
    lngCount = UBound(m_arrTableFieldData)
    
    If lngCount >= 0 Then
        For i = 0 To lngCount
            With m_arrTableFieldData(i)
                If .TableName = strTable Then
                    m_intActiveTableIndex = i
                    lngCount = UBound(.TableFieldInfo)
                    
                    If lngCount > 0 Then
                        For j = 0 To lngCount - 1
                            m_objFieldList.AddSingleItemToList .TableFieldInfo(j).FieldName & _
                                Space(4) & .TableFieldInfo(j).FieldDataType & _
                                Space(8) & .TableFieldInfo(j).FieldMaxLength
                            ' fill the primary key textbox
                            If .TableFieldInfo(j).FieldIsPrimaryKey Then
                                frmClassBuilder.txtPrimaryKey = .TableFieldInfo(j).FieldName
                            End If
                            ' default select all field items
                            m_objFieldList.SelectAllListboxItems
                        Next
                        m_objFieldList.SetListboxListIndex -1
                    End If
                    Exit Sub
                End If
            End With
        Next i
    End If
    
End Sub

Public Sub DeleteTableFromArray()
    ' delete a table from the "table to process" list.  We do
    ' this by copying all the elements from the active m_arrTableFieldData
    ' array except the one tagged for deletion.  Then we reset
    ' the active m_arrTableFieldData to = the newly created copied
    ' array w/o the deleted field
    Dim arrTemp() As TableFieldInfoArray
    Dim strSelected() As String
    Dim lngSelected() As Long
    Dim lngRemoveCount As Long
    Dim lngKeepCount As Long
    Dim lngAllCount As Long
    Dim i As Integer
    Dim j As Integer
    
    On Error GoTo err_DeleteTableFromArray
    
    ' get the string values of the selected tables
    strSelected = m_objProcessedList.ReturnListBoxSelectedText
    
    ' if none selected - get out
    If Not CheckValidArray(strSelected) Then
        MsgBox "You must select at least 1 table to remove.", vbExclamation, "Remove Table"
        Exit Sub
    End If

    ' get the listindex values of the selected tables
    lngSelected = m_objProcessedList.ReturnLBSelectedItemsArray
    
    ' count of all elsements in TableFieldInfoArray
    lngAllCount = UBound(m_arrTableFieldData)
    ' count of all elements to be removed from TableFieldInfoArray
    lngRemoveCount = UBound(strSelected)
    
    ' if all items in array are selected for removal
    If lngAllCount - lngRemoveCount = 0 Then
        Erase m_arrTableFieldData
        m_objProcessedList.ClearListbox
        m_objFieldList.ClearListbox
        Exit Sub
    End If
    
    ' temporary array to hold any listbox items NOT selected for removal
    ReDim arrTemp(lngAllCount - lngRemoveCount - 1)
    
    ' delete the selected tables from the processed list
    For j = lngRemoveCount To 0 Step -1
         m_objProcessedList.DeleteListboxText CLng(lngSelected(j))
    Next
    
    ' now re-create m_arrTableFieldData array
    With m_objProcessedList
        ' select all remaining elements in list
        .SelectAllListboxItems
        ' get the string values of the selected tables
        strSelected = .ReturnListBoxSelectedText
    End With
    lngKeepCount = UBound(strSelected) ' this will be equal to lngRemoveCount
    
    ' clear listbox before refilling it
    m_objProcessedList.ClearListbox
    For i = 0 To lngKeepCount
        For j = 0 To lngAllCount
            If strSelected(i) = m_arrTableFieldData(j).TableName Then
                arrTemp(i) = m_arrTableFieldData(j)
                ' add selected table to to listbox
                m_objProcessedList.AddSingleItemToList m_arrTableFieldData(j).TableName, CLng(i), False
                Exit For
            End If
        Next ' j
    Next ' i
    
    Erase m_arrTableFieldData
    m_arrTableFieldData = arrTemp
    
    Exit Sub
    
err_DeleteTableFromArray:
    With Err
        .Raise .Number, .Source & vbCrLf & "DeleteTableFromArray", .Description
    End With

End Sub

Public Sub CreateClassModules(Optional strpath As String = "C:\")
    Dim intFileNumber As Integer
    Dim strClassHeader As String
    Dim strClassFile As String
    Dim strClassName As String
    Dim strTableName As String
    Dim strPrimaryKey As String
    Dim strPKeyDataType As String
    Dim strPKeyVarName As String
    Dim strFieldName As String
    Dim strVarName  As String
    Dim strRetType As String
    Dim strArgType As String
    Dim strSQLMethod As String
    Dim lngTableCount As Long
    Dim lngFieldCount As Long
    Dim i As Integer
    Dim j As Integer
    
    ' we are going to skip a possible "index out of bounds" error here
    On Error Resume Next
    
    lngTableCount = UBound(m_arrTableFieldData)
    
    Err.Clear
    
    On Error GoTo err_CreateClassModules
    
    If lngTableCount >= 0 Then
        For i = 0 To lngTableCount
            ' table name
            strTableName = m_arrTableFieldData(i).TableName
            ' check for spaces in table name
            If InStr(strTableName, " ") > 0 Then
                strTableName = Replace(strTableName, " ", "_")
            End If
            m_objStatus.PostStatus " Creating class module for table " & strTableName & "..."
            ' class name - w/o path and file extension
            strClassName = "cls" & strTableName
            ' class file name
            strClassFile = AddBackSlash(strpath) & strClassName & ".cls"
            intFileNumber = FreeFile
            ' see if file exists - if it does we
            ' give the user the chance to overwrite the file
            If Not CheckForExistingFile(strClassFile) Then
                ' create the class module header (not seen in the VB environment)
                strClassHeader = CreateInternalClassHeader("cls" & m_arrTableFieldData(i).TableName)
                Open strClassFile For Append As #intFileNumber
                
                Print #intFileNumber, strClassHeader
                Print #intFileNumber, "Option Explicit"
                Print #intFileNumber, ""
                
                Print #intFileNumber, "' *** NOTE *** "
                Print #intFileNumber, "' this is our tie-in to the clsDataAccess"
                Print #intFileNumber, "' class (for data connection and execution)."
                Print #intFileNumber, "Private m_objData As clsDataAccess"
                Print #intFileNumber, ""
                
                lngFieldCount = UBound(m_arrTableFieldData(i).TableFieldInfo)
                ' set field count to fieldcount - 1 because we start at field 0
                lngFieldCount = lngFieldCount - 1
                If lngFieldCount >= 0 Then
                    m_objStatus.PostStatus " Creating General Declarations section for class " & strClassName & "..."
                    ' create the module variables declared in the
                    ' General Declarations area
                    For j = 0 To lngFieldCount
                        ' only create properties for the user selected fields
                        If m_arrTableFieldData(i).TableFieldInfo(j).FieldIsSelected Then
                            strFieldName = CheckFieldNameForSpaces(m_arrTableFieldData(i).TableFieldInfo(j).FieldName)
                            ' check to see if we have selected to convert Integers Into Longs
                            If m_bolConvertIntToLong Then
                                ' change datatype from Int to Long
                                If m_arrTableFieldData(i).TableFieldInfo(j).FieldDataType = "Integer" Then
                                    m_arrTableFieldData(i).TableFieldInfo(j).FieldDataType = "Long"
                                End If
                            End If

                            ' get data type into var
                            strRetType = m_arrTableFieldData(i).TableFieldInfo(j).FieldDataType
                            ' assume any unknown data type is a Variant
                            If Trim$(LCase(strRetType)) = "unknown" Then
                                strRetType = "Variant"
                            End If
                        
                            Print #intFileNumber, "Private " & GetVariablePrefix(strRetType) & _
                                strFieldName & " As " & strRetType
                        End If
                    Next ' i
                    Print #intFileNumber, ""
                    
                    ' create class Set/Get property for clsDataAccess Object
                    Print #intFileNumber, "Public Property Set DataObject(vData As clsDataAccess)"
                    Print #intFileNumber, Space(4) & "Set m_objData = vData"
                    Print #intFileNumber, "End Property" & vbCrLf
                    Print #intFileNumber, "Public Property Get DataObject() As clsDataAccess"
                    Print #intFileNumber, Space(4) & "' external check for valid connection"
                    Print #intFileNumber, Space(4) & "If m_objData Is Nothing Then"
                    Print #intFileNumber, Space(8) & "Set m_objdata = New clsDataAccess"
                    Print #intFileNumber, Space(4) & "End If"
                    Print #intFileNumber, Space(4) & "Set DataObject = m_objData"
                    Print #intFileNumber, "End Property" & vbCrLf
                    
                    
                    ' create the class properties (Let, Get, and Set)
                    For j = 0 To lngFieldCount
                        If m_arrTableFieldData(i).TableFieldInfo(j).FieldIsSelected Then
                            ' database data type
                            strRetType = m_arrTableFieldData(i).TableFieldInfo(j).FieldDataType
                            If Trim$(LCase(strRetType)) = "unknown" Then
                                strRetType = "Variant"
                            End If
                            ' field name
                            ' make sure we take care of field names that contain spaces
                            strFieldName = CheckFieldNameForSpaces(m_arrTableFieldData(i).TableFieldInfo(j).FieldName)
                            strVarName = GetVariablePrefix(strRetType) & strFieldName
                            m_objStatus.PostStatus " Creating Property Get for field " & strFieldName & "..."
                            ' here we want to get the VB data type
                            ' set the TableFieldInfo(j) variable
                            m_arrTableFieldData(i).TableFieldInfo(j).FieldVarName = strVarName
                            strArgType = GetArgumentType(strRetType)
                            ' get the primary key name - 1 per table
                            If m_arrTableFieldData(i).TableFieldInfo(j).FieldIsPrimaryKey Then
                                ' these primary key vars will be needed in the SQLSelect,
                                ' SQLUpdate, SQLInsert, and SQLDelete methods
                                ' primary key
                                strPrimaryKey = m_arrTableFieldData(i).TableFieldInfo(j).FieldName  '''''strFieldName
                                ' data type of Primary key
                                strPKeyDataType = strRetType
                                ' variable name of primary key
                                strPKeyVarName = strVarName
                            End If
                            ' print property get header
                            Print #intFileNumber, cnP_GET & strFieldName & "() As " & strRetType
                            ' print property get body
                            Print #intFileNumber, Space(4) & strFieldName & " = " & strVarName
                            ' print property end
                            Print #intFileNumber, "End Property"
                            Print #intFileNumber, ""
                            
                            m_objStatus.PostStatus " Creating Property Set for field " & strFieldName & "..."
                            ' set up property set
                            If Trim$(LCase(strRetType)) = "object" Or Trim$(LCase(strRetType)) = "collection" _
                                Or Trim$(LCase(strRetType)) = "variant" Then
                                Print #intFileNumber, cnP_SET & strFieldName & strArgType
                                ' print property set body
                                Print #intFileNumber, Space(4) & "Set " & strVarName & " = vData"
                                ' print property end
                                Print #intFileNumber, "End Property"
                                Print #intFileNumber, ""
                            End If
                            
                            m_objStatus.PostStatus " Creating Property Let for field " & strFieldName & "..."
                            ' set up property let
                            If Not Trim$(LCase(strRetType)) = "object" And Not Trim$(LCase(strRetType)) = "collection" Then
                                Print #intFileNumber, cnP_LET & strFieldName & strArgType
                                ' print property set body
                                Print #intFileNumber, Space(4) & strVarName & " = vData"
                                ' print property end
                                Print #intFileNumber, "End Property"
                                Print #intFileNumber, ""
                            End If
                        End If ' if bolIsSelected = True
                    Next ' j
                    'Close #intFileNumber
                End If ' lngFieldCount > 0
                ' now we need to create the Select, Insert, Update, and Deletion methods
                ' select method
                m_objStatus.PostStatus " Creating SQLSelect Method for class " & strClassName & "..."
                strSQLMethod = CreateSQLSelectMethod(strTableName, lngFieldCount, strPrimaryKey, strPKeyDataType, strPKeyVarName, i)
                Print #intFileNumber, ""
                Print #intFileNumber, strSQLMethod
                ' insert method
                m_objStatus.PostStatus " Creating SQLInsert Method for class " & strClassName & "..."
                strSQLMethod = CreateSQLInsertMethod(strTableName, lngFieldCount, strPrimaryKey, strPKeyDataType, strPKeyVarName, i)
                Print #intFileNumber, ""
                Print #intFileNumber, strSQLMethod
                ' update method
                m_objStatus.PostStatus " Creating SQLUpdate Method for class " & strClassName & "..."
                strSQLMethod = CreateSQLUpdateMethod(strTableName, lngFieldCount, strPrimaryKey, strPKeyDataType, strPKeyVarName, i)
                Print #intFileNumber, ""
                Print #intFileNumber, strSQLMethod
                ' delete method
                m_objStatus.PostStatus " Creating SQLDelete Method for class " & strClassName & "..."
                strSQLMethod = CreateSQLDeleteMethod(strTableName, lngFieldCount, strPrimaryKey, strPKeyDataType, strPKeyVarName, i)
                Print #intFileNumber, ""
                Print #intFileNumber, strSQLMethod
                ' fill the created class in the class listbox
                FillCreatedModuleList strClassFile
                Close #intFileNumber
            End If ' Not CheckForExistingFile(strClassFile)
        Next ' i
    End If ' lngTableCount > 0
            
    Erase m_arrTableFieldData
    
    Exit Sub
    
err_CreateClassModules:
    With Err
        m_objError.UpdateLogFile "clsProcess", "CreateClassModules", .Number, .Description, .Source
     End With
     ' let's keep going
     Resume Next
    
End Sub

Private Function CheckFieldNameForSpaces(ByVal strFieldName As String) As String
    ' checks to see if fieldname contains [] - which denotes there are
    ' spaces in the field name (ugh.....)
    If InStr(strFieldName, "[") > 0 Then
        If InStr(strFieldName, "]") > 0 Then
            ' handle the brackets we added
            strFieldName = Replace(strFieldName, "[", "")
            strFieldName = Replace(strFieldName, "]", "")
            ' replace spces in name with underscores
            strFieldName = Replace(strFieldName, " ", "_")
        End If
    End If
    CheckFieldNameForSpaces = strFieldName
End Function

Private Function FillCreatedModuleList(strClassFile As String)
    Dim bolIsDuplicate As Boolean

    bolIsDuplicate = False
    
    ' make sure no duplicates / after we check for listcount
    If m_objClassList.ListCount > 0 Then
        bolIsDuplicate = m_objClassList.FindListboxItem(strClassFile)
    End If
    If Not bolIsDuplicate Then
        m_objClassList.AddSingleItemToList strClassFile, , False
    End If
    
End Function

Private Function CreateInternalClassHeader(strClassName As String) As String
    Dim strTemp As String
    
    strTemp = "VERSION 1.0 CLASS"
    
    strTemp = strTemp & vbCrLf & _
        "BEGIN" & vbCrLf & _
            Space(4) & "MultiUse = -1  'True" & vbCrLf & _
            Space(4) & "Persistable = 0  'NotPersistable" & vbCrLf & _
            Space(4) & "DataBindingBehavior = 0  'vbNone" & vbCrLf & _
            Space(4) & "DataSourceBehavior = 0   'vbNone" & vbCrLf & _
            Space(4) & "MTSTransactionMode = 0   'NotAnMTSObject" & vbCrLf & _
        "End" & vbCrLf
    strTemp = strTemp & "Attribute VB_Name = " & Chr$(34) & strClassName & Chr$(34) & vbCrLf
    strTemp = strTemp & "Attribute VB_GlobalNameSpace = False" & vbCrLf & _
                        "Attribute VB_Creatable = True" & vbCrLf & _
                        "Attribute VB_PredeclaredId = False" & vbCrLf & _
                        "Attribute VB_Exposed = False" & vbCrLf & _
                        "Attribute VB_Ext_KEY = " & Chr$(34) & _
                                "SavedWithClassBuilder6" & Chr$(34) & "," & _
                                Chr$(34) & "No" & Chr$(34) & vbCrLf & _
                        "Attribute VB_Ext_KEY = " & Chr$(34) & "Top_Level" & _
                            Chr$(34) & "," & Chr$(34) & "Yes" & Chr$(34)
    
    CreateInternalClassHeader = strTemp

End Function

Private Function GetVariablePrefix(strRetType As String) As String
    Dim strPrefix As String

    Select Case strRetType
        Case "Long"
            strPrefix = cnP_VAR_LNG
    
        Case "Byte"
            strPrefix = cnP_VAR_BYTE
        
        Case "Boolean"
            strPrefix = cnP_VAR_BOOL
            
        Case "String"
            strPrefix = cnP_VAR_STR
            
        Case "Currency"
            strPrefix = cnP_VAR_CUR
            
        Case "Date"
            strPrefix = cnP_VAR_DATE
            
        Case "Double"
            strPrefix = cnP_VAR_DBL
    
        Case "Integer"
            strPrefix = cnP_VAR_INT
    
        Case "Single"
            strPrefix = cnP_VAR_SINGLE
    
        Case "Object"
            strPrefix = cnP_VAR_OBJ
            
        Case "Collection"
            strPrefix = cnP_VAR_COL
    
        Case Else
            ' anything else is presumed to be a variant
            strPrefix = cnP_VAR_VAR
    End Select
    
    GetVariablePrefix = strPrefix
        
End Function

Private Function GetArgumentType(strRetType As String) As String
    Dim strType As String

    Select Case strRetType
        Case "Long"
            strType = cnP_ARG_LNG
    
        Case "Byte"
            strType = cnP_ARG_BYTE
        
        Case "Boolean"
            strType = cnP_ARG_BOOL
            
        Case "String"
            strType = cnP_ARG_STR
            
        Case "Currency"
            strType = cnP_ARG_CUR
            
        Case "Date"
            strType = cnP_ARG_DATE
            
        Case "Double"
            strType = cnP_ARG_DBL
    
        Case "Integer"
            strType = cnP_ARG_INT
    
        Case "Single"
            strType = cnP_ARG_SINGLE
    
        Case "Object"
            strType = cnP_ARG_OBJ
            
        Case "Collection"
            strType = cnP_ARG_COL
    
        Case Else
            ' anything else is presumed to be a variant
            strType = cnP_ARG_VAR
    End Select
    
    GetArgumentType = strType
        
End Function

Private Function CreateSQLSelectMethod(strTableName As String, lngFieldCount As Long, _
    strPrimaryKey As String, strPKeyDataType As String, strPKeyVarName As String, _
    intIndex As Integer) As String
    Dim strSelect As String
    Dim strArgName As String
    Dim strHold As String
    Dim bolIsLong As Boolean
    Dim bolIsInteger As Boolean
    Dim i As Integer
    
    strSelect = "Public Function SQLSelect(Optional "
    
    ' check for Data type of primary key (we assume only string or long or integer)
    If UCase(strPKeyDataType) = "LONG" Then
        strHold = " As Long = 0"
        bolIsLong = True
    ElseIf UCase(strPKeyDataType) = "INTEGER" Then
        strHold = " As Integer = 0"
        bolIsInteger = True
    Else
        strHold = " As String = " & Chr$(34) & Chr$(34)
    End If
        
    ' method argument name is the same as the var name but without
    ' the m_ -> lngName instead of m_lngName
    strArgName = Mid$(strPKeyVarName, 3)
    strSelect = strSelect & strArgName & strHold & ") As Boolean" & vbCrLf
    
    strSelect = strSelect & Space(4) & "Dim rsData As ADODB.Recordset" & vbCrLf & _
                            Space(4) & "Dim strSQL As String" & vbCrLf & _
                            Space(4) & "On Error GoTo err_SQLSelect" & vbCrLf & vbCrLf & _
                            Space(4) & "SQLSelect = False" & vbCrLf & vbCrLf
    
    strSelect = strSelect & Space(4) & "' the user may enter the primary key as an argument for selection," & vbCrLf & _
        Space(4) & "' otherwise we presume the user set the primary key property (Let)" & vbCrLf
        
        
    ' set up whether user used primary key as argument or
    ' used the property let to set the internal key variable
    
    ' if the primary key is Long
    If bolIsLong Or bolIsInteger Then
        strHold = " <> 0 Then "
    Else
        strHold = " <> " & Chr$(34) & Chr$(34) & " Then "
    End If
    
    strSelect = strSelect & Space(4) & "If " & strArgName & strHold & vbCrLf & _
        Space(8) & strPKeyVarName & " = " & strArgName & vbCrLf & _
        Space(4) & "End If" & vbCrLf & vbCrLf
    
        
    ' create SQL Select statement
    strSelect = strSelect & Space(4) & "strSQL = " & Chr$(34) & "SELECT " & Chr$(34) & vbCrLf
    
    ' iterate through the fields to create file string (i.e. Select field1, field2, etc...)
    For i = 0 To lngFieldCount
        With m_arrTableFieldData(intIndex).TableFieldInfo(i)
            If .FieldIsSelected Then
                If i <> lngFieldCount Then
                    strSelect = strSelect & Space(4) & "strSQL = strSQL & " & Chr$(34) & _
                        .FieldName & ", " & Chr$(34) & vbCrLf
                    
                Else
                    ' last field cannot have a comma after it
                    strSelect = strSelect & Space(4) & "strSQL = strSQL & " & Chr$(34) & _
                        .FieldName & " " & Chr$(34) & vbCrLf
                End If
            End If
        End With
    Next
            
    strSelect = strSelect & Space(4) & "strSQL = strSQL & " & Chr$(34) & "FROM " & strTableName & " " & Chr$(34) & vbCrLf
    
    ' where clause changes dependent on a string or numeric argument
    If bolIsLong Or bolIsInteger Then
        strSelect = strSelect & Space(4) & "strSQL = strSQL & " & Chr$(34) & "WHERE " & _
            strPrimaryKey & " = " & Chr$(34) & " & " & strPKeyVarName & " & " & Chr$(34) & " " & Chr$(34) & vbCrLf
    Else
        strSelect = strSelect & Space(4) & "strSQL = strSQL & " & Chr$(34) & _
            "WHERE " & strPrimaryKey & " = '" & Chr$(34) & " & " & _
                strPKeyVarName & " & " & Chr$(34) & "' " & Chr$(34) & vbCrLf
    End If
    
    strSelect = strSelect & vbCrLf
    
    strSelect = strSelect & Space(4) & "' Create recordset object.  We create a read only " & vbCrLf
    strSelect = strSelect & Space(4) & "' version because we use DML (Insert, Update, & Delete) to change data" & vbCrLf
    strSelect = strSelect & Space(4) & "Set rsData = m_objData.SQLSelect_ReadOnly(strSQL) " & vbCrLf
            
    ' make sure we returned a valid recordset
    strSelect = strSelect & Space(4) & "If m_objData.CheckValidRecordset(rsData) Then" & vbCrLf
    strSelect = strSelect & Space(8) & "' set data propertie values from table field values" & vbCrLf
    strSelect = strSelect & Space(8) & "' *** NOTE: you may want to put NULL error checking here" & vbCrLf
    strSelect = strSelect & Space(8) & "With rsData" & vbCrLf
    ' iterate through the fields collection again to set the class property values
    ' to those in the coresponding table fields
    For i = 0 To lngFieldCount
        With m_arrTableFieldData(intIndex).TableFieldInfo(i)
            If .FieldIsSelected Then
                strSelect = strSelect & Space(8) & Space(4) & .FieldVarName & " = !" & .FieldName & vbCrLf
            End If
        End With
    Next
    
    strSelect = strSelect & Space(8) & "End With" & vbCrLf
    strSelect = strSelect & Space(4) & "End If" & vbCrLf & vbCrLf
    strSelect = strSelect & Space(4) & "SQLSelect = True" & vbCrLf & vbCrLf
    
    strSelect = strSelect & "exit_SQLSelect:" & vbCrLf
    strSelect = strSelect & Space(4) & "On Error Resume Next" & vbCrLf
    strSelect = strSelect & Space(4) & "Set rsData = Nothing" & vbCrLf
    strSelect = strSelect & Space(4) & "Exit Function" & vbCrLf & vbCrLf
    
    strSelect = strSelect & "err_SQLSelect:" & vbCrLf
    strSelect = strSelect & Space(4) & "With Err" & vbCrLf
    strSelect = strSelect & Space(8) & "' *** NOTE ***" & vbCrLf
    strSelect = strSelect & Space(8) & "' This is a hack that will allow us to skip" & vbCrLf
    strSelect = strSelect & Space(8) & "' the invalid use of Null error we get when" & vbCrLf
    strSelect = strSelect & Space(8) & "' we try to push a Null value into a class" & vbCrLf
    strSelect = strSelect & Space(8) & "' property - you are better off using a sub" & vbCrLf
    strSelect = strSelect & Space(8) & "' or function to validate the data coming" & vbCrLf
    strSelect = strSelect & Space(8) & "' in from the database table - I use the" & vbCrLf
    strSelect = strSelect & Space(8) & "' CheckNullValue Function in the modUtilities" & vbCrLf
    strSelect = strSelect & Space(8) & "' Bas file in the DataClassBuilder project" & vbCrLf
    strSelect = strSelect & Space(8) & "If .Number = 94 Then Resume Next" & vbCrLf
    strSelect = strSelect & Space(8) & ".Raise .Number, .Source & vbCrLf & " & _
            Chr$(34) & "SQLSelect" & Chr$(34) & ", .Description" & vbCrLf
    strSelect = strSelect & Space(4) & "End With" & vbCrLf
    strSelect = strSelect & Space(4) & "Resume exit_SQLSelect" & vbCrLf & vbCrLf
    
    strSelect = strSelect & "End Function" & vbCrLf
    
    CreateSQLSelectMethod = strSelect
    
End Function

Private Function CreateSQLInsertMethod(strTableName As String, lngFieldCount As Long, _
    strPrimaryKey As String, strPKeyDataType As String, strPKeyVarName As String, _
    intIndex As Integer) As String
    Dim strInsert As String
    Dim strArgName As String
    Dim strEnding As String
    Dim strCorrectEnd As String
    Dim strHold As String
    Dim bolIsLong As Boolean
    Dim bolIsInteger As Boolean
    Dim bolAddToSQL As Boolean
    Dim i As Integer
    
    strInsert = "Public Function SQLInsert(Optional "
    
    ' check for Data type of primary key (we assume only string or long or integer)
    If UCase(strPKeyDataType) = "LONG" Then
        strHold = " As Long = 0"
        bolIsLong = True
    ElseIf UCase(strPKeyDataType) = "INTEGER" Then
        strHold = " As Integer = 0"
        bolIsInteger = True
    Else
        strHold = " As String = " & Chr$(34) & Chr$(34)
    End If
        
    ' method argument name is the same as the var name but without
    ' the m_ -> lngName instead of m_lngName
    strArgName = Mid$(strPKeyVarName, 3)
    strInsert = strInsert & strArgName & strHold & ") As Boolean" & vbCrLf
    
    strInsert = strInsert & Space(4) & "Dim strSQL As String" & vbCrLf & _
                            Space(4) & "Dim rsData As ADODB.Recordset" & vbCrLf & _
                            Space(4) & "Dim bolArgSent As boolean" & vbCrLf & _
                            Space(4) & "On Error GoTo err_SQLInsert" & vbCrLf & vbCrLf & _
                            Space(4) & "SQLInsert = False" & vbCrLf & vbCrLf
    
    strInsert = strInsert & Space(4) & "' the user may enter the primary key as an argument for Insertion," & vbCrLf & _
        Space(4) & "' otherwise we presume the user set the primary key property (Let)" & vbCrLf & _
        Space(4) & "' (for an INSERT the user should rarely send the primary key value)" & vbCrLf
        
        
    ' set up whether user used primary key as argument or
    ' used the property let to set the internal key variable
    ' if the primary key is Long
    If bolIsLong Or bolIsInteger Then
        strHold = " <> 0 Then "
    Else
        strHold = " <> " & Chr$(34) & Chr$(34) & " Then "
    End If
    
    strInsert = strInsert & Space(4) & "If " & strArgName & strHold & vbCrLf & _
        Space(8) & strPKeyVarName & " = " & strArgName & vbCrLf & _
        Space(8) & "' we need to handle the INSERT stmt differently " & vbCrLf & _
        Space(8) & "' if the user sent the Primary Key as an argument" & vbCrLf & _
        Space(8) & "bolArgSent = True" & vbCrLf & _
        Space(4) & "Else" & vbCrLf & _
        Space(8) & "bolArgSent = False" & vbCrLf & _
        Space(4) & "End If" & vbCrLf & vbCrLf
        
    ' create 2 SQL Insert statements (using IF/Else) so we can take into account
    ' whether the client used the primary key as an argument to this function
    strInsert = strInsert & Space(4) & "' we need to build 2 SQL statments in case the primary" & vbCrLf
    strInsert = strInsert & Space(4) & "' key was used as an argument for this function call." & vbCrLf
    strInsert = strInsert & Space(4) & "' NOTE: it usually is not a good idea to create the key from the client." & vbCrLf
    strInsert = strInsert & Space(4) & "'       Delete the 2nd strSQL statement if you will have the database" & vbCrLf
    strInsert = strInsert & Space(4) & "'       create the primary key automatically (best practice)" & vbCrLf
    strInsert = strInsert & Space(4) & "If Not bolArgSent Then" & vbCrLf
    
    ' ************************************************************************
    ' *** Code to insert record with the database table creating the primary key
    ' ************************************************************************
    strInsert = strInsert & Space(8) & "strSQL = " & Chr$(34) & "INSERT INTO " & Chr$(34) & vbCrLf
    strInsert = strInsert & Space(8) & "strSQL = strSQL & " & Chr$(34) & strTableName & " (" & Chr$(34) & vbCrLf
       
    ' iterate through the fields to create file string (i.e. INSERT INTO table1 Set (field1, field2, etc...)
    For i = 0 To lngFieldCount
        With m_arrTableFieldData(intIndex).TableFieldInfo(i)
            If .FieldIsSelected Then
                If Not .FieldIsPrimaryKey Then
                    If i <> lngFieldCount Then
                        strInsert = strInsert & Space(8) & "strSQL = strSQL & " & Chr$(34) & _
                            .FieldName & ", " & Chr$(34) & vbCrLf
                    Else
                        strInsert = strInsert & Space(8) & "strSQL = strSQL & " & Chr$(34) & _
                            .FieldName & ") " & Chr$(34) & vbCrLf
                    End If
                End If
            End If
        End With
    Next
    
    ' make sure we have the last ) in place, would be missing if the last field
    ' is the Primary Key which is skipped above.
    ' fill var with last 5 characters that should be in strInsert
    ' which is the ), a space, a ", and 1 vbcrlf
    strCorrectEnd = ") " & Chr$(34) & vbCrLf
    ' get actual value of last 5 chars from strInsert
    strEnding = Trim(Right$(strInsert, 5))
    ' compare the 2
    If strEnding <> strCorrectEnd Then
    'If m_arrTableFieldData(intIndex).TableFieldInfo(i - 1).FieldIsPrimaryKey Then
        ' get rid of the last comma and add the closing parenthesis
        strInsert = Left$(strInsert, InStrRev(strInsert, ",") - 1) & strCorrectEnd
    End If
    
    strInsert = strInsert & Space(8) & "strSQL = strSQL & " & Chr$(34) & " VALUES (" & Chr$(34) & vbCrLf
    For i = 0 To lngFieldCount
        With m_arrTableFieldData(intIndex).TableFieldInfo(i)
            If .FieldIsSelected Then
                If Not .FieldIsPrimaryKey Then
                    If i <> lngFieldCount Then
                        ' add double quotes for strings
                        If .FieldDataType = "String" Then
                            strInsert = strInsert & Space(8) & "strSQL = strSQL & " & Chr$(34) & _
                                "'" & Chr$(34) & " & " & .FieldVarName & " & " & Chr$(34) & "', " & Chr$(34) & vbCrLf
                        Else
                            strInsert = strInsert & Space(8) & "strSQL = strSQL & " & .FieldVarName & _
                                " & " & Chr$(34) & ", " & Chr$(34) & vbCrLf
                        End If
                    Else
                        ' last field cannot have an ending comma
                        ' add double quotes for strings
                        If .FieldDataType = "String" Then
                            strInsert = strInsert & Space(8) & "strSQL = strSQL & " & Chr$(34) & _
                                "'" & Chr$(34) & " & " & .FieldVarName & " & " & Chr$(34) & "') " & _
                                    Chr$(34) & vbCrLf & vbCrLf
                        Else
                            strInsert = strInsert & Space(8) & "strSQL = strSQL & " & _
                                .FieldVarName & " & " & Chr$(34) & ") " & _
                                    Chr$(34) & vbCrLf & vbCrLf
                        End If
                    End If
                End If
            End If
        End With
    Next
    
    ' make sure we have the last ) in place
    ' fill var with last 5 characters that should be in strInsert
    ' which is the ), a space, a ", and 2 vbcrlf (which account for 4 characters)
    strCorrectEnd = ") " & Chr$(34) & vbCrLf & vbCrLf
    ' get actual value of last 7 chars from strInsert
    strEnding = Trim(Right$(strInsert, 7))
    ' compare the 2
    If strEnding <> strCorrectEnd Then
    'If m_arrTableFieldData(intIndex).TableFieldInfo(i - 1).FieldIsPrimaryKey Then
        ' get rid of the last comma and add the closing parenthesis
        strInsert = Left$(strInsert, InStrRev(strInsert, ",") - 1) & strCorrectEnd
    End If
    strInsert = strInsert & Space(4) & "Else" & vbCrLf
    
    ' ************************************************************************
    ' *** Code to insert record with a client supplied primary key
    ' ************************************************************************
    ' set up the the 2nd SQL statement to include the primary key in the INSERT
    strInsert = strInsert & Space(8) & "strSQL = " & Chr$(34) & "INSERT INTO " & Chr$(34) & vbCrLf
    strInsert = strInsert & Space(8) & "strSQL = strSQL & " & Chr$(34) & strTableName & " (" & Chr$(34) & vbCrLf
    ' iterate through the fields to create file string (i.e. INSERT INTO table1 Set (field1, field2, etc...)
    ' use all selected fields - client creates the primary key
    For i = 0 To lngFieldCount
        With m_arrTableFieldData(intIndex).TableFieldInfo(i)
            If .FieldIsSelected Then
                If i <> lngFieldCount Then
                    strInsert = strInsert & Space(8) & "strSQL = strSQL & " & Chr$(34) & _
                        .FieldName & ", " & Chr$(34) & vbCrLf
                Else
                    strInsert = strInsert & Space(8) & "strSQL = strSQL & " & Chr$(34) & _
                        .FieldName & ") " & Chr$(34) & vbCrLf
                End If
            End If
        End With
    Next
    
    ' make sure we have the last ) in place
    'If Trim(Right$(strInsert, 3)) <> ")" Then
'        If m_arrTableFieldData(intIndex).TableFieldInfo(i - 1).FieldIsPrimaryKey Then
            ' get rid of the last comma and add the closing parenthesis
     '       strInsert = Left$(strInsert, InStrRev(strInsert, ",") - 1) & ")" & Chr$(34) & vbCrLf
'        End If
    'End If
    
    strInsert = strInsert & Space(8) & "strSQL = strSQL & " & Chr$(34) & " VALUES (" & Chr$(34) & vbCrLf
    ' use all selected fields - client creates the primary key
    For i = 0 To lngFieldCount
        With m_arrTableFieldData(intIndex).TableFieldInfo(i)
            If .FieldIsSelected Then
                If i <> lngFieldCount Then
                    ' add double quotes for strings
                    If .FieldDataType = "String" Then
                        strInsert = strInsert & Space(8) & "strSQL = strSQL & " & Chr$(34) & _
                            "'" & Chr$(34) & " & " & .FieldVarName & " & " & Chr$(34) & "', " & Chr$(34) & vbCrLf
                    Else
                        strInsert = strInsert & Space(8) & "strSQL = strSQL & " & .FieldVarName & _
                            " & " & Chr$(34) & ", " & Chr$(34) & vbCrLf
                    End If
                Else
                    ' last field cannot have an ending comma
                    ' add double quotes for strings
                    If .FieldDataType = "String" Then
                        strInsert = strInsert & Space(8) & "strSQL = strSQL & " & Chr$(34) & _
                            "'" & Chr$(34) & " & " & .FieldVarName & " & " & Chr$(34) & "') " & Chr$(34) & vbCrLf & vbCrLf
                    Else
                        strInsert = strInsert & Space(8) & "strSQL = strSQL & " & .FieldVarName & " & " & Chr$(34) & ") " & Chr$(34) & vbCrLf & vbCrLf
                    End If
                End If
            End If
        End With
    Next
    
    ' make sure we have the last ) in place
    'If Trim(Right$(strInsert, 1)) <> ")" Then
    'If m_arrTableFieldData(intIndex).TableFieldInfo(i - 1).FieldIsPrimaryKey Then
        ' get rid of the last comma and add the closing parenthesis
     '   strInsert = Left$(strInsert, InStrRev(strInsert, ",") - 1) & ")" & Chr$(34) & vbCrLf
    'End If
    
    strInsert = strInsert & Space(4) & "End If" & vbCrLf
    
    strInsert = strInsert & vbCrLf
    
    strInsert = strInsert & Space(4) & "SQLInsert = m_objData.SQLExecute(strSQL) " & vbCrLf & vbCrLf
    
    strInsert = strInsert & Space(4) & "' need to access the primary key if generated by the database table." & vbCrLf
    strInsert = strInsert & Space(4) & "' *** NOTE: Uncomment this code and add your own SQL statment to access the " & vbCrLf
    strInsert = strInsert & Space(4) & "'           primary key." & vbCrLf
    strInsert = strInsert & Space(4) & "'If bolArgSent Then"
    strInsert = strInsert & Space(4) & "'" & Space(4) & "strSQL = " & Chr$(34) & "SELECT..." & Chr$(34) & vbCrLf & ""
    strInsert = strInsert & Space(4) & "'" & Space(4) & "Set rsData = m_objConn.SQLExecute(strSQL)" & vbCrLf
    strInsert = strInsert & Space(4) & "'" & Space(4) & strPKeyVarName & " = rsData(0)" & vbCrLf
    strInsert = strInsert & Space(4) & "'End If" & vbCrLf & vbCrLf
    strInsert = strInsert & Space(4) & "' we can raise our own error if SQLInsert is False" & vbCrLf
    strInsert = strInsert & Space(4) & "' and the SQLExecute function does not raise it's " & vbCrLf
    strInsert = strInsert & Space(4) & "' own error back to this function (should not happen)" & vbCrLf
    
    strInsert = strInsert & Space(4) & "If Not SQLInsert Then" & vbCrLf
    strInsert = strInsert & Space(8) & "' *** NOTE: you may want to put a more meaningful error message here" & vbCrLf
    strInsert = strInsert & Space(8) & "Err.Raise vbObjectError + 2050,," & Chr$(34) & "Error inserting record." & Chr$(34) & vbCrLf
    strInsert = strInsert & Space(4) & "End If" & vbCrLf & vbCrLf
    
    strInsert = strInsert & Space(4) & "SQLInsert = True" & vbCrLf
    strInsert = strInsert & Space(4) & "Exit Function" & vbCrLf & vbCrLf
    
    strInsert = strInsert & "err_SQLInsert:" & vbCrLf
    strInsert = strInsert & Space(4) & "With Err" & vbCrLf
    strInsert = strInsert & Space(8) & ".Raise .Number, .Source & vbCrLf & " & _
            Chr$(34) & "SQLInsert" & Chr$(34) & ", .Description" & vbCrLf
    strInsert = strInsert & Space(4) & "End With" & vbCrLf & vbCrLf
    
    strInsert = strInsert & "End Function" & vbCrLf
    
    CreateSQLInsertMethod = strInsert
    
End Function

Private Function CreateSQLUpdateMethod(strTableName As String, lngFieldCount As Long, _
    strPrimaryKey As String, strPKeyDataType As String, strPKeyVarName As String, _
    intIndex As Integer) As String
    Dim strUpdate As String
    Dim strArgName As String
    Dim strRetType As String
    Dim strModVar As String
    Dim strHold As String
    Dim bolIsLong As Boolean
    Dim bolIsInteger As Boolean
    Dim i As Integer
    
    strUpdate = "Public Function SQLUpdate(Optional "
    
    ' check for Data type of primary key (we assume only string or long or integer)
    If UCase(strPKeyDataType) = "LONG" Then
        strHold = " As Long = 0"
        bolIsLong = True
    ElseIf UCase(strPKeyDataType) = "INTEGER" Then
        strHold = " As Integer = 0"
        bolIsInteger = True
    Else
        strHold = " As String = " & Chr$(34) & Chr$(34)
    End If
        
    ' method argument name is the same as the var name but without
    ' the m_ -> lngName instead of m_lngName
    strArgName = Mid$(strPKeyVarName, 3)
    strUpdate = strUpdate & strArgName & strHold & ") As Boolean" & vbCrLf
    
    strUpdate = strUpdate & Space(4) & "Dim strSQL As String" & vbCrLf & _
                            Space(4) & "On Error GoTo err_SQLUpdate" & vbCrLf & vbCrLf & _
                            Space(4) & "SQLUpdate = False" & vbCrLf & vbCrLf
    
    strUpdate = strUpdate & Space(4) & "' the user may enter the primary key as an argument for Update," & vbCrLf & _
        Space(4) & "' otherwise we presume the user set the primary key property (Let)" & vbCrLf
        
        
    ' set up whether user used primary key as argument or
    ' used the property let to set the internal key variable
    ' if the primary key is Long
    If bolIsLong Or bolIsInteger Then
        strHold = " <> 0 Then "
    Else
        strHold = " <> " & Chr$(34) & Chr$(34) & " Then "
    End If
    
    strUpdate = strUpdate & Space(4) & "If " & strArgName & strHold & vbCrLf & _
        Space(8) & strPKeyVarName & " = " & strArgName & vbCrLf & _
        Space(4) & "End If" & vbCrLf & vbCrLf
    
        
    ' create SQL Update statement
    strUpdate = strUpdate & Space(4) & "strSQL = " & Chr$(34) & "UPDATE " & Chr$(34) & vbCrLf
    strUpdate = strUpdate & Space(4) & "strSQL = strSQL & " & Chr$(34) & strTableName & " SET " & Chr$(34) & vbCrLf
    
    ' iterate through the fields to create file string (i.e. Update field1, field2, etc...)
    For i = 0 To lngFieldCount
        With m_arrTableFieldData(intIndex).TableFieldInfo(i)
            If .FieldIsSelected Then
                ' don't include the primary key in the update
                If Not .FieldIsPrimaryKey Then
                    strRetType = .FieldDataType
                    strModVar = GetVariablePrefix(strRetType) & .FieldName
                    If i <> lngFieldCount Then
                        If .FieldDataType = "String" Then
                            strUpdate = strUpdate & Space(4) & "strSQL = strSQL & " & Chr$(34) & _
                                .FieldName & " = '" & Chr$(34) & " & " & .FieldVarName & " & " & Chr$(34) & "', " & Chr$(34) & vbCrLf
                        Else
                            strUpdate = strUpdate & Space(4) & "strSQL = strSQL & " & Chr$(34) & _
                                .FieldName & " = " & Chr$(34) & " & " & .FieldVarName & " & " & Chr$(34) & ", " & Chr$(34) & vbCrLf
                        End If
                    Else
                        If .FieldDataType = "String" Then
                            ' we do not want to have a comma after the last field
                            strUpdate = strUpdate & Space(4) & "strSQL = strSQL & " & Chr$(34) & _
                                .FieldName & " = '" & Chr$(34) & " & " & .FieldVarName & " & " & Chr$(34) & "' " _
                                    & Chr$(34) & vbCrLf
                        Else
                            strUpdate = strUpdate & Space(4) & "strSQL = strSQL & " & Chr$(34) & _
                                .FieldName & " = " & Chr$(34) & " & " & .FieldVarName & " " & vbCrLf
                        End If
                    End If
                End If
            End If
        End With
    Next
    
    ' where clause changes dependent on a string or numeric argument
    If bolIsLong Or bolIsInteger Then
        strUpdate = strUpdate & Space(4) & "strSQL = strSQL & " & Chr$(34) & " WHERE " & _
            strPrimaryKey & " = " & Chr$(34) & " & " & strPKeyVarName & " & " & _
                Chr$(34) & " " & Chr$(34) & vbCrLf
    Else
        strUpdate = strUpdate & Space(4) & "strSQL = strSQL & " & Chr$(34) & _
            " WHERE " & strPrimaryKey & " = '" & Chr$(34) & " & " & _
                strPKeyVarName & " & " & Chr$(34) & "' " & Chr$(34) & vbCrLf
    End If
    
    strUpdate = strUpdate & vbCrLf
    
    strUpdate = strUpdate & Space(4) & "SQLUpdate = m_objData.SQLExecute(strSQL) " & vbCrLf & vbCrLf
    strUpdate = strUpdate & Space(4) & "' we can raise our own error if SQLUpdate is False" & vbCrLf
    strUpdate = strUpdate & Space(4) & "' and the SQLExecute function does not raise it's " & vbCrLf
    strUpdate = strUpdate & Space(4) & "' own error back to this function (should not happen)" & vbCrLf
    
    strUpdate = strUpdate & Space(4) & "If Not SQLUpdate Then" & vbCrLf
    strUpdate = strUpdate & Space(8) & "' *** NOTE: you may want to put a more meaningful error message here" & vbCrLf
    strUpdate = strUpdate & Space(8) & "Err.Raise vbObjectError + 2050,," & Chr$(34) & "Error updating record." & Chr$(34) & vbCrLf
    strUpdate = strUpdate & Space(4) & "End If" & vbCrLf & vbCrLf
    
    strUpdate = strUpdate & Space(4) & "SQLUpdate = True" & vbCrLf
    strUpdate = strUpdate & Space(4) & "Exit Function" & vbCrLf & vbCrLf
    
    strUpdate = strUpdate & "err_SQLUpdate:" & vbCrLf
    strUpdate = strUpdate & Space(4) & "With Err" & vbCrLf
    strUpdate = strUpdate & Space(8) & ".Raise .Number, .Source & vbCrLf & " & _
            Chr$(34) & "SQLUpdate" & Chr$(34) & ", .Description" & vbCrLf
    strUpdate = strUpdate & Space(4) & "End With" & vbCrLf & vbCrLf
        
    strUpdate = strUpdate & "End Function" & vbCrLf
    
    CreateSQLUpdateMethod = strUpdate
    
End Function

Private Function CreateSQLDeleteMethod(strTableName As String, lngFieldCount As Long, _
    strPrimaryKey As String, strPKeyDataType As String, strPKeyVarName As String, _
    intIndex As Integer) As String
    Dim strDelete As String
    Dim strArgName As String
    Dim strRetType As String
    Dim strModVar As String
    Dim strHold As String
    Dim bolIsLong As Boolean
    Dim bolIsInteger As Boolean
    Dim i As Integer
    
    strDelete = "Public Function SQLDelete(Optional "
    
    ' check for Data type of primary key (we assume only string or long or integer)
    If UCase(strPKeyDataType) = "LONG" Then
        strHold = " As Long = 0"
        bolIsLong = True
    ElseIf UCase(strPKeyDataType) = "INTEGER" Then
        strHold = " As Integer = 0"
        bolIsInteger = True
    Else
        strHold = " As String = " & Chr$(34) & Chr$(34)
    End If
        
    ' method argument name is the same as the var name but without
    ' the m_ -> lngName instead of m_lngName
    strArgName = Mid$(strPKeyVarName, 3)
    strDelete = strDelete & strArgName & strHold & ") As Boolean" & vbCrLf
    
    strDelete = strDelete & Space(4) & "Dim strSQL As String" & vbCrLf & vbCrLf & _
                            Space(4) & "On Error GoTo err_SQLDelete" & vbCrLf & vbCrLf & _
                            Space(4) & "SQLDelete = False" & vbCrLf & vbCrLf
    
    strDelete = strDelete & Space(4) & "' the user may enter the primary key as an argument for selection," & vbCrLf & _
        Space(4) & "' otherwise we presume the user set the primary key property (Let)" & vbCrLf
        
    ' set up whether user used primary key as argument or
    ' used the property let to set the internal key variable
    ' if the primary key is Long
    If bolIsLong Or bolIsInteger Then
        strHold = " <> 0 Then "
    Else
        strHold = " <> " & Chr$(34) & Chr$(34) & " Then "
    End If
    
    strDelete = strDelete & Space(4) & "If " & strArgName & strHold & vbCrLf & _
        Space(8) & strPKeyVarName & " = " & strArgName & vbCrLf & _
        Space(4) & "End If" & vbCrLf & vbCrLf
    
        
    ' create SQL Delete statement
    strDelete = strDelete & Space(4) & "strSQL = " & Chr$(34) & "DELETE FROM " & strTableName & Chr$(34) & vbCrLf
    
    ' where clause changes dependent on a string or numeric argument
    If bolIsLong Or bolIsInteger Then
        strDelete = strDelete & Space(4) & "strSQL = strSQL & " & Chr$(34) & " WHERE " & _
            strPrimaryKey & " = " & Chr$(34) & " & " & strPKeyVarName & " & " & Chr$(34) & " " & Chr$(34) & vbCrLf
    Else
        strDelete = strDelete & Space(4) & "strSQL = strSQL & " & Chr$(34) & _
            " WHERE " & strPrimaryKey & " = '" & Chr$(34) & " & " & _
                strPKeyVarName & " & " & Chr$(34) & "' " & Chr$(34) & vbCrLf
    End If
    
    strDelete = strDelete & vbCrLf
    
    strDelete = strDelete & Space(4) & "' delete record, raise our own error if the call to function " & vbCrLf
    strDelete = strDelete & Space(4) & "' SQLExecute returns a False w/o raising it's own error." & vbCrLf
    strDelete = strDelete & Space(4) & "SQLDelete = m_objData.SQLExecute(strSQL) " & vbCrLf & vbCrLf
    strDelete = strDelete & Space(4) & "' we can raise our own error if SQLDelete is False" & vbCrLf
    strDelete = strDelete & Space(4) & "' and the SQLExecute function does not raise it's " & vbCrLf
    strDelete = strDelete & Space(4) & "' own error back to this function (should not happen)" & vbCrLf
    
    strDelete = strDelete & Space(4) & "If Not SQLDelete Then" & vbCrLf
    strDelete = strDelete & Space(8) & "' *** NOTE: you may want to put a more meaningful error message here" & vbCrLf
    strDelete = strDelete & Space(8) & "Err.Raise vbObjectError + 2050,," & Chr$(34) & "Error deleting record." & Chr$(34) & vbCrLf
    strDelete = strDelete & Space(4) & "End If" & vbCrLf & vbCrLf
    
    strDelete = strDelete & "exit_SQLDelete:" & vbCrLf
    strDelete = strDelete & Space(4) & "SQLDelete = True" & vbCrLf
    strDelete = strDelete & Space(4) & "Exit Function" & vbCrLf & vbCrLf
    
    strDelete = strDelete & "err_SQLDelete:" & vbCrLf
    strDelete = strDelete & Space(4) & "With Err" & vbCrLf
    strDelete = strDelete & Space(8) & ".Raise .Number, .Source & vbCrLf & " & _
            Chr$(34) & "SQLDelete" & Chr$(34) & ", .Description" & vbCrLf
    strDelete = strDelete & Space(4) & "End With" & vbCrLf
    strDelete = strDelete & Space(4) & "Resume exit_SQLDelete" & vbCrLf & vbCrLf
    
    strDelete = strDelete & "End Function" & vbCrLf
    
    
    CreateSQLDeleteMethod = strDelete
    
End Function

Private Function CheckForExistingFile(strFile As String) As Boolean
    Dim objFile As clsFileOperations
    Dim lngRet As Long
    
    CheckForExistingFile = False
    
    Set objFile = New clsFileOperations
    
    With objFile
        If .FileFound(strFile) Then
            lngRet = MsgBox(strFile & " already exists.  Press OK to " & _
                            "overwrite the previous version or " & _
                            "Cancel to skip this module creation.", _
                            vbInformation + vbOKCancel, "Duplicate File")
            If lngRet = vbOK Then
                .FileDelete strFile
            Else
                CheckForExistingFile = True
            End If
        End If
    End With
    
    Set objFile = Nothing
    
End Function

Public Sub ViewClassModInNotepad(strModule As String)
    ' view newly created class module in Notepad
    On Error GoTo err_ViewClassModInNotepad
    
    ShellAndWait "notepad.exe " & strModule, vbNormalFocus
    
    Exit Sub
    
err_ViewClassModInNotepad:
    With Err
        m_objError.UpdateLogFile "clsProcess", "ViewClassInNotepad", .Number, .Description, .Source
     End With
     
End Sub

Public Sub ClearTypeArrays()
    On Error Resume Next
    
    Erase m_typeFieldInfo
    Erase m_arrTableFieldData
    
End Sub

Private Sub Class_Initialize()
    On Error Resume Next
    
    Set m_objData = New clsDataAccess
    
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    
    m_rsTable.Close
    Set m_rsTable = Nothing
    Set m_objTableList = Nothing
    Set m_objProcessedList = Nothing
    Set m_objFieldList = Nothing
    Set m_objClassList = Nothing
    Set m_objData = Nothing
    Set m_objError = Nothing
    Set m_objStatus = Nothing



End Sub
