VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileOperations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'===============================================================
' Class:    clsFileoperations
'
' Project:
'
' Purpose:  This class is instantiated in any calling module.
'           File operations available are:
'                   - Copy file (Function FileCopy)
'                   - Move file (Function FileMove)
'                   - Delete file (Function FileDelete)
'                   - Rename file (Function FileRename)
'                   - Copy a directory (Function DirectoryCopy)
'                   Added additional functionality 9/2001:
'                   - Create a directory (Function CreateFolder)
'                   - Find file (Function FileFound - returns True/False)
'                   - Find files (Function FindFiles - returns string
'                       array of files found in a particular folder)
'
' Author:   Michael J. Nugent
'
' Date:     12/08/2000
'
'
'===============================================================
Option Explicit

' class name (for error reporting/raising)
Private Const CLASSNAME = "clsFileOperations"

' API file operations
Private Const MAX_PATH = 260

'File constants
Private Const FO_COPY = &H2&

Private Const FO_DELETE = &H3&
Private Const FO_MOVE = &H1&
Private Const FO_RENAME = &H4&
'Respond with Yes to All for any dialog box that is displayed.
Private Const FOF_NOCONFIRMATION = &H10&

'Does not display a progress dialog box.
Private Const FOF_SILENT = &H4&
' if set, instead of over-writing a file with the same name (for a copy),
' rename the file to "Copy of file (filename[copy number])
Private Const FOF_RENAMEONCOLLISION = &H8&

' display progress dialog box
Private Const FOF_SIMPLEPROGRESS = &H100

' file structure
Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As String
End Type

' api file operations call
Private Declare Function SHFileOperation Lib "shell32.dll" _
    Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
                      
' instance of the SHFILEOPSTRUCT private type
Private objFile As SHFILEOPSTRUCT

Private Type FILETIME ' 8 Bytes
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Private Type WIN32_FIND_DATA ' 318 Bytes
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved_ As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

' One or more of the following flags identifying
' the file's attributes (dwFileAttributes As Long):
' An archive file (which most files are).
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
' A file residing in a compressed drive or directory.
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
' A directory instead of a file.
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
' A hidden file, not normally visible to the user.
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
' An attribute-less file (cannot be combined with other attributes).
Private Const FILE_ATTRIBUTE_NORMAL = &H80
' A read-only file.
Private Const FILE_ATTRIBUTE_READONLY = &H1
' A system file, used exclusively by the operating system.
Private Const FILE_ATTRIBUTE_SYSTEM = &H4

Private Declare Function FindFirstFile Lib "kernel32" _
    Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData _
    As WIN32_FIND_DATA) As Long
    
Private Declare Function FindNextFile Lib "kernel32.dll" _
    Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData _
    As WIN32_FIND_DATA) As Long

Private Declare Function FindClose Lib "kernel32" _
    (ByVal hFindFile As Long) As Long
    
Private Declare Function CreateDirectory Lib "kernel32" Alias _
    "CreateDirectoryA" (ByVal lpPathname As String, lpSecurityAttributes _
    As SECURITY_ATTRIBUTES) As Long

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Variant
    bInheritHandle As Boolean
End Type

Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, _
                lpSystemTime As SYSTEMTIME) As Long
              
Public Function ReturnFileTime(strFile As String, Optional bolCreationTime As Boolean = False, _
    Optional bolLastWriteTime As Boolean = False, Optional bolLastAccessTime As Boolean = False) As Date
    Dim lpFindFileData As WIN32_FIND_DATA
    Dim sysTime As SYSTEMTIME
    Dim lngReturn As Long
    Dim lngRet As Long
    
    On Error GoTo err_ReturnFileTime

    lngReturn = FindFirstFile(strFile, lpFindFileData)
        
    If lngReturn = 0 Then
        Err.Raise vbObjectError + 1055, , "Error accessing file " & strFile & "."
    End If
       
    ' default to returning Creation time if none selected
    If Not bolCreationTime Then
        If Not bolLastAccessTime Then
            If Not bolLastWriteTime Then
                bolCreationTime = True
            End If
        End If
    End If
        
    If bolCreationTime Then
        lngReturn = FileTimeToSystemTime(lpFindFileData.ftCreationTime, sysTime)
    ElseIf bolLastAccessTime Then
        lngReturn = FileTimeToSystemTime(lpFindFileData.ftLastAccessTime, sysTime)
    ElseIf bolLastWriteTime Then
        lngReturn = FileTimeToSystemTime(lpFindFileData.ftLastWriteTime, sysTime)
    End If
    
    ' return file date and time stamp
    With sysTime
        ReturnFileTime = CDate(.wMonth & "/" & .wDay & "/" & .wYear & " " & _
            .wHour & ":" & .wMinute & ":" & .wSecond)
    End With
    ' close file
    FindClose lngReturn
    Exit Function

err_ReturnFileTime:
    With Err
         .Raise .Number, CLASSNAME & ": ReturnFileTime", .Description
    End With

End Function

Public Function FileMove(strSourceFile As String, strTargetFile As String, _
    Optional bolSilent As Boolean = True, Optional bolCopyOverDuplicateFile As Boolean = True) As Boolean
    
    Dim lngReturn As Long
    
    On Error GoTo err_FileMove
      
    FileMove = False
    
    ' move the desired file if it exists
    
    If FileFound(strSourceFile) Then
        ' fill file structure
        With objFile
            ' file to Copy
            .pFrom = strSourceFile
            ' where to copy file
            .pTo = strTargetFile
            If bolSilent Then
                ' no user interaction
                .fFlags = FOF_NOCONFIRMATION + FOF_SILENT
            Else
                .fFlags = FOF_SIMPLEPROGRESS
            End If
            
            If Not bolCopyOverDuplicateFile Then
                ' create copy of file with duplicate name
                .fFlags = .fFlags + FOF_RENAMEONCOLLISION
            End If
            ' move file flag
            .wFunc = FO_MOVE
            ' move the file
            lngReturn = SHFileOperation(objFile)
            ' if the move was NOT successful
            If lngReturn <> 0 Then
                Err.Raise vbObjectError + lngReturn, , "Unable to move file: " & strSourceFile
                Exit Function
            End If
        End With
    End If
    
    FileMove = True
    Exit Function
    
err_FileMove:
    With Err
         .Raise .Number, CLASSNAME & ": FileMove", .Description
    End With
    Exit Function
    
End Function

Public Function FileCopy(strSourceFile As String, strTargetFile As String, _
    Optional bolSilent As Boolean = True, Optional bolCopyOverDuplicateFile As Boolean = True) As Boolean
    
    Dim lngReturn As Long
    
    On Error GoTo err_FileCopy
      
    FileCopy = False
    
    ' copy the desired file if it exists
    If FileFound(strSourceFile) Then
        ' fill file structure
        With objFile
            ' file to Copy
            .pFrom = strSourceFile
            ' where to copy file
            .pTo = strTargetFile
            If bolSilent Then
                ' no user interaction
                .fFlags = FOF_NOCONFIRMATION + FOF_SILENT
            Else
                .fFlags = FOF_SIMPLEPROGRESS
            End If
            
            If Not bolCopyOverDuplicateFile Then
                ' create copy of file with duplicate name
                .fFlags = .fFlags + FOF_RENAMEONCOLLISION
            End If
            ' Copy file flag
            .wFunc = FO_COPY
            ' copy the file
            lngReturn = SHFileOperation(objFile)
            ' if the copy was NOT successful
            If lngReturn <> 0 Then
                Err.Raise vbObjectError + lngReturn, , "Unable to copy file: " & strSourceFile
                Exit Function
            End If
        End With
    End If
    
    FileCopy = True
    Exit Function
    
err_FileCopy:
    With Err
         .Raise .Number, CLASSNAME & ": FileCopy", .Description
    End With
    Exit Function
    
End Function

Public Function FileDelete(strDeleteFile As String, Optional bolSilent As Boolean = True) As Boolean
    
    Dim lngReturn As Long
    
    On Error GoTo err_FileDelete
          
    FileDelete = False
    
    ' copy the desired file if it exists
    If FileFound(strDeleteFile) Then
        ' fill file structure
        With objFile
            ' file to Delete
            .pFrom = strDeleteFile
            If bolSilent Then
                ' no user interaction
                .fFlags = FOF_NOCONFIRMATION + FOF_SILENT
            Else
                .fFlags = FOF_SIMPLEPROGRESS
            End If
            ' delete file flag
            .wFunc = FO_DELETE
            ' delete the file
            lngReturn = SHFileOperation(objFile)
            ' if the delete was NOT successful
            If lngReturn <> 0 Then
                Err.Raise vbObjectError + lngReturn, , "Unable to delete file: " & strDeleteFile
                Exit Function
            End If
        End With
    End If
    
    FileDelete = True
    Exit Function
    
err_FileDelete:
    With Err
         .Raise .Number, CLASSNAME & ": FileDelete", .Description
    End With
    Exit Function
    
End Function

Public Function FileRename(strSourceFile As String, strTargetFile As String, Optional bolSilent As Boolean = True)
    
    Dim lngReturn As Long
    
    On Error GoTo err_FileRename
    
    FileRename = False
    
    ' rename the desired file if it exists
    If FileFound(strSourceFile) Then
        ' fill file structure
        With objFile
            ' file to rename
            .pFrom = strSourceFile
            ' rename the file
            .pTo = strTargetFile
            If bolSilent Then
                ' no user interaction
                .fFlags = FOF_NOCONFIRMATION + FOF_SILENT
            Else
                .fFlags = FOF_SIMPLEPROGRESS
            End If
            ' delete file flag
            .wFunc = FO_RENAME
            ' copy the file
            lngReturn = SHFileOperation(objFile)
            ' if the rename was NOT successful
            If lngReturn <> 0 Then
                Err.Raise vbObjectError + lngReturn, , "Unable to rename file: " & strSourceFile
                Exit Function
            End If
        End With
    End If
    
    Exit Function
    
err_FileRename:
    With Err
         .Raise .Number, CLASSNAME & ": FileRename", .Description
    End With
    Exit Function
    
End Function

Public Function DirectoryCopy(strSourceDirectory As String, strTargetDirectory As String, Optional bolSilent As Boolean = True) As Boolean
    
    Dim lngReturn As Long
    
    On Error GoTo err_DirectoryCopy
    
    DirectoryCopy = False
    
    ' fill file structure
    With objFile
        ' file to Copy
        .pFrom = strSourceDirectory & Chr(0) & Chr(0)
        ' where to copy file
        .pTo = strTargetDirectory & Chr(0) & Chr(0)
        If bolSilent Then
            ' no user interaction
            .fFlags = FOF_NOCONFIRMATION + FOF_SILENT
        Else
            .fFlags = FOF_SIMPLEPROGRESS
        End If
        ' copy flag
        .wFunc = FO_COPY
        ' copy the directory
        lngReturn = SHFileOperation(objFile)
        ' if the copy was NOT successful
        If lngReturn <> 0 Then
            Err.Raise vbObjectError + lngReturn, , "Unable to copy directory: " & strSourceDirectory
            Exit Function
        End If
    End With

    DirectoryCopy = True
    Exit Function
    
err_DirectoryCopy:
    With Err
         .Raise .Number, CLASSNAME & ": DirectoryCopy", .Description
    End With
    Exit Function
    
End Function

Public Function FileFound(strFileName As String) As Boolean
    ' check if a file exists using the API
    '
    ' Parameter(s):
    '       strFileName (String) - string containing the file and path.
    '
    ' Returns:
    '       True (file is found) / False (file is not found)
    '
    Dim lpFindFileData As WIN32_FIND_DATA
    Dim lngReturn As Long
    
    lngReturn = FindFirstFile(strFileName, lpFindFileData)
      
    If lngReturn > 0 Then
        FindClose lngReturn
        FileFound = True
    Else
        FileFound = False
    End If
    
End Function

Public Function FindFiles(strFiles As String, Optional bolIncludeHiddenSystem _
    As Boolean = False) As String()
    ' Search for all files that match strFiles.
    ' Example: "C:\MyProgram\*.*".  Displays all the
    ' files in the MyPrograms folder.
    '
    ' Code written by:  Paul Kuliniewicz - Copyright 2000
    '                   http://www.vbapi.com/ref/f/findnextfile.html
    '
    ' Modified by:      Michael Nugent 9/21/2001
    '                   added code to check for sub directories
    '                   which we will ignore and also give the
    '                   user a choice to include system,
    '                   hidden files (bolIncludeHiddenSystem),
    '                   error handler, saving files and returning
    '                   as a string array
    '
    Dim hsearch As Long  ' handle to the file search
    Dim findinfo As WIN32_FIND_DATA  ' receives info about matching files
    Dim success As Long  ' will be 1 if successive searches are successful, 0 if not
    Dim retval As Long  ' generic return value
    ' initial file array (grows 50 elements at a time)
    ' contains all found files
    Dim arrFiles() As String
    ' if true, add found item to arrFiles array
    Dim bolAddToArray As Boolean
    ' use this value to redim temp array to only elements
    ' with a value
    Dim i As Integer
    ' array element counter (0-50)
    Dim j As Integer
    
    ' Begin a file search:
    hsearch = FindFirstFile(strFiles, findinfo)
    If hsearch = -1 Then  ' no files match the search string
        Exit Function
    End If
    
    ' Redim array w/50 elements
    ReDim arrFiles(50) As String
    
    ' Display name of each file that matches the search.  Note that the name is displayed, the
    ' next file (if any) is found, and then the loop restarts.  This way the first file
    ' (found above) will also be displayed.
    Do  ' begin loop
        ' Extract the filename from the fixed-length string:
        ' skip directories
        If findinfo.dwFileAttributes <> FILE_ATTRIBUTE_DIRECTORY Then
            ' these are all normal files or user wants all files
            If findinfo.dwFileAttributes = FILE_ATTRIBUTE_ARCHIVE _
                    Or findinfo.dwFileAttributes = FILE_ATTRIBUTE_NORMAL _
                    Or bolIncludeHiddenSystem = True Then
                bolAddToArray = True
            Else
                bolAddToArray = False
            End If
            
            If bolAddToArray Then
                arrFiles(i) = Left$(findinfo.cFileName, InStr(findinfo.cFileName, vbNullChar) - 1)
                ' count how many elements contain data
                i = i + 1
            End If
            
        End If
        ' Get the next matching file and loop if it exists:
        success = FindNextFile(hsearch, findinfo)
        
        ' array loop counter - redim array after every count to +50
        j = j + 1
        If j = 50 Then
            ReDim Preserve arrFiles(i + 50)
            j = 0
        End If
    Loop Until success = 0  ' keep looping until no more matching files are found
    
    ' if we have only 1 item, i will = 1 but element 0
    ' contains the item value
    If i > 0 Then
        ' redimension array to only those elements with values
        ReDim Preserve arrFiles(i - 1)
    End If
    
    FindFiles = arrFiles
    
exit_FindFiles:
    On Error Resume Next
    ' Close the file search handle
    retval = FindClose(hsearch)
    Exit Function
    
err_FindFiles:
    With Err
         .Raise .Number, CLASSNAME & ": FindFiles", .Description
    End With
    Resume exit_FindFiles
    
End Function

Public Function CreateFolder(strFolderPath As String) As Boolean
    Dim secAtt As SECURITY_ATTRIBUTES
    Dim lngReturn As Long
    
    CreateFolder = False
    
    lngReturn = CreateDirectory(strFolderPath, secAtt)
    
    If lngReturn <> 0 Then CreateFolder = True
    
    Exit Function

err_CreateDirectory:
    With Err
         .Raise .Number, CLASSNAME & ": CreateDirectory", .Description
    End With

End Function
