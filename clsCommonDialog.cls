VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommonDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'===============================================================
' Module:   clsCommonDialog
'
' Project:
'
' Purpose:  Class module that provides common dialogbox
'           functionality using the Win API.
'
' Author:   Michael J. Nugent
'
' Date:
'
'===============================================================
Option Explicit

Private Const MAX_PATH = 260
Private Const OFN_EXPLORER = &H80000
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias _
    "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias _
    "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Filename As OPENFILENAME

Private Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type

' API declares, types, and constants used to access the
' directory browse dialog box

Private Const BIF_RETURNONLYFSDIRS = &H1

Private Type SHITEMID
  cb      As Long
  abID    As Byte
End Type

Private Type ITEMIDLIST
  mkid    As SHITEMID
End Type

Private Type BROWSEINFO
  hOwner          As Long
  pidlRoot        As Long
  pszDisplayName  As String
  lpszTitle       As String
  ulFlags         As Long
  lpfn            As Long
  lParam          As Long
  iImage          As Long
End Type

Private Declare Function SHBrowseForFolder Lib _
        "shell32.dll" Alias "SHBrowseForFolderA" _
        (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib _
    "shell32.dll" Alias "SHGetPathFromIDListA" _
    (ByVal pidl As Long, ByVal pszPath As String) As Long
    
Private Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" (ByVal hWnd As Long, _
    ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
    
Private Const SW_SHOWNORMAL = 1

Public Property Let DefaultExtension(Extention As String)
    ' Default extension for files
    Filename.lpstrDefExt = Extention
End Property

Public Property Get DefaultExtension() As String
    ' Default extension for files
    DefaultExtension = Filename.lpstrDefExt
End Property

Public Property Let ObjectOwner(Object As Form)
    ' Object that 'owns' the common dialog.
    Filename.hwndOwner = Object.hWnd
End Property

Public Property Let Filter(CustomFilter As String)
    ' File filter. Format : "Text for filter|filter|Next Text|Next filter". Ex : "All Files (*.*)|*.*|Text files (*.txt)|*.txt"
    Dim intCount As Integer
    
    With Filename
    
        .lpstrFilter = ""
    
        For intCount = 1 To Len(CustomFilter)
            If Mid$(CustomFilter, intCount, 1) = "|" Then
                .lpstrFilter = .lpstrFilter + Chr(0)
            Else
                .lpstrFilter = .lpstrFilter + Mid$(CustomFilter, intCount, 1)
            End If
        Next intCount
    
        .lpstrFilter = .lpstrFilter + Chr(0)
    End With
    
End Property

Public Property Let WindowTitle(Title As String)
    ' Window title
    Filename.lpstrTitle = Title
End Property
 Public Property Get WindowTitle() As String
    ' Window title
    WindowTitle = Filename.lpstrTitle
End Property

Public Property Let InitialDirectory(InitDir As String)
    ' Initial folder
    Filename.lpstrInitialDir = InitDir
End Property

Public Property Get InitialDirectory() As String
    ' Initial folder
    InitialDirectory = Filename.lpstrInitialDir
End Property

Public Function GetFileOpenName(Optional Multiselect As Boolean = False) As String
    ' Show common dialog open file control and returns file name. Use GetMultiFilename function to
    ' extract filenames when the multiselect flag is true. The Count function returns the number
    ' of files selected.
    
    On Error GoTo err_GetFileOpenName
    
    With Filename
        .hInstance = App.hInstance
        .lpstrFile = Chr(0) & Space(259)
        .nMaxFile = MAX_PATH
        .lStructSize = Len(Filename)
        GetOpenFileName Filename
        GetFileOpenName = Left$(.lpstrFile, InStr(.lpstrFile, Chr$(0)) - 1)
    End With
    
    Exit Function
    
err_GetFileOpenName:
    With Err
        .Raise .Number, .Source & vbCrLf & "GetFileOpenName", .Description
    End With
    
End Function

Public Function GetFileSaveName() As String
    ' Show common dialog save file control and returns file name.
    
    On Error GoTo err_GetFileSaveName
    
    With Filename
        .hInstance = App.hInstance
        .lpstrFile = Chr(0) & Space(259)
        .nMaxFile = MAX_PATH
        .flags = &H80000 Or &H4
        .lStructSize = Len(Filename)
        .flags = OFN_EXPLORER
        GetSaveFileName Filename
        ' return the saved file and path
        GetFileSaveName = .lpstrFile
    End With
    
    Exit Function
    
err_GetFileSaveName:
    With Err
        .Raise .Number, .Source & vbCrLf & "GetFileSaveName", .Description
    End With
    
End Function

Public Function BrowseDirectory() As String
    ' code taken from the EZVB -Home of Visual Basic Excellence
    ' web site (www.ezvb.com) - author unknown
    '
    ' Returns user selected folder
    '
    Dim typInfo As BROWSEINFO
    Dim typIDL As ITEMIDLIST
    Dim lRet As Long
    Dim lIDL As Long
    Dim sPath As String
    
    On Error GoTo err_BrowseDirectory
    
   ' typInfo.hOwner = Form1.hWnd
    typInfo.pidlRoot = 0&
    typInfo.lpszTitle = "Browse"
    typInfo.ulFlags = BIF_RETURNONLYFSDIRS
    
   'get the folder
    lIDL = SHBrowseForFolder(typInfo)
    
    sPath = Space$(512)
    lRet = SHGetPathFromIDList(ByVal lIDL, ByVal sPath)
      
    If lRet Then 'success
          sPath = Left$(sPath, InStr(sPath, Chr$(0)) - 1) & "\"
          BrowseDirectory = sPath
    End If
    
    Exit Function
    
err_BrowseDirectory:
    With Err
        .Raise .Number, .Source & vbCrLf & "BrowseDirectory", .Description
    End With
    
End Function


