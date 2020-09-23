VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComboList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'===============================================================
' Module:   clsComboList
'
' Project:
'
' Purpose:  Class module that provides combobox and listbox
'           functionality using the Win API.
'
' Author:   Michael J. Nugent
'
' Date:
'
'===============================================================
Option Explicit

Public Event SelectArray(lngItemArr() As Long)
Public Event IncrementProgressBar(intValue As Integer)

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
            (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
            lParam As Any) As Long
            
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Private Const LB_ADDSTRING = &H180
Private Const LB_SETCURSEL = &H186
Private Const LB_DIR = &H18D
Private Const LB_SETITEMDATA = &H19A
Private Const LB_RESETCONTENT = &H184
Private Const LB_GETTEXT = &H189
Private Const LB_GETTEXTLEN = &H18A
Private Const LB_GETSELCOUNT = &H190
Private Const LB_GETSELITEMS = &H191
Private Const LB_GETITEMDATA = &H199
Private Const LB_GETCURSEL = &H188
Private Const LB_GETCOUNT = &H18B
Private Const LB_SETSEL = &H185
Private Const LB_DELETESTRING = &H182
Private Const LB_SETTABSTOPS = &H192
Private Const LB_ERR = -1
Private Const LB_SETHORIZONTALEXTENT = &H194
Private Const LB_INSERTSTRING = &H181
Private Const LB_FINDSTRING = &H18F
Private Const LB_FINDSTRINGEXACT = &H1A2


' listbox object
Private m_objListbox As ListBox
' hWnd of the listbox control
Private m_lngListboxHwnd As Long
' number of items in the listbox
Private m_lngListCount As Long

Private Const CB_ADDSTRING = &H143
Private Const CB_SETCURSEL = &H14E
Private Const CB_DIR = &H145
Private Const CB_GETDROPPEDWIDTH = &H15F
Private Const CB_SETDROPPEDWIDTH = &H160
Private Const CB_SETITEMDATA = &H151
Private Const CB_RESETCONTENT = &H14B
Private Const CB_GETCOUNT = &H146
Private Const CB_GETITEMDATA = &H150
Private Const CB_FINDSTRING = &H14C
Private Const CB_FINDSTRINGEXACT = &H158
Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_GETCURSEL = &H147
Private Const CB_GETLBTEXT = &H148
Private Const CB_DELETESTRING = &H144
Private Const CB_INSERTSTRING = &H14A

' drive, directory, and file constants
Private Const DDL_READWRITE = &H0
Private Const DDL_READONLY = &H1
Private Const DDL_HIDDEN = &H2
Private Const DDL_SYSTEM = &H4
Private Const DDL_DIRECTORY = &H10
Private Const DDL_ARCHIVE = &H20
Private Const DDL_DRIVES = &H4000
Private Const DDL_EXCLUSIVE = &H8000&

' how many items in the combo
Private m_lngComboCount As Long
' combobox object
Private m_objCombobox As ComboBox
' hWnd of combobox control
Private m_lngComboboxHwnd As Long
' did we add a blank to the combo list - if so,
' this changes how we can return the listindex value
' in Function ReturnComboItemData
Private m_bolBlankAdded As Boolean

Public Property Set ListboxObject(objList As ListBox)
    Set m_objListbox = objList
    ' get hwnd property of listbox
    m_lngListboxHwnd = m_objListbox.hWnd
End Property

Public Property Set ComboboxObject(objList As ComboBox)
    Set m_objCombobox = objList
    ' get hwnd property of listbox
    m_lngComboboxHwnd = m_objCombobox.hWnd
End Property

Public Property Get ComboItemsCount() As Long
    ' make sure we have the correct count
    If m_lngComboCount > 0 Then
        ComboItemsCount = m_lngComboCount
    Else
        ComboItemsCount = SendMessage(m_lngComboboxHwnd, CB_GETCOUNT, ByVal 0&, 0&)
    End If
End Property

Public Property Get ListCount() As Long
    ' make sure we have the correct count
    If m_lngListCount > 0 Then
        ListCount = m_lngListCount
    Else
        ListCount = SendMessage(m_lngListboxHwnd, LB_GETCOUNT, ByVal 0&, 0&)
    End If
End Property

Private Property Let ListCount(NewValue As Long)
    m_lngListCount = NewValue
End Property

Public Sub ReturnListBoxSelectedItem()
    Dim lngSelectCount As Long
    Dim lngItems() As Long
    Dim lngRet As Long
    Dim i As Integer
    
    On Error GoTo err_ReturnListBoxSelectedItem
    
    lngSelectCount = SendMessage(m_lngListboxHwnd, LB_GETSELCOUNT, ByVal 0&, 0&)
    
    If lngSelectCount > 0 Then
        ReDim lngItems(lngSelectCount - 1) As Long
        ' Get the indexes of all selected items.
        lngRet = SendMessage(m_lngListboxHwnd, LB_GETSELITEMS, ByVal lngSelectCount, lngItems(0))
        
        ' fill array with ItemData values
        For i = 0 To lngSelectCount - 1
            lngRet = SendMessage(m_lngListboxHwnd, LB_GETITEMDATA, ByVal lngItems(i), ByVal 0&)
            lngItems(i) = lngRet
        Next i
        ' send array back to client
        RaiseEvent SelectArray(lngItems)
    End If
    
    Exit Sub
    
err_ReturnListBoxSelectedItem:
    With Err
        .Raise .Number, "clsComboList: ReturnListBoxSelectedItem", .Description
    End With
    
End Sub

Public Function ReturnLBSelectedItemsArray() As Long()
    Dim lngSelectCount As Long
    Dim lngItems() As Long
    Dim lngRet As Long
    Dim i As Integer
    
    On Error GoTo err_ReturnLBSelectedItemsArray
    
    lngSelectCount = SendMessage(m_lngListboxHwnd, LB_GETSELCOUNT, ByVal 0&, 0&)
    
    If lngSelectCount > 0 Then
        ReDim lngItems(lngSelectCount - 1) As Long
        ' Get the indexes of all selected items.
        lngRet = SendMessage(m_lngListboxHwnd, LB_GETSELITEMS, ByVal lngSelectCount, lngItems(0))
        
        ' fill array with ItemData values
        For i = 0 To lngSelectCount - 1
            lngRet = SendMessage(m_lngListboxHwnd, LB_GETITEMDATA, ByVal lngItems(i), ByVal 0&)
            lngItems(i) = lngRet
        Next i
        ' set function to return array to client
        ReturnLBSelectedItemsArray = lngItems
    End If
    
    Exit Function
    
err_ReturnLBSelectedItemsArray:
    With Err
        .Raise .Number, "clsComboList: ReturnLBSelectedItemsArray", .Description
    End With
    
End Function

Public Function ReturnListBoxSelectedText() As String()
    Dim strSelectedText As String
    Dim arrSelectedText() As String
    Dim lngSelectCount As Long
    Dim lngItems() As Long
    Dim lngRet As Long
    Dim lngTextLen As Long
    Dim i As Integer
    
    On Error GoTo err_ReturnListBoxSelectedText
    
    lngSelectCount = SendMessage(m_lngListboxHwnd, LB_GETSELCOUNT, ByVal 0&, 0&)
    If lngSelectCount > 0 Then
        ReDim lngItems(lngSelectCount - 1) As Long
        ReDim arrSelectedText(lngSelectCount - 1) As String
        ' Get the indexes of all selected items.
        lngRet = SendMessage(m_lngListboxHwnd, LB_GETSELITEMS, ByVal lngSelectCount, lngItems(0))
        
        For i = 0 To lngSelectCount - 1
            lngTextLen = SendMessage(m_lngListboxHwnd, LB_GETTEXTLEN, ByVal lngItems(i), ByVal CLng(0))
            strSelectedText = Space(lngTextLen) & vbNullChar
            ' Get the item text and remove the trailing null.
            lngTextLen = SendMessage(m_lngListboxHwnd, LB_GETTEXT, ByVal lngItems(i), ByVal strSelectedText)
            strSelectedText = Left$(strSelectedText, lngTextLen)

            arrSelectedText(i) = strSelectedText
        Next i
    
    End If
    ReturnListBoxSelectedText = arrSelectedText
    Exit Function
    
err_ReturnListBoxSelectedText:
    With Err
        .Raise .Number, "clsComboList: ReturnListBoxSelectedText", .Description
    End With
    
End Function

Public Function FindListboxItem(strSearchItem As String, Optional bolExactMatch As Boolean = True, _
    Optional lngStartPt As Long = -1) As Boolean
    Dim lngRet As Long
    
    If bolExactMatch Then
        ' find an exact match (-1 means search the entire list - the starting point)
        lngRet = SendMessage(m_lngListboxHwnd, LB_FINDSTRINGEXACT, lngStartPt, ByVal strSearchItem)
    Else
        lngRet = SendMessage(m_lngListboxHwnd, LB_FINDSTRING, lngStartPt, ByVal strSearchItem)
    End If
    If lngRet <> -1 Then
        ' display found item
        lngRet = SendMessage(m_lngListboxHwnd, LB_SETCURSEL, ByVal lngRet&, ByVal 0&)
        FindListboxItem = True
    Else
        FindListboxItem = False
    End If
    
End Function

Public Sub LoadListBoxADOArray(arrData As Variant, Optional bolClearList As Boolean = True)
    Dim strBuff As String
    Dim lngRet As Long
    Dim lnghwnd As Long
    Dim i As Integer
    Dim j As Integer
    Dim lngItemData As Long
    
    On Error GoTo err_LoadListBoxADOArray
    
    ' set long variable lngHwnd to hWnd value of the combo box.
    lnghwnd = m_lngListboxHwnd
    j = 0
    LockWindowUpdate m_lngListboxHwnd
    
    If bolClearList Then
        ' clear the list box
        ClearListbox
    End If
    
    ' arrData is an Array of recordset values returned from an ADO call
    For i = 0 To UBound(arrData, 2)
         If Not IsNull(arrData(1, i)) Then
            ' you CANNOT reference the array or object such as a recordset directly
            ' in the SendMessage call, so set the value to a simple variable (strBuff)
            strBuff = arrData(1, i)
            lngItemData = arrData(0, i)
            ' add the item to the listbox
            lngRet = SendMessage(lnghwnd, LB_ADDSTRING, 0, ByVal strBuff)
            ' add index of the item to itemdata property
            lngRet = SendMessage(lnghwnd, LB_SETITEMDATA, i, ByVal lngItemData)
            ' we need j as an incrementor in case any of the (arrData(1, i))
            ' elements are NULL because of the SendMessage(lnghwnd, LB_SETITEMDATA, i, ByVal lngItemData)
            ' call which sets the item data property (the count must be uninterupted (which WILL
            ' occur when the (arrData(1, i) is NULL)
            j = j + 1
        End If
    Next i
    ' set listbox count (numbers of items in the list)
    m_lngListCount = SendMessage(lnghwnd, LB_GETCOUNT, ByVal 0, 0)
    LockWindowUpdate 0
    Exit Sub
    
err_LoadListBoxADOArray:
    With Err
        .Raise .Number, "clsComboList: LoadListBoxADOArray", .Description
    End With

End Sub

Public Sub LoadListBoxADORecordset(rsData As ADODB.Recordset, lngListFieldNum As Long, _
        Optional lngIdFieldNum As Long = -1, Optional bolAddBlank As Boolean = True)
    Dim strBuff As String
    Dim lngRet As Long
    Dim objDisplayField As Field
    Dim objIDField As Field
    Dim i As Integer
    Dim j As Integer
    Dim lngItemData As Long
    
    On Error GoTo err_LoadListBoxADOArray
    
    j = 0
    LockWindowUpdate m_lngListboxHwnd
    ' clear the Combo box
    ClearListbox
    ' add an empty value
    strBuff = "  "
    ' add empty blank to combo list if true
    If bolAddBlank Then lngRet = SendMessage(m_lngListboxHwnd, LB_ADDSTRING, 0, ByVal strBuff)
    m_bolBlankAdded = bolAddBlank
    ' rsData is a recordset values returned from an ADO call
    With rsData
        .MoveFirst
        Set objDisplayField = .Fields(lngListFieldNum)
        If lngIdFieldNum <> -1 Then Set objIDField = .Fields(lngIdFieldNum)
        For i = 1 To .RecordCount
            ' disregard NULLS
            If Not IsNull(objDisplayField) Then
                ' you CANNOT reference the array or object such as a recordset directly
                ' in the SendMessage call, so set the value to a simple variable (strBuff)
                strBuff = objDisplayField.Value
                If lngIdFieldNum <> -1 Then
                    lngItemData = objIDField.Value
                Else
                    lngItemData = i
                End If

                ' add the item to the Combobox
                lngRet = SendMessage(m_lngListboxHwnd, LB_ADDSTRING, 0, ByVal strBuff)
                 ' add index of the item to itemdata property
                lngRet = SendMessage(m_lngListboxHwnd, LB_SETITEMDATA, j, ByVal lngItemData)
                ' fire off event  - used by client to increment progressbar
                RaiseEvent IncrementProgressBar(j)
                ' we need j as an incrementor in case any of the (arrData(1, i))
                ' elements are NULL because of the SendMessage(lnghwnd, CB_SETITEMDATA, i, ByVal lngItemData)
                ' call which sets the item data property (the count must be uninterupted (which WILL
                ' occur when the (arrData(1, i) is NULL)
                j = j + 1
            End If
            .MoveNext
        Next i
    End With
    ' set Combobox count (numbers of items in the Combo)
    m_lngListCount = SendMessage(m_lngListboxHwnd, LB_GETCOUNT, ByVal 0, 0)
    'lngRet = SendMessage(lnghwnd, CB_SETCURSEL, 0, 0)
    LockWindowUpdate 0
    Exit Sub
    
err_LoadListBoxADOArray:
    With Err
        .Raise .Number, "clsComboList: LoadListBoxADOArray", .Description
    End With

End Sub

Public Sub LoadListBoxArray(arrData As Variant, Optional bolClearList As Boolean = True)
    Dim strBuff As String
    Dim lngRet As Long
    Dim i As Integer
    Dim lngItemData As Long
    
    On Error GoTo err_LoadListBoxArray
    
    ' make sure argument array is valid
    If Not CheckValidArray(arrData) Then
        ' we still may need to clear the listbox
        If bolClearList Then ClearListbox
        Exit Sub
    End If
    
    LockWindowUpdate m_lngListboxHwnd
    
    If bolClearList Then
        ' clear the list box
        ClearListbox
    End If
    
    For i = 0 To UBound(arrData) - 1
        If Not IsNull(arrData) Then
            ' you CANNOT reference the array or object such as a recordset directly
            ' in the SendMessage call, so set the value to a simple variable (strBuff)
            strBuff = arrData(i)
            ' add the item to the listbox
            lngRet = SendMessage(m_lngListboxHwnd, LB_ADDSTRING, 0, ByVal strBuff)
        End If
    Next i
    LockWindowUpdate 0
    Exit Sub
    
err_LoadListBoxArray:
    With Err
        .Raise .Number, "clsComboList: LoadListBoxArray", .Description
    End With

End Sub

Public Sub ClearListbox()
    Dim lngRet As Long
    
    ' clear the list box
    lngRet = SendMessage(m_lngListboxHwnd, LB_RESETCONTENT, ByVal 0, 0)
    m_lngListCount = 0
    
End Sub

Public Sub AddFilesToListbox(strpath As String, Optional bolIncludeHidden _
        As Boolean = False, Optional bolIncludeSystem As Boolean = False, _
        Optional bolClearList As Boolean = True)
    Dim lngRet As Long
    
    ' clear listbox
    If bolClearList Then ClearListbox
    
    ' include only archived files
    If Not bolIncludeHidden And Not bolIncludeSystem Then
        lngRet = SendMessage(m_lngListboxHwnd, LB_DIR, DDL_EXCLUSIVE Or DDL_ARCHIVE, ByVal strpath)
    ElseIf bolIncludeHidden And Not bolIncludeSystem Then
        lngRet = SendMessage(m_lngListboxHwnd, LB_DIR, DDL_EXCLUSIVE Or DDL_ARCHIVE _
            Or DDL_HIDDEN, ByVal strpath)
    ElseIf Not bolIncludeHidden And bolIncludeSystem Then
        lngRet = SendMessage(m_lngListboxHwnd, LB_DIR, DDL_EXCLUSIVE Or DDL_ARCHIVE _
            Or DDL_SYSTEM, ByVal strpath)
    Else
        lngRet = SendMessage(m_lngListboxHwnd, LB_DIR, DDL_EXCLUSIVE Or DDL_ARCHIVE _
            Or DDL_HIDDEN Or DDL_SYSTEM, ByVal strpath)
    End If
    
End Sub

Public Sub AddFoldersToListbox(strpath As String, Optional bolClearList As Boolean = True)
    Dim lngRet As Long
    
    ' clear listbox
    If bolClearList Then ClearListbox
    ' add directories to listbox
    lngRet = SendMessage(m_lngListboxHwnd, LB_DIR, DDL_DIRECTORY, ByVal strpath)
    
End Sub

Public Sub AddDrivesToListbox(strpath As String, Optional bolClearList As Boolean = True)
    Dim lngRet As Long
    
    ' clear listbox
    If bolClearList Then ClearListbox
    ' add directories to listbox
    lngRet = SendMessage(m_lngListboxHwnd, LB_DIR, DDL_DRIVES, ByVal strpath)
    
End Sub

Public Sub SetLBTabStops(TabStops() As Long)
    ' Written by David Stevens 3/29/2000 5:33:21 PM
    ' Code found at Planet Source Code (www.planet-source-code.com)
    '----- Local Variables used in this function -----'
    ' 10/17/2001 MJN changed
    Dim alTabStops() As Long
    Dim lCtr As Long
    Dim lColumns As Long
    Dim lRet As Long
    
    'PURPOSE: Set TabSTops for a list box using the hwnd property.
    'This creates columns separated by a tab character
    
    'USAGE:
    'Pass ListBox Object and a comma delimited
    'list of tab stops.  Tab stops are expressed
    'in dialog units which approximately equal
    '1/4 the width of a character
    
    On Error GoTo err_SetLBTabStops:
    
    ' make sure argument array is valid
    If Not CheckValidArray(TabStops) Then Exit Sub
    
    ReDim alTabStops(UBound(TabStops)) As Long
    
    For lCtr = 0 To UBound(TabStops)
        alTabStops(lCtr) = TabStops(lCtr)
    Next
    
    lColumns = UBound(alTabStops) + 1
    
    
    lRet = SendMessage(m_lngListboxHwnd, LB_SETTABSTOPS, lColumns, alTabStops(0))
    
    Exit Sub
    
err_SetLBTabStops:
    With Err
        .Raise .Number, "clsComboList: SetLBTabStops", .Description
    End With

End Sub

' code to put horizontal scrollbars on the list box
Public Sub RedrawListbox()
    Dim intNewWidth As Integer
    Dim lngRet As Long
        
    ' set the scrollbars
    lngRet = SendMessage(m_lngListboxHwnd, LB_SETHORIZONTALEXTENT, m_objListbox.Width, ByVal 0&)

    ' can dynamically reset width of list box (to take into account the adding of the scrollbar)
    intNewWidth = m_objListbox.Width + 100 'new width in pixels
    
    ' set the scrollbars to the wider list box
    lngRet = SendMessage(m_lngListboxHwnd, LB_SETHORIZONTALEXTENT, intNewWidth, ByVal 0&)

End Sub

Public Sub SelectAllListboxItems()
    Dim lngListCount As Long
    Dim lngItem As Long
    Dim lngRet As Long
    
    ' get count of list items
    lngListCount = SendMessage(m_lngListboxHwnd, LB_GETCOUNT, ByVal 0&, ByVal 0&)
    
    If lngListCount > 0 Then
        LockWindowUpdate m_lngListboxHwnd
        lngListCount = lngListCount - 1
        For lngItem = 0 To lngListCount
            lngRet = SendMessage(m_lngListboxHwnd, LB_SETSEL, ByVal 1&, ByVal lngItem)
        Next lngItem
        ' display top of list (with all selected)
        lngRet = SendMessage(m_lngListboxHwnd, LB_SETSEL, ByVal 1&, ByVal 0&)
        LockWindowUpdate 0
    End If
    
End Sub

Public Sub DeselectAllListboxItems()
    Dim lngRet As Long
    
    lngRet = SendMessage(m_lngListboxHwnd, LB_SETSEL, ByVal 0&, ByVal -1&)
    
End Sub

Public Sub DeleteListboxText(lngItemIndex As Long)
    Dim lngRet As Long
    
    On Error GoTo err_DeleteListboxText
    
    lngRet = SendMessage(m_lngListboxHwnd, LB_DELETESTRING, ByVal lngItemIndex, ByVal 0&)
    
    If lngRet = LB_ERR Then
        Err.Raise vbObjectError + 1070, , "Error deleting listbox text"
    End If
    
    ' set listbox count (numbers of items in the listbox)
    m_lngListCount = SendMessage(m_lngListboxHwnd, LB_GETCOUNT, ByVal 0, 0)
    
    Exit Sub
    
err_DeleteListboxText:
    With Err
        .Raise .Number, "clsComboList: DeleteListboxText", .Description
    End With
    
End Sub

Public Sub SetListboxTabStops(ParamArray TabStops())
    ' Written by David Stevens 3/29/2000 5:33:21 PM
    ' Code found at Planet Source Code (www.planet-source-code.com)
    '----- Local Variables used in this function -----'
    Dim alTabStops() As Long
    Dim lCtr As Long
    Dim lColumns As Long
    Dim lRet As Long
    
    'PURPOSE: Set TabSTops for a list box using the hwnd property.
    'This creates columns separated by a tab character
    
    'USAGE:
    'Pass ListBox Object and a comma delimited
    'list of tab stops.  Tab stops are expressed
    'in dialog units which approximately equal
    '1/4 the width of a character
    
    On Error GoTo err_SetListboxTabStops:
    
    ReDim alTabStops(UBound(TabStops)) As Long
    
    For lCtr = 0 To UBound(TabStops)
    alTabStops(lCtr) = TabStops(lCtr)
    Next
    
    lColumns = UBound(alTabStops) + 1
    
    lRet = SendMessage(m_lngListboxHwnd, LB_SETTABSTOPS, lColumns, alTabStops(0))
    
    Exit Sub
    
err_SetListboxTabStops:
    With Err
        .Raise .Number, "clsComboList: SetListboxTabStops", .Description
    End With

End Sub

Public Sub ResetComboWidth(Optional intAddWidth As Integer = 300)
    ' set the dropdown width to user set width (Pixels)
    SendMessage m_lngComboboxHwnd, CB_SETDROPPEDWIDTH, intAddWidth, 0
End Sub

Public Function ReturnComboWidth() As Long
    ' return the dropdown width (if needed)
    ReturnComboWidth = SendMessage(m_lngComboboxHwnd, CB_GETDROPPEDWIDTH, 0, 0)
End Function

Public Sub ShowComboDropDown(bolState As Boolean)
    ' immediately shows combobox dropdown
    SendMessage m_lngComboboxHwnd, CB_SHOWDROPDOWN, ByVal bolState, 0
End Sub

Public Sub SetListboxListIndex(lngListIndex As Long)
    Dim lngRet As Long
    
    ' set listbox listindex
    lngRet = SendMessage(m_lngListboxHwnd, LB_SETCURSEL, ByVal lngListIndex, 0)
    
End Sub

Public Function ReturnListboxListIndex() As Long
    ' obtain list listindex
    ReturnListboxListIndex = SendMessage(m_lngListboxHwnd, LB_GETCURSEL, 0, 0)
End Function

Public Sub AddSingleItemToList(strBuff As String, Optional lngItemData As Long = -1, _
    Optional bolMakeCurrentItem As Boolean = True)
    Dim lngRet As Long
    
    On Error GoTo err_AddSingleItemToList
    
    ' add the item to the Listbox
    lngRet = SendMessage(m_lngListboxHwnd, LB_ADDSTRING, 0, ByVal strBuff)
    If lngItemData <> -1 Then
        ' add index of the item to itemdata property
        lngRet = SendMessage(m_lngListboxHwnd, LB_SETITEMDATA, lngRet, ByVal lngItemData)
    End If
    ' set listbox count (numbers of items in the listbox)
    m_lngListCount = SendMessage(m_lngListboxHwnd, LB_GETCOUNT, ByVal 0, 0)
    If bolMakeCurrentItem Then
        ' display added text at in combo text
        lngRet = SendMessage(m_lngListboxHwnd, LB_SETCURSEL, m_lngListCount - 1, 0)
    End If
    
    Exit Sub
    
err_AddSingleItemToList:
    With Err
        .Raise .Number, "clsComboList: AddSingleItemToList", .Description
    End With
    
End Sub

Public Sub InsertSingleItemToList(strBuff As String, lngItemData As Long, _
    lngIndexPosition As Long, bolMakeCurrentItem As Boolean)
    Dim lngRet As Long
    
    On Error GoTo err_InsertSingleItemToList
    
    ' insert the item to the Listbox
    lngRet = SendMessage(m_lngListboxHwnd, LB_INSERTSTRING, lngIndexPosition, ByVal strBuff)
    ' add index of the item to itemdata property
    lngRet = SendMessage(m_lngListboxHwnd, LB_SETITEMDATA, lngRet, ByVal lngItemData)
    ' set listbox count (numbers of items in the listbox)
    m_lngListCount = SendMessage(m_lngListboxHwnd, LB_GETCOUNT, ByVal 0, 0)
    If bolMakeCurrentItem Then
        ' display added text at in combo text
        lngRet = SendMessage(m_lngListboxHwnd, LB_SETCURSEL, lngIndexPosition, 0)
    End If
    
    Exit Sub
    
err_InsertSingleItemToList:
    With Err
        .Raise .Number, "clsComboList: InsertSingleItemToList", .Description
    End With
    
End Sub

Public Function DeleteListboxItem(lngListIndex As Long) As Boolean
    Dim lngRet As Long
    
    On Error GoTo err_DeleteListboxItem
    
    DeleteListboxItem = False
    
    lngRet = SendMessage(m_lngListboxHwnd, LB_DELETESTRING, ByVal lngListIndex, 0)
    
    ' set listbox count (numbers of items in the listbox)
    m_lngListCount = SendMessage(m_lngListboxHwnd, LB_GETCOUNT, ByVal 0, 0)
    
    DeleteListboxItem = True
    
    Exit Function
    
err_DeleteListboxItem:
    With Err
        .Raise .Number, "clsComboList: DeleteListboxItem", .Description
    End With
    
End Function

Public Sub LoadComboBoxADOArray(arrData As Variant, Optional bolAddBlank As Boolean = True)
    Dim strBuff As String
    Dim lngRet As Long
    Dim lnghwnd As Long
    Dim i As Integer
    Dim j As Integer
    Dim lngItemData As Long
    
    On Error GoTo err_LoadComboBoxADOArray
    
    ' set long variable lngHwnd to hWnd value of the combo box.
    lnghwnd = m_lngComboboxHwnd
    j = 0
    LockWindowUpdate m_lngComboboxHwnd
    ' clear the Combo box
    lngRet = SendMessage(lnghwnd, CB_RESETCONTENT, ByVal 0, 0)
    ' add an empty value
    strBuff = "  "
    If bolAddBlank Then lngRet = SendMessage(m_lngComboboxHwnd, CB_ADDSTRING, 0, ByVal strBuff)
    m_bolBlankAdded = bolAddBlank
    ' arrData is an Array of recordset values returned from an ADO call
    For i = 0 To UBound(arrData, 2)
        If Not IsNull(arrData(0, i)) Then
            ' you CANNOT reference the array or object such as a recordset directly
            ' in the SendMessage call, so set the value to a simple variable (strBuff)
            strBuff = arrData(0, i)
           ' lngItemData = arrData(1, i)
            ' add the item to the Combobox
            lngRet = SendMessage(lnghwnd, CB_ADDSTRING, 0, ByVal strBuff)
            ' add index of the item to itemdata property
           ' lngRet = SendMessage(lnghwnd, CB_SETITEMDATA, j, ByVal lngItemData)
            ' fire off event  - used by client to increment progressbar
            RaiseEvent IncrementProgressBar(j)
            ' we need j as an incrementor in case any of the (arrData(1, i))
            ' elements are NULL because of the SendMessage(lnghwnd, CB_SETITEMDATA, i, ByVal lngItemData)
            ' call which sets the item data property (the count must be uninterupted (which WILL
            ' occur when the (arrData(1, i) is NULL)
            j = j + 1
        End If
    Next i
    ' set Combobox count (numbers of items in the Combo)
    m_lngComboCount = SendMessage(lnghwnd, CB_GETCOUNT, ByVal 0, 0)
    'lngRet = SendMessage(lnghwnd, CB_SETCURSEL, 0, 0)
    LockWindowUpdate 0
    Exit Sub
    
err_LoadComboBoxADOArray:
    With Err
        .Raise .Number, "clsComboList: LoadComboBoxADOArray", .Description
    End With

End Sub

Public Sub LoadComboBoxArray(arrData As Variant, Optional bolAddBlank As Boolean = True)
    Dim strBuff As String
    Dim lngRet As Long
    Dim i As Integer
    Dim lngItemData As Long
    
    On Error GoTo err_LoadComboBoxArray
    
    ' make sure argument array is valid
    If Not CheckValidArray(arrData) Then Exit Sub
    
    LockWindowUpdate m_lngComboboxHwnd
    ' clear the Combo box
    lngRet = SendMessage(m_lngComboboxHwnd, CB_RESETCONTENT, ByVal 0, 0)
    ' add an empty value
    strBuff = "  "
    If bolAddBlank Then lngRet = SendMessage(m_lngComboboxHwnd, CB_ADDSTRING, 0, ByVal strBuff)
    m_bolBlankAdded = bolAddBlank
    ' arrData is an Array values
    For i = 0 To UBound(arrData) - 1
        If Not IsNull(arrData) Then
            ' you CANNOT reference the array or object such as a recordset directly
            ' in the SendMessage call, so set the value to a simple variable (strBuff)
            strBuff = arrData(i)
            ' add the item to the Combobox
            lngRet = SendMessage(m_lngComboboxHwnd, CB_ADDSTRING, 0, ByVal strBuff)
        End If
    Next i
    LockWindowUpdate 0
    Exit Sub
    
err_LoadComboBoxArray:
    With Err
        .Raise .Number, "clsComboList: LoadComboBoxArray", .Description
    End With

End Sub

Public Sub LoadComboBoxADORecordset(rsData As ADODB.Recordset, lngListFieldNum As Long, _
        lngIdFieldNum As Long, Optional bolAddBlank As Boolean = True)
    Dim strBuff As String
    Dim lngRet As Long
    Dim objDisplayField As Field
    Dim objIDField As Field
    Dim i As Integer
    Dim j As Integer
    Dim lngItemData As Long
    
    On Error GoTo err_LoadComboBoxADOArray
    
    j = 0
    LockWindowUpdate m_lngComboboxHwnd
    ' clear the Combo box
    lngRet = SendMessage(m_lngComboboxHwnd, CB_RESETCONTENT, ByVal 0, 0)
    ' add an empty value
    strBuff = "  "
    ' add empty blank to combo list if true
    If bolAddBlank Then lngRet = SendMessage(m_lngComboboxHwnd, CB_ADDSTRING, 0, ByVal strBuff)
    m_bolBlankAdded = bolAddBlank
    ' rsData is a recordset values returned from an ADO call
    With rsData
        .MoveFirst
        Set objDisplayField = .Fields(lngListFieldNum)
        Set objIDField = .Fields(lngIdFieldNum)
        For i = 1 To .RecordCount
            ' disregard NULLS
            If Not IsNull(objDisplayField) Then
                ' you CANNOT reference the array or object such as a recordset directly
                ' in the SendMessage call, so set the value to a simple variable (strBuff)
                strBuff = objDisplayField.Value
                lngItemData = objIDField.Value

                ' add the item to the Combobox
                lngRet = SendMessage(m_lngComboboxHwnd, CB_ADDSTRING, 0, ByVal strBuff)
                 ' add index of the item to itemdata property
                lngRet = SendMessage(m_lngComboboxHwnd, CB_SETITEMDATA, j, ByVal lngItemData)
                ' fire off event  - used by client to increment progressbar
                RaiseEvent IncrementProgressBar(j)
                ' we need j as an incrementor in case any of the (arrData(1, i))
                ' elements are NULL because of the SendMessage(lnghwnd, CB_SETITEMDATA, i, ByVal lngItemData)
                ' call which sets the item data property (the count must be uninterupted (which WILL
                ' occur when the (arrData(1, i) is NULL)
                j = j + 1
            End If
            .MoveNext
        Next i
    End With
    ' set Combobox count (numbers of items in the Combo)
    m_lngComboCount = SendMessage(m_lngComboboxHwnd, CB_GETCOUNT, ByVal 0, 0)
    'lngRet = SendMessage(lnghwnd, CB_SETCURSEL, 0, 0)
    LockWindowUpdate 0
    Exit Sub
    
err_LoadComboBoxADOArray:
    With Err
        .Raise .Number, "clsComboList: LoadComboBoxADOArray", .Description
    End With

End Sub

Public Sub AddSingleItemToCombo(strBuff As String, Optional lngItemData As Long = -1, _
    Optional bolMakeCurrentItem As Boolean = True)
    Dim lngRet As Long
    
    On Error GoTo err_AddSingleItemToCombo
    
    ' add the item to the Combobox (bottom of list)
    lngRet = SendMessage(m_lngComboboxHwnd, CB_ADDSTRING, 0, ByVal strBuff)
    If lngItemData <> -1 Then
        ' add index of the item to itemdata property
        lngRet = SendMessage(m_lngComboboxHwnd, CB_SETITEMDATA, lngRet, ByVal lngItemData)
    End If
    ' set Combobox count (numbers of items in the Combo)
    m_lngComboCount = SendMessage(m_lngComboboxHwnd, CB_GETCOUNT, ByVal 0, 0)
    If bolMakeCurrentItem Then
        ' display added text at in combo text
        lngRet = SendMessage(m_lngComboboxHwnd, CB_SETCURSEL, m_lngComboCount - 1, 0)
    End If
    
    Exit Sub
    
err_AddSingleItemToCombo:
    With Err
        .Raise .Number, "clsComboList: AddSingleItemToCombo", .Description
    End With
End Sub

Public Sub InsertSingleItemToCombo(strBuff As String, lngItemData As Long, _
    lngIndexPosition As Long, bolMakeCurrentItem As Boolean)
    Dim lngRet As Long
    
    On Error GoTo err_InsertSingleItemToCombo
    
    ' insert the item to the Combobox
    lngRet = SendMessage(m_lngComboboxHwnd, CB_INSERTSTRING, lngIndexPosition, ByVal strBuff)
    ' add index of the item to itemdata property
    lngRet = SendMessage(m_lngComboboxHwnd, CB_SETITEMDATA, lngRet, ByVal lngItemData)
    ' set Combobox count (numbers of items in the Combo)
    m_lngComboCount = SendMessage(m_lngComboboxHwnd, CB_GETCOUNT, ByVal 0, 0)
    If bolMakeCurrentItem Then
        ' display added text at in combo text
        lngRet = SendMessage(m_lngComboboxHwnd, CB_SETCURSEL, lngIndexPosition, 0)
    End If
    
    Exit Sub
    
err_InsertSingleItemToCombo:
    With Err
        .Raise .Number, "clsComboList: InsertSingleItemToCombo", .Description
    End With
End Sub

Public Function FindComboListItem(strSearchItem As String, Optional bolExactMatch As Boolean = True, _
    Optional lngStartPt As Long = -1) As Boolean
    Dim lngRet As Long
    
    If bolExactMatch Then
        ' find an exact match (-1 means search the entire list - the starting point)
        lngRet = SendMessage(m_lngComboboxHwnd, CB_FINDSTRINGEXACT, lngStartPt, ByVal strSearchItem)
    Else
        lngRet = SendMessage(m_lngComboboxHwnd, CB_FINDSTRING, lngStartPt, ByVal strSearchItem)
    End If
    If lngRet <> -1 Then
        ' display found item
        lngRet = SendMessage(m_lngComboboxHwnd, CB_SETCURSEL, ByVal lngRet&, ByVal 0&)
        FindComboListItem = True
    Else
        FindComboListItem = False
    End If
    
End Function

Public Function ReturnComboSelectedItem() As String
    Dim lngRet As Long
    Dim strBuffer As String
    
    strBuffer = String(255, vbNull)
    
    lngRet = SendMessage(m_lngComboboxHwnd, CB_GETCURSEL, 0, 0)
    ' make sure something is selected
    If lngRet > -1 Then
        lngRet = SendMessage(m_lngComboboxHwnd, CB_GETLBTEXT, ByVal lngRet, ByVal strBuffer)
        ' trim the selected item down
        ReturnComboSelectedItem = Left$(strBuffer, lngRet)
    End If
    
End Function

Public Function ReturnComboItemData() As Long
    Dim lngRet As Long
    
    lngRet = SendMessage(m_lngComboboxHwnd, CB_GETCURSEL, 0, 0)
    ' make sure something is selected
    If lngRet > -1 Then
        If m_bolBlankAdded Then
            ' return the itemdata value
            ReturnComboItemData = SendMessage(m_lngComboboxHwnd, CB_GETITEMDATA, ByVal lngRet - 1, ByVal 0)
        Else
            ReturnComboItemData = SendMessage(m_lngComboboxHwnd, CB_GETITEMDATA, ByVal lngRet, ByVal 0)
        End If
    End If
    
End Function

Public Function ReturnComboListIndex() As Long
    ' obtain combo listindex
    ReturnComboListIndex = SendMessage(m_lngComboboxHwnd, CB_GETCURSEL, 0, 0)
End Function
    
Public Sub SetComboListIndex(lngListIndex As Long)
    Dim lngRet As Long
    
    ' set combo listindex
    lngRet = SendMessage(m_lngComboboxHwnd, CB_SETCURSEL, ByVal lngListIndex, 0)
    
End Sub

Public Function DeleteComboListItem(lngListIndex As Long) As Boolean
    Dim lngRet As Long
    
    On Error GoTo err_DeleteComboListItem
    
    DeleteComboListItem = False
    
    lngRet = SendMessage(m_lngComboboxHwnd, CB_DELETESTRING, ByVal lngListIndex, 0)
    
    ' set Combobox count (numbers of items in the Combo)
    m_lngComboCount = SendMessage(m_lngComboboxHwnd, CB_GETCOUNT, ByVal 0, 0)
    
    DeleteComboListItem = True
    
    Exit Function
    
err_DeleteComboListItem:
    With Err
        .Raise .Number, "clsComboList: DeleteComboListItem", .Description
    End With
    
End Function

Public Sub AddFilesToCombobox(strpath As String, Optional bolIncludeHidden _
        As Boolean = False, Optional bolIncludeSystem As Boolean = False, _
        Optional bolClearList As Boolean = True)
    Dim lngRet As Long
    
    ' clear listbox
    If bolClearList Then ClearCombobox
    
    ' include only archived files
    If Not bolIncludeHidden And Not bolIncludeSystem Then
        lngRet = SendMessage(m_lngListboxHwnd, CB_DIR, DDL_EXCLUSIVE Or DDL_ARCHIVE, ByVal strpath)
    ElseIf bolIncludeHidden And Not bolIncludeSystem Then
        lngRet = SendMessage(m_lngListboxHwnd, CB_DIR, DDL_EXCLUSIVE Or DDL_ARCHIVE _
            Or DDL_HIDDEN, ByVal strpath)
    ElseIf Not bolIncludeHidden And bolIncludeSystem Then
        lngRet = SendMessage(m_lngListboxHwnd, CB_DIR, DDL_EXCLUSIVE Or DDL_ARCHIVE _
            Or DDL_SYSTEM, ByVal strpath)
    Else
        lngRet = SendMessage(m_lngListboxHwnd, CB_DIR, DDL_EXCLUSIVE Or DDL_ARCHIVE _
            Or DDL_HIDDEN Or DDL_SYSTEM, ByVal strpath)
    End If
    
End Sub

Public Sub AddFoldersToCombobox(strpath As String, Optional bolClearList As Boolean = True)
    Dim lngRet As Long
    
    ' clear listbox
    If bolClearList Then ClearCombobox
    ' add directories to listbox
    lngRet = SendMessage(m_lngListboxHwnd, CB_DIR, DDL_DIRECTORY, ByVal strpath)
    
End Sub

Public Sub AddDrivesToCombobox(strpath As String, Optional bolClearList As Boolean = True)
    Dim lngRet As Long
    
    ' clear listbox
    If bolClearList Then ClearCombobox
    ' add directories to listbox
    lngRet = SendMessage(m_lngListboxHwnd, CB_DIR, DDL_DRIVES, ByVal strpath)
    
End Sub

Public Sub ClearCombobox()
    Dim lngRet As Long
    
    ' clear the Combo box
    lngRet = SendMessage(m_lngComboboxHwnd, CB_RESETCONTENT, ByVal 0, 0)
    m_lngComboCount = 0
    
End Sub

Private Function CheckValidArray(varArray As Variant) As Boolean
    ' make sure we have a valid array before we work with it
    Dim lngUpper As Long
    
    On Error Resume Next
    
    ' check upper bound
    lngUpper = UBound(varArray)
    If Err.Number = 9 Then
        CheckValidArray = False
        Err.Clear
    Else
        CheckValidArray = True
    End If
    
End Function
