VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormUtilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'===============================================================
' Module:   clsFormUtilities
'
' Project:
'
' Purpose:  Provides Form GUI functionality via Win32 APIs
'
' Author:   Michael J. Nugent
'
' Date:
'
'===============================================================
Option Explicit
    
Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

' Toggles enabled state of menu item.
Private Const MF_ENABLED = &H0&
Private Const MF_GRAYED = &H1&
Private Const MF_DISABLED = &H2&

' System menu command values commonly used by VB.
Private Const SC_SIZE = &HF000&
Private Const SC_MOVE = &HF010&
Private Const SC_MINIMIZE = &HF020&
Private Const SC_MAXIMIZE = &HF030&
Private Const SC_CLOSE = &HF060&
Private Const SC_RESTORE = &HF120&

' Menu information constants.
Private Const MIIM_STATE As Long = &H1
Private Const MIIM_ID As Long = &H2
Private Const MIIM_SUBMENU As Long = &H4
Private Const MIIM_CHECKMARKS As Long = &H8
Private Const MIIM_TYPE As Long = &H10
Private Const MIIM_DATA As Long = &H20
Private Const MF_BYPOSITION = &H400


' Used to get window style bits.
Private Const GWL_STYLE = (-16)

' Style bits.
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_THICKFRAME = &H40000
Private Const WS_MINIMIZEBOX = &H20000

' position of the Maximize and Size selections on a forms system menu
' (0-based menu array)
Private Const cnMAX_MENU_POSITION = &H4
Private Const cnSIZE_MENU_POSITION = &H2

Private Declare Function GetMenuItemInfo Lib "user32" _
    Alias "GetMenuItemInfoA" _
   (ByVal hMenu As Long, ByVal uItem As Long, _
    ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long

Private Declare Function GetMenuItemCount Lib "user32" _
   (ByVal hMenu As Long) As Long
      
Private Declare Function GetSystemMenu Lib "user32" _
   (ByVal hWnd As Long, _
    ByVal bRevert As Long) As Long
       
Private Declare Function DrawMenuBar Lib "user32" _
   (ByVal hWnd As Long) As Long
          
Private Declare Function SetWindowPos Lib "user32" _
    (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
     ByVal X As Long, ByVal Y As Long, _
     ByVal cx As Long, ByVal cy As Long, _
     ByVal wFlags As Long) As Long

Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" _
    (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Long, _
    lpcMenuItemInfo As MENUITEMINFO) As Long
    
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    
Private m_lng_hWnd As Long
    
Public Property Let Form_hWnd(vData As Long)
     m_lng_hWnd = vData
End Property

Public Sub DisableMaximizeMenu()
    DisableFormMenuSelection cnMAX_MENU_POSITION
End Sub

Public Sub DisableSizeMenu()
    DisableFormMenuSelection cnSIZE_MENU_POSITION
End Sub

Public Sub AddMinimizeButtonToDialogWindow()
    EnableWindowsFunction WS_MINIMIZEBOX
End Sub

Public Sub DisableFormMaxButton()
    DisableWindowsFunction WS_MAXIMIZEBOX
End Sub

Public Sub DisableFormResize()
    DisableWindowsFunction WS_THICKFRAME
End Sub

Public Sub SetToCustomDialog()
    ' access all the above subs to get to a moveable, nosize
    ' form that the user can still minimize
    
    DisableMaximizeMenu
    DisableSizeMenu
    DisableFormMaxButton
    DisableFormResize
    AddMinimizeButtonToDialogWindow
    
End Sub

Private Sub DisableFormMenuSelection(lngPosition As Long)
    ' Disables Maximize button of the calling form
    ' Modified (by MJN)from code found on the following
    ' sites:
    ' VBnet, Randy Birch - http://www.mvps.org/vbnet/
    ' Karl E. Peterson - http://www.mvps.org/vb
    Dim hMenu As Long
    Dim nPosition As Long
    Dim uFlags As Long
    Dim mii As MENUITEMINFO
    
    ' Retrieve handle to system menu.
    hMenu = GetSystemMenu(m_lng_hWnd, False)

    ' Initialize structure.
    With mii
        .cbSize = Len(mii)
        .fMask = MIIM_STATE Or MIIM_ID
        .dwTypeData = String$(80, 0)
        .cch = Len(mii.dwTypeData)
        GetMenuItemInfo hMenu, lngPosition, MF_BYPOSITION, mii
        .fState = MF_GRAYED
     
        ' New ID uses highbit to signify that
        ' the menu item is enabled.
        .wID = hMenu

        ' Modify the menu!
        .fMask = MIIM_STATE Or MIIM_ID
        SetMenuItemInfo hMenu, lngPosition, MF_BYPOSITION, mii
    End With
   
End Sub

Private Sub DisableWindowsFunction(lngFunction As Long)
   Dim nStyle As Long
   
   ' Retrieve current style bits.
   nStyle = GetWindowLong(m_lng_hWnd, GWL_STYLE)
   
   ' disable Maximize button
   nStyle = nStyle And Not lngFunction
   
   Call SetWindowLong(m_lng_hWnd, GWL_STYLE, nStyle)
   
End Sub

Private Sub EnableWindowsFunction(lngFunction As Long)
   Dim nStyle As Long
   
   ' Retrieve current style bits.
   nStyle = GetWindowLong(m_lng_hWnd, GWL_STYLE)
   
   ' enable Minimize button
   nStyle = nStyle Or lngFunction
   
   Call SetWindowLong(m_lng_hWnd, GWL_STYLE, nStyle)
   
End Sub

