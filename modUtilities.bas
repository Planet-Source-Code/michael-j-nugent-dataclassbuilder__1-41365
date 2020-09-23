Attribute VB_Name = "modUtilities"
'===============================================================
' Module:   ModUtilities
'
' Project:
'
' Purpose:  Application utilities module
'           Provides generic application
'           utility functions/subs.
'
' Author:   Michael J. Nugent
'
' Date:
'
'===============================================================
Option Explicit

' APIs used to look for, and set, any running application window
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    
Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type WINDOWPLACEMENT
        Length As Long
        flags As Long
        showCmd As Long
        ptMinPosition As POINTAPI
        ptMaxPosition As POINTAPI
        rcNormalPosition As RECT
End Type

Private Declare Function SetForegroundWindow Lib "user32" _
    (ByVal hwnd As Long) As Long
Private Declare Function GetWindowPlacement Lib "user32" _
    (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function SetWindowPlacement Lib "user32" _
    (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Const SW_SHOWNORMAL = 1
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_RESTORE = 9

'***** Find file API calls and types
Private Const MAX_PATH = 260

'***** Format API error strings
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" _
    (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, _
    Arguments As Long) As Long
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000

' get correct file path name
Private Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" _
    (ByVal lpFileName As String, ByVal nBufferLength As Long, _
    ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
    
' use to print text in picturebox
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, _
        ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
        ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
        ByVal dwRop As Long) As Long

Private Const SRCCOPY = &HCC0020

'****** used for API Shell routines
' open a process
Private Declare Function OpenProcess Lib "kernel32" _
    (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
    ByVal dwlngProcessID As Long) As Long
' close the process/handle
Private Declare Function CloseHandle Lib "kernel32" _
    (ByVal hObject As Long) As Long
' wait for process to finish before calling program continues
Private Declare Function WaitForSingleObject Lib "kernel32" _
    (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
' get exit code of process
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, _
    lpExitCode As Long) As Long
    
Private Const WAIT_FAILED = -1&        'Error on call
Private Const WAIT_OBJECT_0 = 0        'Normal completion
Private Const WAIT_ABANDONED = &H80&   'Wait abandoned
Private Const WAIT_TIMEOUT = &H102&    'Timeout period elapsed
Private Const IGNORE = 0               'Ignore signal
Private Const INFINITE = -1&           'Infinite timeout
Private Const SYNCHRONIZE = &H100000

' retrieve process information (status)
Private Const PROCESS_QUERY_INFORMATION = &H400
' process is still active
Private Const STILL_ACTIVE = &H103


' similar to doEvents - suspends the execution of the
' current thread for a specified interval.
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' allow the users to change the default system
' disabled text color to something more readable
Private Declare Function SetSysColors Lib "user32" _
    (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long
Private Declare Function GetSysColor Lib "user32" _
    (ByVal nIndex As Long) As Long
Private Const COLOR_GREYTEXT = 17
' default system disabled text color
' usually this value = 8421504
Private Const DEFAULT_COLOR = 8421504

' used to lock a window (or control) to prevent screen flicker
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public Function CheckNullValue(vCheckValue As Variant, vTargetValue As Variant) As Variant

    On Error Resume Next
    
    Select Case vTargetValue
        Case vbInteger, vbLong, vbDouble
            If IsNull(vCheckValue) Then
                CheckNullValue = 0
            Else
                CheckNullValue = vCheckValue
            End If
        Case vbString
            If IsNull(vCheckValue) Then
                CheckNullValue = ""
            Else
                CheckNullValue = vCheckValue
            End If
        Case vbDate
            If IsNull(vCheckValue) Then
                CheckNullValue = CDate(Empty)
            Else
                CheckNullValue = vCheckValue
            End If
            
        Case vbBoolean
            If IsNull(vCheckValue) Then
                CheckNullValue = ""
            Else
                CheckNullValue = vCheckValue
            End If
        Case Else
            ' do nothing
            
    End Select
    
End Function

Public Function SetStringToBoolean(vBooleanValue As Variant) As String
    On Error Resume Next
        
    If UCase(vBooleanValue) = "FALSE" Or vBooleanValue = "0" Then
        SetStringToBoolean = "0"
    Else
        SetStringToBoolean = "1"
    End If
    
End Function

Public Function SetBooleanToString(vBooleanValue As Variant) As String
    On Error Resume Next
    
    If vBooleanValue = "0" Or vBooleanValue = 0 Then
        SetBooleanToString = "0"
    Else
        SetBooleanToString = "-1"
    End If
    
End Function

Public Function ReturnInstrString(strFullString, varChr As Variant, intIteration As Integer, _
                    Optional bolResetPlaceHolder As Boolean = False) As String
    ' returns (parses) string based on iteration and delimiter string
    ' ex:
    '       If there is a string which will contain 3 separate values, this function will
    '       be called 3 times.
    '       Example string:
    '           \\pedcqinf\team;WE\xxalurt;xxalurt (; is the delimiter)
    '               1st time through will return (intIteration = 1) -> "\\pedcqinf\team"
    '               2nd time will return (intIteration = 2) -> "WE\xxalurt"
    '               3rd time will return (intIteration = 3) -> "xxalurt"
    '
    ' Parameter(s):
    '       StrFullString (String) - entire string to select needed section
    '           ex: \\pedcqinf\team;WE\xxalurt;xxalurt
    '       varChr (String) - delimiter (in this app - the semi-colon)
    '       intIteration (Integer) - what number of times have we been stripping the
    '           same string (different processes dependent on how many times we have
    '           come into this function with the same string to parse).
    '       bolResetPlaeceHolder (Boolan) - 9/11/2001 MJN : added optinal boolean
    '           to allow calling procedure to automatically reset the intPlaceHold
    '           variable back to 0  .We will want to do this if this function is called
    '           within a loop and sometimes the loop will skip it's iteration through
    '           this function - via code - although the interations will be set to 0
    '           through code via the loop code that calls this functions, we need to
    '           sometimes reset the intPlaceholder (where we begin - character number -
    '           to check a string)
    ' Returns:
    '       Parsed string
    '
    Dim i As Integer
    Dim strTemp As String
    Static strHoldTemp As String
    Static intPlaceHold As Integer
    
    On Error GoTo err_ReturnInstrString
    
    ' caller manually needs to reset intPlaceHold
    If bolResetPlaceHolder = True Then intPlaceHold = 0
    ' if this is the 1st time through select the string up to
    ' (and NOT including) the string delimeter.
    If intIteration = 1 And intPlaceHold = 0 Then
        strTemp = Trim(Left$(strFullString, InStr(strFullString, varChr) - 1))
        ' out next starting point (for the same string)
        ' for the 2nd iteration
        intPlaceHold = InStr(strFullString, varChr) + 1
        ' actual string we will check on the 2nd time through
        strHoldTemp = Trim(Mid$(strFullString, intPlaceHold))
        intPlaceHold = 0
    Else
        ' this is for any iteration other than the last one for
        ' the original strFullString string
        If InStr(strHoldTemp, varChr) <> 0 Then
            ' the required parsed string
            strTemp = Trim(Mid$(strHoldTemp, 1, InStr(strHoldTemp, varChr) - 1))
            ' out next starting point (for the same string)
            ' for 3rd or above iteration
            intPlaceHold = InStr(strHoldTemp, varChr) + 1
            ' the actual string we will check the next iteration
            ' for the same strFullString value
            strHoldTemp = Trim(Mid$(strHoldTemp, intPlaceHold))
        Else
            ' last iteration to parse the end of the string
            strTemp = Trim(strHoldTemp)
            ' reset place holder for next (new) string
            intPlaceHold = 0
            ' reset temp string for next (new) string
            strHoldTemp = ""
        End If
    End If
    
    ' returned the parsed string
    ReturnInstrString = Trim(strTemp)
    
    Exit Function
    
err_ReturnInstrString:
    With Err
        .Raise .Number, "UtilitiesMod - ReturnInstrString", .Description
    End With
    Exit Function
    
End Function

Public Function GetErrorString(ByVal lngLastErrorValue As Long) As String
    ' Format API error messages
    '
    ' Parameter(s):
    '       LastErrorValue (Long) - error number to get formatted description
    '
    ' Returns:
    '       API error description
    '
    Dim lngBytes As Long
    Dim strDesc As String
    
    On Error Resume Next
    
    ' pad/fill description variable
    strDesc = String$(255, 0)
    
    ' get API error message
    lngBytes = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, _
        lngLastErrorValue, 0, strDesc, 128, 0)
        
    If lngBytes > 0 Then
        ' on success, return error description
        GetErrorString = Left$(strDesc, lngBytes)
    End If

End Function

Public Function IsCompiled() As Boolean
   ' Determine if running from EXE/IDE.
   On Error Resume Next
   Debug.Print 1 / 0
   IsCompiled = (Err.Number = 0)
End Function

Public Function StripNulls(strValueWithNulls As String) As String
    ' Strip nulls from parameter string
    '
    ' Parameter(s):
    '       strValueWithNulls (String) - string variable (with possible trailing
    '           null characters.
    '
    ' Returns:
    '       strValueWithNulls without trailing nulls
    '
    StripNulls = Left$(strValueWithNulls, InStr(strValueWithNulls, Chr$(0)) - 1)
    
End Function

Public Function AddBackSlash(strData As String) As String
    ' Add a backslash where needed (usually at the end of
    ' a network or local path string)
    '
    ' Parameter(s):
    '       strData (String) - string that may need a backslash
    '           concatenated with it.
    '
    ' Returns:
    '       String with a backslash as the last character.
    '
    If Right$(strData, 1) = "\" Then
        ' no backslash needed
        AddBackSlash = strData
    Else
        ' add backslash to string
        AddBackSlash = strData & "\"
    End If
    
End Function

Public Function GetCorrectPath() As String
    ' returns application path with a backslash
    '
    ' Returns:
    '       App path with backslash
    '
    Dim strpath As String
    Dim strBuffer As String
    Dim lngFilePart As Long
    Dim lngReturn As Long
    Dim strAppName As String
        
    strpath = App.Path
    
    ' check if app.Path returns the UNC path
    If Left$(strpath, 2) = "\\" Then
        If IsCompiled Then
            strAppName = App.EXEName & ".exe"
        Else
            strAppName = App.EXEName & ".vbp"
        End If
        ' try to get the network path of the application file (also includes
        ' application file name)
        strBuffer = Space$(MAX_PATH)
        lngReturn = GetFullPathName(strAppName, Len(strBuffer), strBuffer, lngFilePart)
        If lngReturn Then
            ' get rid of application name from strpath
            strpath = Left$(strBuffer, lngReturn - Len(strAppName))
        End If
    End If
    
    If Right$(strpath, 1) <> "\" Then
        strpath = strpath & "\"
    End If
    
    GetCorrectPath = strpath
    
End Function

Public Function CheckForApostrophe(strName As String) As String
    Dim strTemp As String
    Dim lngPosition As Long
    
    ' we need to add an extra apostrophe to any variable which contains an apostophe
    lngPosition = InStr(strName, "'")
    If lngPosition = 0 Then
        strTemp = strName
    Else
        strTemp = Mid$(strName, 1, lngPosition - 1) & "'" & Mid$(strName, lngPosition)
    End If
    
    CheckForApostrophe = strTemp
    
End Function

Public Function ReturnNoApostrophe(strName As String) As String
    Dim strTemp As String
    Dim lngPosition As Long
       
    ' we need to get rid of any apostrophes from variable
    lngPosition = InStr(strName, "'")
    
    If lngPosition = 0 Then
        strTemp = strName
    Else
        strTemp = Replace(strName, "'", "")
        If InStr(strName, "'") > 0 Then
            ' recursive routine
            ReturnNoApostrophe strTemp
        End If
    End If
    
    ReturnNoApostrophe = strTemp
    
End Function

Public Function ReturnDoubleApostrophe(strName As String) As String
    Dim strTemp As String
    Dim lngPosition As Long
       
    ' we need to get rid of any apostrophes from variable
    lngPosition = InStr(strName, "'")
    
    If lngPosition = 0 Then
        strTemp = strName
    Else
        strTemp = Replace(strName, "'", "''")
    End If
    
    ReturnDoubleApostrophe = strTemp
    
End Function

Public Sub CenterForm(objForm As Form, Optional objMDI As MDIForm)
    
    If objMDI Is Nothing Then
        ' move form to middle of screen
        With objForm
            .Move (Screen.Width - .Width) / 2, (Screen.Height - .Height) / 2
        End With
    Else
        With objForm
            .Move (objMDI.Width - .Width) / 2, (objMDI.Height - .Height) / 2
        End With
    End If
    
End Sub

Public Sub CenterControl(objCtl As Control, objForm As Form, Optional BolVerticalMove As Boolean = True)
    On Error Resume Next
    
    ' move control to middle of screen
    With objCtl
        If BolVerticalMove Then
            .Move (objForm.Width - .Width) / 2, (objForm.Height - .Height) / 2
        Else
            ' in case we want to center the control to the forms' width
            ' and NOT height
            .Move (objForm.Width - .Width) / 2, .Top
        End If
    End With
    
End Sub

Public Sub CenterFrameControl(objCtl As Control, objFrame As Control, Optional BolVerticalMove As Boolean = True)
    On Error Resume Next
    
    ' move control to middle of frame
    With objCtl
        If BolVerticalMove Then
            .Move (objFrame.Width - .Width) / 2, (objFrame.Height - .Height) / 2
        Else
            ' in case we want to center the control to the frames' width
            ' and NOT height
            .Move (objFrame.Width - .Width) / 2, .Top
        End If
    End With
    
End Sub

Public Function IsFormLoaded(strForm As String) As Boolean
    Dim frm As Form
    
    IsFormLoaded = False
    For Each frm In Forms
        If frm.Name = strForm Then
            IsFormLoaded = True
            Exit For
        End If
    Next
    
End Function

Public Function SelectAllText(objControl As Object)

    With objControl
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
        
End Function

Public Function ValidTest(intKeyIn As Integer, strValidate As String, bolEditable As Boolean, Optional bolSpaceBar As Boolean = False) As Integer
    ' Validates data entered into a control
    ' Written by Juan Lozano
    '
    ' Called in a controls KeyPress event:
    '               KeyAscii = ValidTest(KeyAscii, "0123456789", True, [True])
    '
    ' Arguments:
    '               intKeyIn        - Key pressed
    '               strValidate     - Allowed characters
    '               boleditable     - if the backspace key is allowed
    '               bolSpaceBar     - if the space bar is allowed
    ' Returns:
    '               0    Invalid key is pressed
    '               Any other return means a valid key is pressed
    '
    Dim strValidateList As String
    Dim intKeyOut As Integer
    
    
    If bolEditable Then
        strValidateList = UCase(strValidate) & Chr$(8)
    Else
        strValidateList = UCase(strValidate)
    End If
    
    ' skip this code section if the backspace key was hit (intKey = 8)
    If intKeyIn <> 8 Then
        If bolSpaceBar Then
            strValidateList = UCase(strValidate) & Chr$(32)
        Else
            strValidateList = UCase(strValidate)
        End If
    End If
    
    If InStr(1, strValidateList, UCase(Chr$(intKeyIn)), 1) > 0 Then
        intKeyOut = intKeyIn
    Else
        intKeyOut = 0
    End If
    
    ValidTest = intKeyOut
    
End Function

Public Function AddApos(ReplaceString As String) As String

    Dim lngCount As Long
    Dim lngLen As Long
    
    On Error GoTo error_addapos
    
    If InStr(1, ReplaceString, "'") > 0 Then
        lngLen = Len(ReplaceString)
        For lngCount = 1 To lngLen
            If Mid$(ReplaceString, lngCount, 1) = "'" Then
                AddApos = AddApos & "'" & "'"
            Else
                AddApos = AddApos & Mid$(ReplaceString, lngCount, 1)
            End If
        Next lngCount
    Else
        AddApos = ReplaceString
    End If
    
    Exit Function
    
error_addapos:
    
    '  Just return the passed-in string.
    AddApos = ReplaceString

End Function

Public Function RemoveLeadingZeros(strNumber As String, Optional intHowManyZeros As Integer = 0) As String
    ' strip leading 0's off a string number
    Dim lngTempNumber As Long
    
    On Error GoTo err_RemoveLeadingZeros
    
    If Len(Trim(strNumber)) = 0 Then Exit Function
    
    ' if we know how many leading 0's we have
    If intHowManyZeros > 0 Then
        ' add 1 because we start AFTER the last 0
        intHowManyZeros = intHowManyZeros + 1
        RemoveLeadingZeros = Mid(strNumber, intHowManyZeros)
        Exit Function
    Else
        ' we don't know how many 0's we have
        ' make sure variable is numeric (contains all numbers)
        If IsNumeric(strNumber) Then
            ' cast to a long variable (gets rid of leading 0's)
            lngTempNumber = CLng(strNumber)
            ' recast to string
            RemoveLeadingZeros = CStr(lngTempNumber)
        End If
    End If
        
    Exit Function
    
err_RemoveLeadingZeros:
    ' just return passed-in string
    RemoveLeadingZeros = strNumber
        
End Function

Public Sub PrintTextInPicturebox(objPic As PictureBox, strText As String, _
    intTextIterations As Integer, Optional lngFontSize As Long = 10, _
    Optional lngPictureBackColor As Long = &H808080, _
    Optional lngFontDepthColor As Long = 0, _
    Optional lngFontForeColor As Long = &HFFFF&)
    
    Dim lngPicTextHeight As Long
    Dim lngPicHeight As Long
    Dim lngPicWidth As Long
    Dim X As Long
    Dim Y As Long
    Dim i As Integer
    
    With objPic
        .AutoRedraw = True
        .FontSize = lngFontSize
        .FontBold = True
        .BackColor = lngPictureBackColor
        .ScaleMode = vbPixels
        lngPicTextHeight = .TextHeight("X")
        lngPicHeight = .ScaleHeight
        lngPicWidth = .ScaleWidth
    
        X = BitBlt(.hDC, 0, -lngPicTextHeight, lngPicWidth, lngPicHeight, .hDC, 0, 0, SRCCOPY)
        objPic.Line (0, lngPicHeight - lngPicTextHeight)-(lngPicWidth, lngPicHeight), .BackColor, BF
        .CurrentY = lngPicHeight - lngPicTextHeight
    
        .CurrentX = (lngPicWidth / 2) - (.TextWidth(strText) / 2)
        .ForeColor = lngFontDepthColor
        X = .CurrentX
        Y = .CurrentY

        For i = 1 To intTextIterations
            objPic.Print strText
            X = X + 1
            Y = Y + 1
            .CurrentX = X
            .CurrentY = Y
        Next i
        .ForeColor = lngFontForeColor
        objPic.Print strText
    End With
    
End Sub

Public Function CheckForPreviousInstance() As Boolean
    ' this will bring up the previous instance of this application to
    ' the front of the desktop.
    ' This sub is only called when a user tries to start multiple
    ' instances of this application (only 1 allowed at a time).
    Dim lng_hWnd As Long
    Dim typWin As WINDOWPLACEMENT
    Dim strAppTitle As String
    
    strAppTitle = App.Title
    
    ' change this app title so we don't find it when looking
    ' for a previous instance
    App.Title = "#$#"
    
    ' try to find a previous instance (looking for the original
    ' app title)
    lng_hWnd = FindWindow("ThunderRT6MDIForm", strAppTitle)

    ' previous instance IS found
    If lng_hWnd <> 0 Then
        typWin.Length = Len(typWin)
        ' get the state of the window (minimized or not)
        GetWindowPlacement lng_hWnd, typWin
        If typWin.showCmd = SW_SHOWMINIMIZED Then
            ' if minimized, display the window in it's normal mode
            typWin.showCmd = SW_SHOWNORMAL
            SetWindowPlacement lng_hWnd, typWin
        End If
        ' bring the application window to the front
        SetForegroundWindow lng_hWnd
    Else
        ' reset this apps title if a previous instance
        ' was NOT found
        App.Title = strAppTitle
    End If
    
End Function

Public Function ShellAndWait(ByVal strJobToDo As String, Optional lngExecMode, Optional lngTimeOut) As Long
    ' Shells a new process and waits for it to complete.
    ' Calling application is totally non-responsive while
    ' new process executes.
    '
    ' Parameter(s):
    '       strJobToDo (String) - application to run
    '       lngExecMode (Long) - window state of application while running
    '       lngTimeOut (Long) - timeout period for app
    '
    ' Function/Sub calls:
    '       Shell (API call)
    '       OpenProcess (API call)
    '       WaitForSingleObject (API call)
    '       CloseHandle (API call)
    '
    ' Returns:
    '       0 (success) / any non 0 number (error number)
    '
    Dim lngProcessID As Long
    Dim lngProcess As Long
    Dim lngReturn As Long
    Const fdwAccess = SYNCHRONIZE

    On Error GoTo err_ShellAndWait
    
    ' default app to a minimized/no focus window state
    If IsMissing(lngExecMode) Then
        lngExecMode = vbMinimizedNoFocus
    Else
        If lngExecMode < vbHide Or lngExecMode > vbMinimizedNoFocus Then
             lngExecMode = vbMinimizedNoFocus
        End If
    End If
    
    ' shell out to the called application (strJobToDo)
    lngProcessID = Shell(strJobToDo, CLng(lngExecMode))
    If Err Then
        'err shelling out to the called program
        ShellAndWait = lngProcessID
        Exit Function
    End If
   
    ' default timeout to Infinity
    If IsMissing(lngTimeOut) Then
        lngTimeOut = INFINITE
    End If
    
    ' open process and return the handle of the called program
    lngProcess = OpenProcess(fdwAccess, False, lngProcessID)
    ' wait for app process to run
    lngReturn = WaitForSingleObject(lngProcess, CLng(lngTimeOut))
    ' close open process (handle)
    Call CloseHandle(lngProcess)

    ' return success or failure of application run call
    Select Case lngReturn
        Case WAIT_OBJECT_0
            ' do nothing - app call was successful
            
        Case WAIT_TIMEOUT
            ' timeout period elapsed
            Err.Raise vbObjectError + lngReturn, , strJobToDo & " Timed out!  " & _
                GetErrorString(lngReturn)

        Case WAIT_ABANDONED
            ' timeout period was abandoned
            Err.Raise vbObjectError + lngReturn, , "Wait for " & strJobToDo & " was abandoned!  " & _
                GetErrorString(lngReturn)

        Case WAIT_FAILED
            ' timeout period failed
            Err.Raise vbObjectError + lngReturn, , "The wait for " & strJobToDo & " failed!  " & _
                GetErrorString(lngReturn)
        
        Case Else
            ' unknown error
            Err.Raise Err.LastDllError, , "Unknown process error calling " & strJobToDo & "!"
            
    End Select
   
    ' return success (0) or failure (any non-zero number)
    ShellAndWait = lngReturn
   
    Exit Function
   
err_ShellAndWait:
    With Err
        .Raise .Number, .Source & vbCrLf & "Function: ShellAndWait", .Description
    End With
    Exit Function
   
End Function

Public Function ShellAndLoop(ByVal strJobToDo As String, Optional lngExecMode) As Long
    '
    ' Shells a new process and waits for it to complete.
    ' Calling application is responsive while new process
    ' executes. It will react to new events, though execution
    ' of the current thread will not continue.
    '
    Dim lngProcessID As Long
    Dim lngHProcess As Long
    Dim lngRet As Long
       
    
    If IsMissing(lngExecMode) Then
        lngExecMode = vbNormalFocus
    Else
        If lngExecMode < vbHide Or lngExecMode > vbMinimizedNoFocus Then
            lngExecMode = vbNormalFocus
        End If
    End If
   
    lngProcessID = Shell(strJobToDo, lngExecMode)
    lngHProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, lngProcessID)
    
    Do
        GetExitCodeProcess lngHProcess, lngRet
        Sleep 0
    Loop While (lngRet = STILL_ACTIVE)
    CloseHandle lngHProcess
   
    ShellAndLoop = lngRet
    
'   On Error Resume Next
'      ProcessID = Shell(JobToDo, CLng(ExecMode))
'      If Err Then
'         ShellAndLoop = vbObjectError + Err.Number
'         Exit Function
'      End If
'   On Error GoTo 0
'
'   hProcess = OpenProcess(fdwAccess, False, ProcessID)
'   Do
'      GetExitCodeProcess hProcess, nRet
'      DoEvents
'      sleep 100
'   Loop While nRet = STILL_ACTIVE
'   Call CloseHandle(hProcess)
'

End Function

Public Function CheckValidArray(varArray As Variant) As Boolean
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

Public Function BinarySearch(vntData As Variant, ByVal sItem As String, _
    Optional ByVal lSearchCol As Long = 1) As Long
    Dim lLevel As Long
    Dim lb As Long, ub As Long, X As Long
    Dim res As Long
    ' Taken from PlanetSOurce Code web site
    ' http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=37185&lngWId=1
    '    by Josip Habjan
    '       Habjan Software
    '       softdevelopers@go.com
    '
    
    
    ' lLevel shows you how many iterations
    lLevel = 0
    
    BinarySearch = -1

    If Not IsEmpty(vntData) Then 'in Case if vntData is empty
        ub = UBound(vntData, 2) 'get number of rows
        lb = 0

        Do
            lLevel = lLevel + 1
            X = (lb + ub) \ 2 'calculating
            res = StrComp(sItem, vntData(lSearchCol, X), vbTextCompare) 'compare

            If res = 0 Then 'case Item FIND
                BinarySearch = X 'return row index
                Exit Do
            ElseIf res = -1 Then 'check how near are till ITEM
                ub = X - 1
            Else
                lb = X + 1
            End If
        Loop While ub >= lb
    End If

End Function

Public Sub DisabledTextChange(bolUseDefault As Boolean)
    
    If bolUseDefault Then
        SetSysColors 1, COLOR_GREYTEXT, DEFAULT_COLOR
    Else
        SetSysColors 1, COLOR_GREYTEXT, 0
    End If
End Sub

Public Sub LockWindow(hwnd As Long)
    ' if 0, the calling window/control is unlocked
    LockWindowUpdate hwnd
End Sub
