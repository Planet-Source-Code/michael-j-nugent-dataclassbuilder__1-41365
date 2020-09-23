Attribute VB_Name = "MainMod"
'===============================================================
' Module:   MainMod
'
' Project:
'
' Purpose:  Application startup module
'
' Author:   Michael J. Nugent
'           wiscmike@yahoo.com
'
' Date:
'
'===============================================================
Option Explicit

' application error log file name
Private Const cnERRLOGFILE = "ClassBuilderErr.log"

Private m_objError As clsError
Private m_objProcess As clsProcess

Public Sub Main()

    On Error GoTo err_Main
    Dim lblAppLoad As Label
    
    Set m_objError = New clsError
    With m_objError
        .LogFileName = cnERRLOGFILE
        .DisplayErrMsg = True
    End With
    
    If App.PrevInstance Then
        CheckForPreviousInstance
        EndProgram
        Exit Sub
    End If
    
    Load frmSplash
            
    With frmClassBuilder
        Set .ObjError = m_objError
        Set .ObjAppProcess = m_objProcess
        .Show
    End With
        
    Unload frmSplash
    Exit Sub
    
err_Main:
    With Err
        m_objError.UpdateLogFile "MainMod", "Main", .Number, .Description, .Source
    End With
    EndProgram
    
End Sub

Public Sub EndProgram()
    Dim objForm As Form
    
    On Error Resume Next
    
    For Each objForm In Forms
        Unload objForm
    Next
        
    Set m_objError = Nothing
    Set m_objProcess = Nothing
    
End Sub


