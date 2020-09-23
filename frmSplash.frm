VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3030
   ClientLeft      =   3465
   ClientTop       =   4095
   ClientWidth     =   6885
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picLogo 
      BackColor       =   &H8000000B&
      Height          =   975
      Left            =   1245
      ScaleHeight     =   915
      ScaleWidth      =   4305
      TabIndex        =   3
      Top             =   150
      Width           =   4365
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   5700
      Picture         =   "frmSplash.frx":0000
      Stretch         =   -1  'True
      Top             =   150
      Width           =   1095
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4860
      TabIndex        =   2
      Top             =   1965
      Width           =   1185
   End
   Begin VB.Label lblCopyright 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "© Copyright 2002 Michael J. Nugent All Rights Reserved"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1860
      TabIndex        =   1
      Top             =   2520
      Width           =   3240
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblAppTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data ClassBuilder"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1260
      TabIndex        =   0
      Top             =   1380
      Width           =   4110
   End
   Begin VB.Image imgCompany 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   60
      Picture         =   "frmSplash.frx":030A
      Stretch         =   -1  'True
      Top             =   135
      Width           =   1095
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Form_Load()
    SetFormLoad
End Sub

Private Sub SetFormLoad()
    
    On Error Resume Next
    
    Me.MousePointer = vbHourglass
    CenterForm Me
    CreateTitle_Logo
    DisplayFormProcessing
    
End Sub

Private Sub CreateTitle_Logo()
    Dim strLogo As String
    
    strLogo = "Class Builder"
    PrintTextInPicturebox picLogo, strLogo, 7, 28, &H80000004, 0, &HFF0000
End Sub

Private Sub DisplayFormProcessing()
    Dim aTitle(0 To 17) As String
    
    On Error Resume Next
    
    Dim i As Integer
    Dim j As Integer
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim lngPosition As Long
    Dim lngBound As Long
    
    On Error Resume Next
    
    CenterForm Me
    
    With lblAppTitle
        .Caption = ""
        .AutoSize = True
        CenterControl lblAppTitle, Me, BolVerticalMove:=False
    End With
    
    With Me
        .Show
        lngWidth = .Width
        lngHeight = .Height
        .Width = 0
        .Height = 0
        For i = 0 To lngWidth
            .Width = i
            If j <> lngHeight Then
                .Height = j
                j = j + 1
            End If
            CenterForm Me
            Sleep 0
        Next
        .Refresh
    End With
    
    lblVersion.Caption = lblVersion.Caption & " " & App.Major & "." & App.Minor

    aTitle(0) = "D"
    aTitle(1) = "a"
    aTitle(2) = "t"
    aTitle(3) = "a"
    aTitle(4) = " "
    aTitle(5) = "C"
    aTitle(6) = "l"
    aTitle(7) = "a"
    aTitle(8) = "s"
    aTitle(9) = "s"
    aTitle(10) = " "
    aTitle(11) = "B"
    aTitle(12) = "u"
    aTitle(13) = "i"
    aTitle(14) = "l"
    aTitle(15) = "d"
    aTitle(16) = "e"
    aTitle(17) = "r"
    
    For i = 0 To 17
        With lblAppTitle
            .Caption = .Caption & aTitle(i)
            .Visible = True
            CenterControl lblAppTitle, Me, False
        End With
        Sleep 25
    Next i
    Sleep 1000
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Me.MousePointer = vbDefault
End Sub

