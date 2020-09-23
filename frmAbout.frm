VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Telecom Circuits Viewer"
   ClientHeight    =   4200
   ClientLeft      =   2385
   ClientTop       =   3525
   ClientWidth     =   8625
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2898.916
   ScaleMode       =   0  'User
   ScaleWidth      =   8099.321
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Left            =   285
      Top             =   3465
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   7215
      TabIndex        =   6
      Top             =   3630
      Width           =   1185
   End
   Begin VB.PictureBox Picture2 
      Height          =   540
      Left            =   6525
      Picture         =   "frmAbout.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   3390
      Width           =   540
   End
   Begin VB.PictureBox Picture1 
      Height          =   540
      Left            =   1485
      Picture         =   "frmAbout.frx":05A5
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   3390
      Width           =   540
   End
   Begin VB.PictureBox picAppTitle 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   1125
      ScaleHeight     =   1440
      ScaleWidth      =   6225
      TabIndex        =   2
      Top             =   1845
      Width           =   6285
      Begin VB.PictureBox picLogoTitle 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   -15
         ScaleHeight     =   705
         ScaleWidth      =   6240
         TabIndex        =   3
         Top             =   -15
         Width           =   6240
      End
   End
   Begin VB.PictureBox picCredits 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   3750
      ScaleHeight     =   1440
      ScaleWidth      =   4620
      TabIndex        =   1
      Top             =   45
      Width           =   4680
   End
   Begin VB.PictureBox picLaserShow 
      BackColor       =   &H8000000D&
      Height          =   1500
      Left            =   90
      Picture         =   "frmAbout.frx":0708
      ScaleHeight     =   1440
      ScaleMode       =   0  'User
      ScaleWidth      =   3585
      TabIndex        =   0
      Top             =   45
      Visible         =   0   'False
      Width           =   3645
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   270
      Top             =   1635
   End
   Begin VB.Line Line3 
      X1              =   6113.227
      X2              =   4197.561
      Y1              =   2505.492
      Y2              =   2505.492
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   345
      Left            =   4005
      Top             =   3480
      Width           =   480
   End
   Begin VB.Line Line2 
      X1              =   3760.902
      X2              =   1845.237
      Y1              =   2505.492
      Y2              =   2505.492
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   1352.234
      X2              =   6577.118
      Y1              =   1159.566
      Y2              =   1159.566
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   1366.32
      X2              =   6577.118
      Y1              =   1169.92
      Y2              =   1169.92
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, _
        ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
        ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
        ByVal dwRop As Long) As Long

Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, _
        ByVal Y As Long, ByVal crColor As Long) As Long

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, _
        ByVal X As Long, ByVal Y As Long) As Long

Private Const SRCCOPY = &HCC0020
Private Const SRCINVERT = &H660046
Private Const SRCPAINT = &HEE0086
Private Const SRCERASE = &H440328

Private m_lngHeight As Long
Private m_intPicLaserPointY As Integer
Private m_intPicLaserPointX  As Integer
Private m_lngPicPointColor As String

Private m_lngPicTextHeight As Long
Private m_lngPicHeight As Long
Private m_lngPicWidth As Long

Private m_lngPicIconHeight As Long
Private m_lngPicIconWidth As Long

Private m_strCredits(12) As String

Private m_lngMsgLeft As Long
Private Const cnUPPERARRAY = 12

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Function PicureboxLaser(objPicturebox As PictureBox, intLineStartX As Integer, intLineStartY As Integer, intLeftStartPos As Integer, intTopStartPos As Integer, lngLineColor As Long)
    Dim lngScaleWidth As Long
    Dim lngScaleHeight As Long
    Static bolOnceThrough As Boolean
    
    With objPicturebox
        .ScaleMode = vbPixels
        .AutoRedraw = True
        lngScaleWidth = .ScaleWidth
        lngScaleHeight = .ScaleHeight
    End With
    
    For m_intPicLaserPointX = 0 To lngScaleWidth
        Sleep 0
    
        For m_intPicLaserPointY = 0 To lngScaleHeight
            m_lngPicPointColor = objPicturebox.Point(m_intPicLaserPointX, m_intPicLaserPointY)
            Line (intLineStartX, intLineStartY)-(intLeftStartPos + m_intPicLaserPointX, intTopStartPos + m_intPicLaserPointY), m_lngPicPointColor
        Next
    
        Line (intLineStartX, intLineStartY)-(intLeftStartPos + m_intPicLaserPointX, intTopStartPos + lngScaleHeight), lngLineColor
    Next
    
    For m_intPicLaserPointX = 0 To lngScaleHeight
       Line (intLineStartX, intLineStartY)-(intLeftStartPos + lngScaleWidth, intTopStartPos + m_intPicLaserPointX), lngLineColor
    Next
    If Not bolOnceThrough Then
        SetFormLoadDisplay
    End If
    bolOnceThrough = True
    
End Function

Private Sub SetFormLoadDisplay()
    Dim i As Integer
    Dim X As Integer
    Dim Y As Integer
    
    Dim strAppTitle As String
    
    Me.Height = m_lngHeight
    picLaserShow.Visible = True
    
    strAppTitle = "Data"
    PrintTextInPicturebox picLogoTitle, strAppTitle, 4, 30, &H404040, 0, vbWhite
    
    strAppTitle = "Class Builder"
    PrintTextInPicturebox picAppTitle, strAppTitle, 4, 26, &H808080, 0, vbWhite
    
End Sub

Private Sub Form_Activate()
    PaintLaser
    StartTimingFunctions
End Sub

Private Sub PaintLaser()
    Dim intX As Integer
    Dim intY As Integer
    
    With picLaserShow
        intX = .Top - 2
        intY = .Left - 2
    End With
    PicureboxLaser picLaserShow, intX, intY, 10, 10, Me.BackColor
    
End Sub
Private Sub StartTimingFunctions()
    
    CreateCreditsMessageArray
    
    With picCredits
        .AutoRedraw = True
        .FontSize = 10
        .FontBold = True
        .BackColor = &H808080
        .ScaleMode = vbPixels
        m_lngPicTextHeight = .TextHeight("X")
        m_lngPicHeight = .ScaleHeight
        m_lngPicWidth = .ScaleWidth
    End With
       
    Timer1.Enabled = True
    Timer1.Interval = 500
    
    Timer2.Enabled = True
    Timer2.Interval = 500
End Sub

Private Sub Form_Load()
   Dim lngPicHeight As Long
      
   CenterForm Me
   With Me
        m_lngHeight = .Height
        lngPicHeight = picLaserShow.Height
        .ScaleMode = vbPixels
        .Height = lngPicHeight + 1000
    End With
  
End Sub

Private Sub CreateCreditsMessageArray()
    
    m_strCredits(0) = "Data Class Builder"
    m_strCredits(1) = "Version: " & App.Major & "." & App.Minor
    m_strCredits(2) = ""
    m_strCredits(3) = "© Copyright 2002 "
    m_strCredits(4) = "All Rights Are Reserved"
    m_strCredits(5) = "Unauthorized use is strictly prohibited:-)"
    m_strCredits(6) = ""
    m_strCredits(7) = "Application Developer: Michael J. Nugent"
    m_strCredits(8) = "email: wiscmike@yahoo.com"
    m_strCredits(9) = "Application Date: December 2002"
    m_strCredits(10) = ""
    m_strCredits(11) = "Today is: " & CStr(Format(Date, "Long date"))
    m_strCredits(12) = ""
    
End Sub

Private Sub PrintText(Text As String)
    Dim X As Long
    Dim Y As Long
    Dim i As Integer
    
    With picCredits
        X = BitBlt(.hdc, 0, -m_lngPicTextHeight, m_lngPicWidth, m_lngPicHeight, .hdc, 0, 0, SRCCOPY)
        picCredits.Line (0, .ScaleHeight - m_lngPicTextHeight)-(m_lngPicWidth, m_lngPicHeight), .BackColor, BF
        .CurrentY = m_lngPicHeight - m_lngPicTextHeight - 5
    
        .CurrentX = (.ScaleWidth / 2) - (.TextWidth(Text) / 2)
        .ForeColor = 0
        X = .CurrentX
        Y = .CurrentY

        For i = 1 To 3
            picCredits.Print Text
            X = X + 1
            Y = Y + 1
            .CurrentX = X
            .CurrentY = Y
        Next i
        .ForeColor = &HFFFF&    '&HFFC0C0
        picCredits.Print Text
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Set frmAbout = Nothing
End Sub

Private Sub picCredits_Click()
    Timer1.Enabled = Not Timer1.Enabled
End Sub

Private Sub picLaserShow_DblClick()
    PaintLaser
End Sub

Private Sub Picture1_DblClick()
    frmOval.Show vbModal
End Sub

Private Sub Timer1_Timer()
    Static i As Integer

    If i <= cnUPPERARRAY Then
        PrintText m_strCredits(i)
    End If
    i = i + 1
    If i > cnUPPERARRAY Then
        i = 0
    End If

End Sub

Private Sub Timer2_Timer()
    
    If Line2.BorderStyle = 2 Then
        Line2.BorderStyle = 1
        Line3.BorderStyle = 2
        Shape1.FillColor = vbRed
    Else
        Line2.BorderStyle = 2
        Line3.BorderStyle = 1
        Shape1.FillColor = vbBlack
    End If
        
End Sub
