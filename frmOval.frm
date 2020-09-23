VERSION 5.00
Begin VB.Form frmOval 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Data Class Builder"
   ClientHeight    =   5745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   690
      Top             =   5235
   End
   Begin VB.PictureBox picGreeting 
      BackColor       =   &H00000000&
      Height          =   5100
      Left            =   0
      ScaleHeight     =   5040
      ScaleWidth      =   6825
      TabIndex        =   1
      Top             =   60
      Width           =   6885
      Begin VB.PictureBox picIcon 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   2250
         Left            =   135
         Picture         =   "frmOval.frx":0000
         ScaleHeight     =   2250
         ScaleWidth      =   2250
         TabIndex        =   2
         Top             =   165
         Width           =   2250
      End
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00000000&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2850
      MaskColor       =   &H00000000&
      TabIndex        =   0
      Top             =   5265
      Width           =   1335
   End
End
Attribute VB_Name = "frmOval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, _
        ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Declare Function CreateRoundRectRgn Lib "gdi32" _
        (ByVal X1 As Long, ByVal Y1 As Long, _
        ByVal X2 As Long, ByVal Y2 As Long, _
        ByVal X3 As Long, ByVal Y3 As Long) As Long

Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, _
        ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, _
        ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
        
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
        lParam As Any) As Long


Private Declare Function DeleteObject Lib "gdi32" _
        (ByVal hObject As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Dim xp As Long
Dim yp As Long

Private m_lngChildCmdRegion As Long
Private m_lngChildPicRegion As Long
Private m_lngChildFormRegion As Long

Private m_lngCmdWidth As Long
Private m_lngCmdHeight As Long
Private m_lngPicWidth As Long
Private m_lngPicHeight As Long
Private m_lngFormWidth As Long
Private m_lngFormHeight As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub cmdClose_Click()
    Me.Visible = False
    Unload Me
End Sub

Private Sub Form_Activate()
    DisplayMessage
End Sub

Private Sub DisplayMessage()
    Dim strMessage(9) As String
    Dim X As Long
    Dim Y As Long
    Dim i As Integer
    Dim j As Integer
    
    strMessage(0) = "Hello There!"
    strMessage(1) = ""
    strMessage(2) = vbTab & "Today is: " & Format(Date, "Long Date")
    strMessage(3) = ""
    strMessage(4) = vbTab & "Application: Data Class Builder"
    strMessage(5) = ""
    strMessage(6) = vbTab & "Developed By:"
    strMessage(7) = vbTab & vbTab & "Michael J. Nugent"
    strMessage(8) = vbTab & vbTab & "wiscmike@yahoo.com"
    strMessage(9) = ""
    
    With picGreeting
        .AutoRedraw = True
        .FontSize = 12
        .FontBold = True
        .BackColor = &H0
        .ScaleMode = vbPixels
        .CurrentX = (.ScaleWidth / 2) - (.TextWidth(strMessage(0)) / 2)

        For j = 0 To UBound(strMessage) ' - 1
            X = .CurrentX
            Y = .CurrentY
            .ForeColor = &HFF0000
            For i = 1 To 3
                picGreeting.Print strMessage(j)
                X = X + 1
                Y = Y + 1
                .CurrentX = X
                .CurrentY = Y
            Next i
            .ForeColor = &HFF00FF   '&HFFFF&
            picGreeting.Print strMessage(j)
        Next j
    End With
End Sub

Private Sub Form_Load()
        
    CenterForm Me
    xp = Screen.TwipsPerPixelX
    yp = Screen.TwipsPerPixelY
    m_lngCmdWidth = cmdClose.Width
    m_lngCmdHeight = cmdClose.Height
    m_lngPicWidth = picGreeting.Width
    m_lngPicHeight = picGreeting.Height
    m_lngFormWidth = Me.Width
    m_lngFormHeight = Me.Height
   
    CreateFormShape
End Sub

Private Sub CreateFormShape()
        
    '  Respectively (x,y) upperleft, (x,y) lowerright, ellipse height, ellipsewidth
    m_lngChildFormRegion = CreateRoundRectRgn(0, 0, m_lngFormWidth / xp, m_lngFormHeight / yp, 340, 340)
    SetWindowRgn Me.hWnd, m_lngChildFormRegion, False
    '  Respectively (x,y) upperleft, (x,y) lowerright, ellipse height, ellipsewidth
    m_lngChildCmdRegion = CreateRoundRectRgn(0, 0, m_lngCmdWidth / xp, m_lngCmdHeight / yp, 240, 240)
    SetWindowRgn cmdClose.hWnd, m_lngChildCmdRegion, False
    '  Respectively (x,y) upperleft, (x,y) lowerright, ellipse height, ellipsewidth
    m_lngChildPicRegion = CreateRoundRectRgn(0, 0, m_lngPicWidth / xp, m_lngPicHeight / yp, 540, 540)
    SetWindowRgn picGreeting.hWnd, m_lngChildPicRegion, False
    
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then
         Exit Sub
    End If
    ReleaseCapture
    SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SetWindowRgn cmdClose.hWnd, 0, False
    DeleteObject m_lngChildCmdRegion
    SetWindowRgn picGreeting.hWnd, 0, False
    DeleteObject m_lngChildPicRegion
    SetWindowRgn Me.hWnd, 0, False
    DeleteObject m_lngChildFormRegion
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmOval = Nothing
End Sub

Private Sub picGreeting_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        ReleaseCapture
        SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
    
End Sub

Private Sub Timer1_Timer()
     Static intStart As Integer
    Static VelX As Long
    Static VelY As Long
    Static MaxX As Long
    Static MaxY As Long
    Static X As Long
    Static Y As Long
    
    If intStart = 0 Then
        picGreeting.ScaleMode = 3 'Pixels
        picIcon.ScaleMode = 3
        VelX = 3 'Horizontal and vertical velocities
        VelY = 3 'Measured In pixels
        MaxX = picGreeting.ScaleWidth - picIcon.ScaleWidth 'Max X and Y
        MaxY = picGreeting.ScaleHeight - picIcon.ScaleHeight
        X = 0
        Y = 0
        intStart = 1
    End If

    'This is the main loop
    X = X + VelX ' Increase X and Y coordinates
    Y = Y + VelY


    If X <= 0 Or X >= MaxX Then ' If X is out of bounds,
        VelX = -VelX ' invert the horizontal
    End If ' velocity In order
    ' to make it move back


    If Y <= 0 Or Y >= MaxY Then ' The same thing With the
        VelY = -VelY ' vertical (Y) coordinate
    End If
    
    picIcon.Move X, Y ' Move the picture To the
    ' calculated X and

    
End Sub
