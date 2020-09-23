VERSION 5.00
Begin VB.Form frmClassBuilder 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Class Builder"
   ClientHeight    =   8385
   ClientLeft      =   270
   ClientTop       =   780
   ClientWidth     =   14250
   Icon            =   "frmClassBuilder.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8385
   ScaleWidth      =   14250
   Begin VB.CheckBox chkIntegerToLong 
      Caption         =   "Convert Integer Fields To Longs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6690
      TabIndex        =   47
      Top             =   6105
      Width           =   3090
   End
   Begin VB.Frame fraPKeyInstructions 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1905
      Left            =   -2025
      TabIndex        =   41
      Top             =   -675
      Width           =   2265
      Begin VB.Label lblPKeyInstructions 
         BorderStyle     =   1  'Fixed Single
         Height          =   1905
         Left            =   45
         TabIndex        =   42
         Top             =   0
         Width           =   2205
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraPrimaryKey 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   6615
      TabIndex        =   36
      Top             =   6585
      Width           =   3285
      Begin VB.CommandButton cmdPrimaryKey 
         Caption         =   "Primary Key"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1050
         Picture         =   "frmClassBuilder.frx":030A
         TabIndex        =   43
         Top             =   1110
         Width           =   1245
      End
      Begin VB.TextBox txtPrimaryKey 
         Height          =   315
         Left            =   225
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   735
         Width           =   2820
      End
      Begin VB.Label lblPrimaryKeyHelp 
         AutoSize        =   -1  'True
         Caption         =   " ? "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3000
         TabIndex        =   40
         Top             =   1185
         Width           =   255
      End
      Begin VB.Label lblInstructions5 
         AutoSize        =   -1  'True
         Caption         =   "Click the button to select the highlighted field as the Primary Key for this table."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   330
         TabIndex        =   37
         Top             =   135
         Width           =   2670
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraDeselectTables 
      Caption         =   "   Deselect Tables   "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1920
      Left            =   3360
      TabIndex        =   30
      Top             =   6165
      Width           =   2895
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   855
         TabIndex        =   10
         Top             =   1395
         Width           =   1110
      End
      Begin VB.Label lblInstructions 
         Caption         =   "Check the table(s) you wish to remove from the class process list above, then click the &Remove button below."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   135
         TabIndex        =   31
         Top             =   315
         Width           =   2610
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line12 
         BorderColor     =   &H80000003&
         X1              =   375
         X2              =   2475
         Y1              =   1230
         Y2              =   1230
      End
      Begin VB.Line Line11 
         BorderColor     =   &H80000005&
         X1              =   375
         X2              =   2475
         Y1              =   1245
         Y2              =   1245
      End
   End
   Begin VB.Frame fraDatabaseTables 
      Caption         =   "   Select Tables   "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1920
      Left            =   135
      TabIndex        =   28
      Top             =   6165
      Width           =   2895
      Begin VB.CommandButton cmdCreateProperties 
         Caption         =   "&Properties"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   855
         TabIndex        =   8
         Top             =   1395
         Width           =   1110
      End
      Begin VB.Line Line6 
         BorderColor     =   &H80000005&
         X1              =   375
         X2              =   2475
         Y1              =   1245
         Y2              =   1245
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000003&
         X1              =   375
         X2              =   2475
         Y1              =   1230
         Y2              =   1230
      End
      Begin VB.Label lblInstructions1 
         Caption         =   "Check the table(s) you wish to create a data class for from the list above, then click the &Properties button below."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   135
         TabIndex        =   29
         Top             =   315
         Width           =   2610
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraClassModule 
      Caption         =   "   Class Modules   "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6645
      Left            =   10095
      TabIndex        =   27
      Top             =   1440
      Width           =   3930
      Begin VB.CommandButton cmdClose 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2190
         TabIndex        =   18
         Top             =   6120
         Width           =   1110
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "&View"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   645
         TabIndex        =   17
         Top             =   6120
         Width           =   1110
      End
      Begin VB.ListBox lstCreatedClasses 
         Height          =   2790
         Left            =   345
         TabIndex        =   16
         Top             =   3210
         Width           =   3225
      End
      Begin VB.CommandButton cmdCreateClassMods 
         Caption         =   "C&reate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2190
         TabIndex        =   15
         Top             =   2295
         Width           =   1110
      End
      Begin VB.TextBox txtFolderPath 
         Height          =   330
         Left            =   180
         TabIndex        =   13
         Top             =   1785
         Width           =   3615
      End
      Begin VB.CommandButton cmdFolder 
         Caption         =   "&Folder..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   645
         TabIndex        =   14
         Top             =   2295
         Width           =   1110
      End
      Begin VB.Label lblSaveDefault 
         AutoSize        =   -1  'True
         Caption         =   "(Defaults to C:\)"
         Height          =   195
         Left            =   1080
         TabIndex        =   44
         Top             =   1560
         Width           =   1125
      End
      Begin VB.Label lblCreatedClasses 
         AutoSize        =   -1  'True
         Caption         =   "Class modules created this session:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   390
         TabIndex        =   35
         Top             =   3000
         Width           =   3030
      End
      Begin VB.Label lblInstructions3 
         Caption         =   "Press the &Create button to create and save the selected data class modules."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   225
         TabIndex        =   34
         Top             =   1050
         Width           =   3615
      End
      Begin VB.Label lblInstructions2 
         Caption         =   "Enter the path you want to save the class module(s) into (the &Folder button diplays all available folders to select from).   "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   225
         TabIndex        =   33
         Top             =   300
         Width           =   3630
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line14 
         BorderColor     =   &H80000005&
         X1              =   360
         X2              =   3540
         Y1              =   2835
         Y2              =   2835
      End
      Begin VB.Line Line13 
         BorderColor     =   &H80000003&
         X1              =   360
         X2              =   3540
         Y1              =   2820
         Y2              =   2820
      End
      Begin VB.Label lblSaveTo 
         AutoSize        =   -1  'True
         Caption         =   "Save To:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   225
         TabIndex        =   32
         Top             =   1560
         Width           =   795
      End
   End
   Begin VB.ListBox lstFields 
      Height          =   4110
      Left            =   6615
      Style           =   1  'Checkbox
      TabIndex        =   11
      Top             =   1545
      Width           =   3285
   End
   Begin VB.ListBox lstProcessedTables 
      Height          =   4560
      Left            =   3360
      Style           =   1  'Checkbox
      TabIndex        =   9
      Top             =   1545
      Width           =   2895
   End
   Begin VB.ListBox lstTables 
      Height          =   4560
      Left            =   135
      Style           =   1  'Checkbox
      TabIndex        =   7
      Top             =   1545
      Width           =   2895
   End
   Begin VB.Frame fraDatabase 
      Caption         =   "   Database Connection   "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   135
      TabIndex        =   19
      Top             =   105
      Width           =   13890
      Begin VB.CommandButton cmdPath 
         Caption         =   "&Find DB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8655
         TabIndex        =   3
         Top             =   465
         Width           =   960
      End
      Begin VB.TextBox txtDatabase 
         Height          =   300
         Left            =   5145
         TabIndex        =   2
         Top             =   465
         Width           =   3375
      End
      Begin VB.ComboBox cboDatabaseType 
         Height          =   315
         Left            =   330
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   465
         Width           =   1635
      End
      Begin VB.CommandButton cmdConnection 
         Caption         =   "Connect"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12630
         TabIndex        =   6
         Top             =   345
         Width           =   1110
      End
      Begin VB.TextBox txtPassword 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   11070
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   465
         Width           =   975
      End
      Begin VB.TextBox txtUserName 
         Height          =   300
         Left            =   9825
         TabIndex        =   4
         Top             =   465
         Width           =   975
      End
      Begin VB.TextBox txtServer 
         Height          =   300
         Left            =   2325
         TabIndex        =   1
         Top             =   465
         Width           =   2490
      End
      Begin VB.Label lblDatabase 
         AutoSize        =   -1  'True
         Caption         =   "Database:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5145
         TabIndex        =   45
         Top             =   240
         Width           =   885
      End
      Begin VB.Line Line8 
         BorderColor     =   &H80000005&
         X1              =   12315
         X2              =   12315
         Y1              =   285
         Y2              =   735
      End
      Begin VB.Line Line7 
         BorderColor     =   &H80000003&
         X1              =   12300
         X2              =   12300
         Y1              =   285
         Y2              =   735
      End
      Begin VB.Label lblDatabaseType 
         AutoSize        =   -1  'True
         Caption         =   "Type:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   330
         TabIndex        =   26
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblPassword 
         AutoSize        =   -1  'True
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   11070
         TabIndex        =   22
         Top             =   240
         Width           =   885
      End
      Begin VB.Label lblUserName 
         AutoSize        =   -1  'True
         Caption         =   "User Id:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9825
         TabIndex        =   21
         Top             =   240
         Width           =   690
      End
      Begin VB.Label lblServer 
         AutoSize        =   -1  'True
         Caption         =   "Server:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2325
         TabIndex        =   20
         Top             =   240
         Width           =   630
      End
   End
   Begin VB.Label lblDefaultsInfo 
      AutoSize        =   -1  'True
      Caption         =   "(Defaults to all)"
      Height          =   195
      Left            =   7620
      TabIndex        =   50
      Top             =   5895
      Width           =   1050
   End
   Begin VB.Label lblIntToLongHelp 
      AutoSize        =   -1  'True
      Caption         =   " ? "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9735
      TabIndex        =   49
      Top             =   6120
      Width           =   255
   End
   Begin VB.Label lblMoreInfo 
      AutoSize        =   -1  'True
      Caption         =   "(Applies to all tables in Process list)"
      Height          =   195
      Left            =   6990
      TabIndex        =   48
      Top             =   6375
      Width           =   2445
   End
   Begin VB.Label lblAbout 
      AutoSize        =   -1  'True
      Caption         =   "About DataClassBuilder..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   12060
      TabIndex        =   46
      Top             =   8190
      Width           =   2190
   End
   Begin VB.Label lblActions 
      BorderStyle     =   1  'Fixed Single
      Height          =   270
      Left            =   -30
      TabIndex        =   39
      Top             =   8130
      Width           =   14340
   End
   Begin VB.Label lblInstructions6 
      AutoSize        =   -1  'True
      Caption         =   "Check the fields to include as class properties."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6675
      TabIndex        =   38
      Top             =   5700
      Width           =   3195
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line10 
      BorderColor     =   &H80000003&
      X1              =   6420
      X2              =   6420
      Y1              =   1785
      Y2              =   5910
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000005&
      X1              =   6435
      X2              =   6435
      Y1              =   1785
      Y2              =   5910
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000005&
      X1              =   3195
      X2              =   3195
      Y1              =   1785
      Y2              =   5910
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000003&
      X1              =   3180
      X2              =   3180
      Y1              =   1785
      Y2              =   5910
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      Caption         =   "Table Fields (Class Properties):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6615
      TabIndex        =   25
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label lblProcessedTables 
      AutoSize        =   -1  'True
      Caption         =   "Tables to Process (Class Objects):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3360
      TabIndex        =   24
      Top             =   1320
      Width           =   2940
   End
   Begin VB.Label lblTableList 
      AutoSize        =   -1  'True
      Caption         =   "Database Tables:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   165
      TabIndex        =   23
      Top             =   1320
      Width           =   1515
   End
End
Attribute VB_Name = "frmClassBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' instance of error object
Private m_objError As clsError
Private m_objProcess As clsProcess
Private m_objTableList As clsComboList
Private m_objProcessedList As clsComboList
Private m_objFieldList As clsComboList
Private m_objClassList As clsComboList
Private WithEvents m_objStatus As clsStatus
Attribute m_objStatus.VB_VarHelpID = -1

Public Property Set ObjAppProcess(vData As clsProcess)
    Set m_objProcess = vData
End Property

Public Property Set ObjError(vData As clsError)
    Set m_objError = vData
End Property

Private Sub cboDatabaseType_Click()
    If cboDatabaseType.Text = "Access" Then
        With txtDatabase
            .Enabled = True
            .BackColor = &H80000005
        End With
        With txtServer
            .Enabled = False
            .BackColor = &H8000000F
        End With
        lblDatabase.Caption = "Access MDB Name/Path:"
        cmdPath.Enabled = True
    Else
        If cboDatabaseType.Text = "SQL Server" Then
            With txtDatabase
                .BackColor = &H80000005
                .Enabled = True
            End With
        Else
            With txtDatabase
                .Enabled = False
                .BackColor = &H8000000F
            End With
        End If
        With txtServer
            .Enabled = True
            .BackColor = &H80000005
        End With
        cmdPath.Enabled = False
        lblDatabase.Caption = "Database:"
    End If
    
End Sub

Private Sub chkIntegerToLong_Click()
    m_objProcess.ConvertIntegerToLong = chkIntegerToLong.Value
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdConnection_Click()
    If CheckForDatabaseValues Then
        SetDatabaseConnection
        ClearTableFieldListboxes
        m_objProcess.ClearTypeArrays
    End If
End Sub

Private Function CheckForDatabaseValues() As Boolean
    ' make sure everything is filled in to create a connection
    Dim strMsg As String
    Dim strDatabaseType As String
    
    CheckForDatabaseValues = True
    
    strDatabaseType = cboDatabaseType.Text
    
    strMsg = "Please fill in the following values for the " & strDatabaseType & _
                " database connection:" & vbCrLf
    
    If Len(Trim(strDatabaseType)) = 0 Then
        strMsg = strMsg & vbCrLf & "Database Type" & vbCrLf
        CheckForDatabaseValues = False
    End If
    If Len(Trim(txtServer.Text)) = 0 And strDatabaseType <> "Access" Then
        strMsg = strMsg & vbCrLf & "Server" & vbCrLf
        CheckForDatabaseValues = False
    End If
    If strDatabaseType = "Access" Or strDatabaseType = "SQL Server" Then
        If Len(Trim(txtDatabase.Text)) = 0 Then
            strMsg = strMsg & vbCrLf & "Database" & vbCrLf
            CheckForDatabaseValues = False
        End If
    End If
    If Len(Trim(txtUserName.Text)) = 0 And strDatabaseType <> "Access" Then
        strMsg = strMsg & vbCrLf & "User Id" & vbCrLf
        CheckForDatabaseValues = False
    End If
    If Len(Trim(txtPassword.Text)) = 0 And strDatabaseType <> "Access" Then
        strMsg = strMsg & vbCrLf & "Password" & vbCrLf
        CheckForDatabaseValues = False
    End If
    
    If Not CheckForDatabaseValues Then
        MsgBox strMsg, vbExclamation, "Database Connection"
    End If

End Function

Private Sub SetDatabaseConnection()
    Dim bolSuccess As Boolean
    
    On Error GoTo err_SetDatabaseConnection
    
    Me.MousePointer = vbHourglass
    
    m_objStatus.PostStatus " Making " & cboDatabaseType.Text & " database connection - Please wait..."

    bolSuccess = m_objProcess.MakeDataBaseConnection(cboDatabaseType.Text, txtServer.Text, txtUserName.Text, txtPassword.Text, txtDatabase.Text)
    If bolSuccess Then
        m_objStatus.PostStatus " Creating table schema from the database - Please wait..."
        m_objProcess.OpenTableSchema
    Else
        Err.Raise vbObjectError + 1077, , "Error making database connection."
    End If
    
exit_SetDatabaseConnection:
    Me.MousePointer = vbDefault
    m_objStatus.PostStatus " "
    Exit Sub
    
err_SetDatabaseConnection:
    Me.MousePointer = vbDefault
    With Err
        m_objError.UpdateLogFile "frmClassBuilder", "SetDatabaseConnection", _
            .Number, .Description, .Source
    End With
    Resume exit_SetDatabaseConnection

End Sub

Private Function CheckConnectionEntry() As Boolean
    ' verify user filled in all necessary elements to make
    ' a database connection
    Dim strEntry As String
        
    CheckConnectionEntry = False
    
    If Len(cboDatabaseType.Text) = 0 Or _
        Len(txtServer.Text) = 0 Or _
        (cboDatabaseType.Text <> "Access" And Len(txtUserName.Text) = 0) Or _
        (cboDatabaseType.Text <> "Access" And Len(txtPassword.Text) = 0) Then
            MsgBox "Please fill in all the necessary connection information." & vbCrLf & vbCrLf & _
                "*** Note ***" & vbCrLf & _
                "User Id and Password may not be necessary for Access databases.", vbInformation, _
                App.EXEName & " - Connection String"
    Else
        CheckConnectionEntry = True
    End If
    
End Function

Private Sub cmdCreateClassMods_Click()
    If lstProcessedTables.ListCount > 0 Then
        CallCreateClassModules
    End If
End Sub

Private Sub CallCreateClassModules()
    Dim strSavePath As String
    Dim lngRet As Long
    
    On Error GoTo err_CallCreateClassModules
    
    strSavePath = Trim(txtFolderPath.Text)
    
    ' if user puts nothing in, we default to C:\
    If Len(Trim(strSavePath)) <> 0 Then
        strSavePath = AddBackSlash(strSavePath)
        ' make sure folder exists
        If Len(Dir$(strSavePath, vbDirectory)) = 0 Then
            lngRet = MsgBox(strSavePath & " is not a valid path.  Press OK to save " & _
                        "the class module in the default C:\ drive or Cancel " & _
                        "to select a new folder to save into.", vbOKCancel + vbQuestion, _
                        "Invalid Path")
            If lngRet = vbCancel Then
                SelectAllText txtFolderPath
                Exit Sub
            Else
                strSavePath = "C:\"
            End If
        End If
    Else
        strSavePath = "C:\"
    End If
    Screen.MousePointer = vbHourglass
    m_objProcess.CreateClassModules strSavePath
    ClearTableFieldListboxes
    
exit_CallCreateClassModules:
    m_objStatus.PostStatus " "
    Screen.MousePointer = vbDefault
    Exit Sub
    
err_CallCreateClassModules:
    Me.MousePointer = vbDefault
    With Err
        m_objError.UpdateLogFile "frmClassBuilder", "CallCreateClassModules", _
            .Number, .Description, .Source
    End With
    Resume exit_CallCreateClassModules
    
End Sub

Private Sub cmdCreateProperties_Click()
    CreateFieldProperties
End Sub

Private Sub CreateFieldProperties()
    Dim lngSelected As Long
    
    On Error GoTo err_CreateFieldProperties
    
    Me.MousePointer = vbHourglass
    ' get the number of selected tables
    lngSelected = lstTables.SelCount
    If lngSelected > 0 Then
        m_objProcess.CreateTableClassProperties lngSelected
        ' highlight 1st item in listbox
        If m_objProcessedList.ListCount > 0 Then
            m_objProcessedList.SetListboxListIndex 0
            lstProcessedTables_Click
        End If
    End If
    
exit_CreateFieldProperties:
    Me.MousePointer = vbDefault
    m_objStatus.PostStatus " "
    Exit Sub
    
err_CreateFieldProperties:
    Me.MousePointer = vbDefault
    With Err
        m_objError.UpdateLogFile "frmClassBuilder", "CreateFieldProperties", _
            .Number, .Description, .Source
    End With
    Resume exit_CreateFieldProperties
    
End Sub

Private Sub ClearTableFieldListboxes()
    m_objProcessedList.ClearListbox
    m_objFieldList.ClearListbox
    m_objTableList.DeselectAllListboxItems
    Me.txtPrimaryKey = ""
    Me.chkIntegerToLong.Value = 0
End Sub

Private Sub cmdFolder_Click()
    Dim objFolders As clsCommonDialog
    
    Set objFolders = New clsCommonDialog
    
    txtFolderPath.Text = objFolders.BrowseDirectory
    
    Set objFolders = Nothing
    
End Sub

Private Sub cmdPath_Click()
    Dim objFile As clsCommonDialog
    
    Set objFile = New clsCommonDialog
    
    With objFile
        .ObjectOwner = Me
        .Filter = "MDB Files (*.mdb)|*.mdb|"
        .WindowTitle = "Select Access Database"
        txtDatabase.Text = .GetFileOpenName
    End With
    
    Set objFile = Nothing
    
End Sub

Private Sub cmdPrimaryKey_Click()
    With lstFields
        If Len(Trim(.Text)) > 0 Then
            ' drop selected field into textbox
            ' handle any spaces in selected Primary Key
            If InStr(.Text, "]") = 0 Then
                txtPrimaryKey.Text = Trim$(Left$(.Text, InStr(.Text, " ") - 1))
            Else
                txtPrimaryKey.Text = Trim$(Left$(.Text, InStr(.Text, "]") + 1))
            End If
            ' make sure the item dropped into this textbox is also selected
            ' in the listbox for property creation
            .Selected(lstFields.ListIndex) = True
            ' fire off the listbox click event
            lstFields_Click
            m_objProcess.SetPrimaryField .ListIndex
        End If
    End With

End Sub

Private Sub cmdRemove_Click()
    DeleteTablesFromSelectionList
End Sub

Private Sub DeleteTablesFromSelectionList()
    Dim lngRet As Long

    On Error GoTo err_DeleteTablesFromSelectionList
                    
    If m_objProcessedList.ListCount > 0 Then
        lngRet = MsgBox("Are you sure you want to remove the selected table(s) from " & _
                    "the Tables to Process list?", vbYesNo + vbQuestion, "Remove Table")
        If lngRet = vbYes Then
            m_objProcess.DeleteTableFromArray
        
            ' if no error in the above call we need to
            ' reset the fields list
            If m_objProcessedList.ListCount > 0 Then
                'm_objProcessedList.SetComboListIndex 0
                lstProcessedTables_Click
            Else
                ' make sure primary key text is blank
                txtPrimaryKey.Text = ""
            End If
        End If
    End If
    
    Exit Sub
    
err_DeleteTablesFromSelectionList:
    With Err
        m_objError.UpdateLogFile "frmClassBuilder", "DeleteTablesFromSelectionList", _
            .Number, .Description, .Source
    End With
        
End Sub


Private Sub cmdView_Click()
 ' view newly created class in notepad
    If m_objClassList.ListCount > 0 Then
        If lstCreatedClasses.SelCount > 0 Then
            m_objProcess.ViewClassModInNotepad lstCreatedClasses.Text
        End If
    End If
End Sub

Private Sub Form_Load()
    
    Set m_objStatus = New clsStatus
    
    SetFormDisplay
    CreateListObjects
    FillServerCombo
    SetAppProcessClass
    
End Sub

Private Sub SetFormDisplay()
    Dim objForm As clsFormUtilities
    
    Set objForm = New clsFormUtilities
    
    With objForm
        .Form_hWnd = Me.hWnd
        .SetToCustomDialog
    End With
    CenterForm Me
    CenterControl fraDatabase, Me, False
    fraPKeyInstructions.Visible = False
    cboDatabaseType_Click
    Set objForm = Nothing
    
End Sub

Private Sub SetAppProcessClass()
    ' instatiate process class (handles all non GUI stuff)
    Set m_objProcess = New clsProcess
    With m_objProcess
        Set .ObjError = m_objError
        Set .ObjTableListbox = m_objTableList
        Set .ObjProcessedList = m_objProcessedList
        Set .ObjFieldList = m_objFieldList
        Set .ObjClassList = m_objClassList
        Set .ObjStatus = m_objStatus
    End With
End Sub

Private Sub CreateListObjects()
    ' create 3 intances of the clsComboList class to get
    ' API functionality for the form listboxes
    Dim lngTabStops(3) As Long
    
    On Error GoTo err_CreateListObjects
    
    Set m_objTableList = New clsComboList
    With m_objTableList
        Set .ListboxObject = lstTables
        .RedrawListbox
    End With
    
    Set m_objProcessedList = New clsComboList
    With m_objProcessedList
        Set .ListboxObject = lstProcessedTables
        .RedrawListbox
    End With
    
    Set m_objFieldList = New clsComboList
    With m_objFieldList
        Set .ListboxObject = lstFields
        .RedrawListbox
        ' add tabs to field listbox control
        lngTabStops(0) = 100
        lngTabStops(1) = 15
        lngTabStops(2) = 10
        .SetLBTabStops lngTabStops
    End With
    
    Set m_objClassList = New clsComboList
    With m_objClassList
        Set .ListboxObject = lstCreatedClasses
        .RedrawListbox
    End With
    
    Exit Sub
    
err_CreateListObjects:
    With Err
        m_objError.UpdateLogFile "frmClassBuilder", "CreateListObjects", _
            .Number, .Description, .Source
    End With
    
End Sub

Private Sub FillServerCombo()
    Dim objCombo As clsComboList
    Dim arrList(4) As String
    
    arrList(0) = "Access"
    arrList(1) = "Oracle"
    arrList(2) = "Sybase"
    arrList(3) = "SQL Server"
    
    Set objCombo = New clsComboList
    
    With objCombo
        Set .ComboboxObject = cboDatabaseType
        .LoadComboBoxArray arrList
    End With
    
exit_FillServerCombo:
    Set objCombo = Nothing
    Exit Sub
    
err_FillServerCombo:
    With Err
        m_objError.UpdateLogFile "frmClassBuilder", "FillServerCombo", _
            .Number, .Description, .Source
    End With
    Resume exit_FillServerCombo
    
End Sub

Private Sub lblAbout_Click()
    frmAbout.Show
End Sub

Private Sub lblAbout_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblAbout.ForeColor = &HFF00FF
End Sub

Private Sub lblAbout_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblAbout.ForeColor = &H80000012
End Sub

Private Sub lblIntToLongHelp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngTextHeight As Long
    
    lngTextHeight = TextHeight("A")
    
    lblPrimaryKeyHelp.ForeColor = &HFF0000
    
    With lblPKeyInstructions
        .ForeColor = &HFF0000
        .BackColor = vbWhite
        .FontBold = True
        .Caption = "Note: Because some databases Integer data types can contain values that are " & _
                "greater than the Max value of a VB Integer, " & _
                "you can select to have every Integer field set up as a Long in the " & _
                "class module."
        .ZOrder 0
        .Visible = True
    End With
    
    With fraPKeyInstructions
        .Move fraPrimaryKey.Left + fraPrimaryKey.Width, fraPrimaryKey.Top - (lblPrimaryKeyHelp.Height * 3) ' - .Height
        .Visible = True
    End With
    
End Sub

Private Sub lblIntToLongHelp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblPrimaryKeyHelp.ForeColor = &H80000012
    fraPKeyInstructions.Visible = False
End Sub

Private Sub lblPrimaryKeyHelp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngTextHeight As Long
    
    lngTextHeight = TextHeight("A")
    
    lblPrimaryKeyHelp.ForeColor = &HFF0000
    
    With lblPKeyInstructions
        .ForeColor = &HFF0000
        .BackColor = vbWhite
        .FontBold = True
        .Caption = "Note: The primary key is automatically selected " & _
            "by the Class Builder application if one is not found in the table.  You can " & _
            "replace the primary key with any field from the Table Fields listbox."
          
        .ZOrder 0
        .Visible = True
    End With
    
    With fraPKeyInstructions
        .Move fraPrimaryKey.Left + fraPrimaryKey.Width, fraPrimaryKey.Top - (lblPrimaryKeyHelp.Height * 3) ' - .Height
        .Visible = True
    End With
    
End Sub

Private Sub lblPrimaryKeyHelp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblPrimaryKeyHelp.ForeColor = &H80000012
    fraPKeyInstructions.Visible = False
End Sub

Private Sub lstCreatedClasses_DblClick()
    cmdView_Click
End Sub

Private Sub lstFields_Click()
    Dim intSelected As Integer
    Dim bolIsSelected As Boolean
    
    If lstFields.ListCount > 0 Then
        intSelected = (CInt(m_objFieldList.ReturnListboxListIndex))
        bolIsSelected = lstFields.Selected(intSelected)
           
        ' check if user is try to deselect the Primary Key - Not allowed
        If Trim$(Left$(lstFields.Text, InStr(lstFields.Text, " ") - 1)) <> Trim$(txtPrimaryKey.Text) Then
            ' select/deselect field for property creation
            m_objProcess.SetFieldForPropertyCreation CLng(intSelected), bolIsSelected
        Else
            If lstFields.Selected(intSelected) = True Then Exit Sub
            MsgBox "You cannot deselect the Primary Key.  If you really want to deselect " & _
                "this field, Please select a new Primary key and then uncheck this field."
            m_objFieldList.SetListboxListIndex CLng(intSelected)
            lstFields.Selected(intSelected) = True
        End If
    End If
    
End Sub

Private Sub lstProcessedTables_Click()
    m_objFieldList.ClearListbox
    If lstProcessedTables.ListCount > 0 Then
        m_objProcess.DisplayFieldProperties lstProcessedTables.Text
    End If
End Sub

Private Sub m_objStatus_StatusChange(sDescription As String)
    lblActions.Caption = sDescription
    lblActions.Refresh
End Sub
