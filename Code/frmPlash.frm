VERSION 5.00
Object = "{D1ECD258-D329-4388-AB83-DEC261A66B86}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form frmPlash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin UniControls.UniLabel UniLabel4 
      Height          =   255
      Left            =   1560
      Top             =   2760
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   450
      Alignment       =   2
      BackStyle       =   0
      Caption         =   "Sa3n pha63m cu3a Perfect Software"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   8421504
   End
   Begin UniControls.UniLabel UniLabel3 
      Height          =   495
      Left            =   2400
      Top             =   1920
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      Caption         =   "5"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   49152
   End
   Begin UniControls.UniLabel UniLabel2 
      Height          =   615
      Left            =   240
      Top             =   1200
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1085
      Caption         =   "Nha61n phi1m Shift 5 la62n lie6n tie61p d9e63 kho73i d9o65ng chu7o7ng tri2nh"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   255
   End
   Begin UniControls.UniLabel UniLabel1 
      Height          =   1215
      Left            =   120
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2143
      Alignment       =   1
      BackStyle       =   0
      Caption         =   "Ma1y ti1nh cu3a ba5n d9ang d9u7o75c ba3o ve65 tra1nh kho3i  su75 co61 kho6ng the63 Log On."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   4320
      Top             =   720
   End
   Begin VB.Timer Timer2 
      Left            =   4320
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Left            =   4320
      Top             =   360
   End
   Begin VB.Image Image1 
      Height          =   900
      Left            =   1080
      Picture         =   "frmPlash.frx":0000
      Top             =   1800
      Width           =   2250
   End
End
Attribute VB_Name = "frmPlash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000

Dim m_lAlpha

Private Sub Form_Load()
    Dim lStyle As Long
    lStyle = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
    lStyle = lStyle Or WS_EX_LAYERED
    SetWindowLong Me.hWnd, GWL_EXSTYLE, lStyle
    SetLayeredWindowAttributes Me.hWnd, 0, 0, LWA_ALPHA
    Timer1.Interval = 50
    Timer2.Interval = 50
    Timer2.Enabled = False
    Timer1.Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then
        Cancel = True
        Timer2.Enabled = True
    End If
End Sub

Private Sub Timer1_Timer()
    m_lAlpha = m_lAlpha + 15
    If (m_lAlpha > 255) Then
        m_lAlpha = 255
        Timer1.Enabled = False
        Timer3.Enabled = True
    Else
        SetLayeredWindowAttributes Me.hWnd, 0, m_lAlpha, LWA_ALPHA
    End If
End Sub

Private Sub Timer2_Timer()
    m_lAlpha = m_lAlpha - 15
    If (m_lAlpha < 0) Then
        m_lAlpha = 0
        Unload Me
    Else
        SetLayeredWindowAttributes Me.hWnd, 0, m_lAlpha, LWA_ALPHA
    End If
End Sub

Private Sub Timer3_Timer()
Timer2.Enabled = True
Timer3.Enabled = False
End Sub

