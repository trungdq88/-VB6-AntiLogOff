VERSION 5.00
Object = "{D1ECD258-D329-4388-AB83-DEC261A66B86}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anti LogOFF - Perfect Software"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9705
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   9705
   StartUpPosition =   2  'CenterScreen
   Begin UniControls.UniButton UniButton5 
      Height          =   375
      Left            =   7560
      TabIndex        =   10
      Top             =   4800
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Icon            =   "frmMain.frx":57E2
      Style           =   2
      Caption         =   "Tho6ng tin chu7o7ng tri2nh"
      IconAlign       =   3
      iNonThemeStyle  =   2
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedBordersByTheme=   0   'False
      ShowFocusRectangle=   0   'False
   End
   Begin UniControls.UniFrame UniFrame2 
      Height          =   1695
      Left            =   120
      Top             =   5280
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   2990
      MaskColor       =   16711935
      FrameColor      =   -2147483629
      Style           =   0
      Caption         =   "Hu7o71ng da64n su73 du5ng"
      TextColor       =   13579779
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin UniControls.UniLabel UniLabel6 
         Height          =   255
         Left            =   120
         Top             =   1320
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   450
         Caption         =   $"frmMain.frx":57FE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UniControls.UniLabel UniLabel5 
         Height          =   495
         Left            =   120
         Top             =   840
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   873
         Caption         =   $"frmMain.frx":5887
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UniControls.UniLabel UniLabel4 
         Height          =   255
         Left            =   120
         Top             =   600
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   450
         Caption         =   $"frmMain.frx":5965
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin UniControls.UniLabel UniLabel3 
         Height          =   255
         Left            =   120
         Top             =   360
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   450
         Caption         =   "- Nha61n va2o nu1t phi1a tre6n d9e63 ba65t chu71c na8ng pho2ng cho61ng lo64i kho6ng the63 Logon."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   8160
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin UniControls.UniFrame UniFrame3 
      Height          =   2535
      Left            =   6360
      Top             =   1560
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   4471
      MaskColor       =   16711935
      FrameColor      =   12632256
      Caption         =   "Tie65n i1ch"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin UniControls.UniButton UniButton10 
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1920
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         Icon            =   "frmMain.frx":5A0C
         Style           =   2
         Caption         =   "Xo1a toa2n bo65 ta65p tin Autorun"
         IconAlign       =   3
         iNonThemeStyle  =   2
         MaskColor       =   16711935
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedBordersByTheme=   0   'False
         ShowFocusRectangle=   0   'False
      End
      Begin UniControls.UniButton UniButton9 
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         Icon            =   "frmMain.frx":5A28
         Style           =   2
         Caption         =   "Phu5c ho62i ta65p tin Hosts"
         IconAlign       =   3
         iNonThemeStyle  =   2
         MaskColor       =   16711935
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedBordersByTheme=   0   'False
         ShowFocusRectangle=   0   'False
      End
      Begin UniControls.UniButton UniButton7 
         Height          =   735
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1296
         Icon            =   "frmMain.frx":5A44
         Style           =   2
         Caption         =   "Phu5c ho62i ta61t ca3 ca1c chu71c na8ng cu3a he65 d9ie62u ha2nh (Task Manager, Registry...)"
         IconAlign       =   3
         iNonThemeStyle  =   2
         MaskColor       =   16711935
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedBordersByTheme=   0   'False
         ShowFocusRectangle=   0   'False
      End
   End
   Begin UniControls.UniFrame UniFrame1 
      Height          =   2535
      Left            =   120
      Top             =   1560
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   4471
      MaskColor       =   16711935
      Caption         =   "Co6ng cu5"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin UniControls.UniButton UniButton4 
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   2040
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         Icon            =   "frmMain.frx":5A60
         Style           =   2
         Caption         =   "Co6ng cu5 die65t Virus vo71i ma64u"
         IconAlign       =   3
         iNonThemeStyle  =   2
         MaskColor       =   16711935
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedBordersByTheme=   0   'False
         ShowFocusRectangle=   0   'False
      End
      Begin UniControls.UniButton UniButton3 
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         Icon            =   "frmMain.frx":5FFA
         Style           =   2
         Caption         =   "Windows Command Promtp"
         IconAlign       =   3
         iNonThemeStyle  =   2
         MaskColor       =   16711935
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedBordersByTheme=   0   'False
         ShowFocusRectangle=   0   'False
      End
      Begin UniControls.UniButton UniButton2 
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         Icon            =   "frmMain.frx":6154
         Style           =   2
         Caption         =   "Windows Registry Editor"
         IconAlign       =   3
         iNonThemeStyle  =   2
         MaskColor       =   16711935
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedBordersByTheme=   0   'False
         ShowFocusRectangle=   0   'False
      End
      Begin UniControls.UniButton UniButton1 
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         Icon            =   "frmMain.frx":62AE
         Style           =   2
         Caption         =   "Windows Task Manager"
         IconAlign       =   3
         iNonThemeStyle  =   2
         MaskColor       =   16711935
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedBordersByTheme=   0   'False
         ShowFocusRectangle=   0   'False
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   975
      Left            =   8400
      Picture         =   "frmMain.frx":6408
      ScaleHeight     =   915
      ScaleWidth      =   795
      TabIndex        =   1
      Top             =   -120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Left            =   8640
      Picture         =   "frmMain.frx":C825
      ScaleHeight     =   1035
      ScaleWidth      =   1035
      TabIndex        =   0
      Top             =   -240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin UniControls.UniLabel UniLabel2 
      Height          =   495
      Left            =   240
      Top             =   4560
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   873
      Caption         =   "Chu71c na8ng pho2ng cho61ng lo64i kho6ng the63 Log Off d9ang:"
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
      Height          =   975
      Left            =   1320
      Top             =   360
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   1720
      Alignment       =   1
      Caption         =   "Ne61u ma1y cu3a ba5n hie65n ta5i kho6ng the63 Log On. Ha4y nha61n va2o nu1t Su74a Chu74a be6n du7o71i"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   33023
      Link            =   ""
   End
   Begin VB.Image Image1 
      Height          =   330
      Left            =   3840
      Picture         =   "frmMain.frx":12DD2
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Image TurnOnOff 
      Height          =   915
      Left            =   5640
      Top             =   4320
      Width           =   855
   End
   Begin VB.Image cmdFixLogOff 
      Height          =   1920
      Left            =   3960
      Picture         =   "frmMain.frx":195F1
      Top             =   2280
      Width           =   1920
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub InitCommonControls Lib "comctl32" ()
Private Declare Function LoadCursor Lib "user32.dll" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Long) As Long

Const IDC_HAND As Long = 32649 'Ban tay



Private Sub cmdFixLogOff_Click()
SetCursor LoadCursor(ByVal 0&, IDC_HAND)
RegistryClean
UniMsgBox "D9a4 su73a xong lo64i kho6ng the63 Log On!", vbOKOnly + vbInformation, "Tho6ng ba1o", Me.hWnd
End Sub

Private Sub cmdFixLogOff_DblClick()
SetCursor LoadCursor(ByVal 0&, IDC_HAND)
End Sub

Private Sub cmdFixLogOff_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SetCursor LoadCursor(ByVal 0&, IDC_HAND)
End Sub

Private Sub cmdFixLogOff_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SetCursor LoadCursor(ByVal 0&, IDC_HAND)
End Sub

Private Sub cmdFixLogOff_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
SetCursor LoadCursor(ByVal 0&, IDC_HAND)
End Sub



Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_Load()
If CheckLog = True Then
TurnOnOff.Picture = Picture1.Picture
Else
TurnOnOff.Picture = Picture2.Picture
End If
End Sub

Private Sub Label1_Click()

End Sub



Private Sub Form_Unload(Cancel As Integer)
UniMsgBox "Ne61u nhu7 ba5n d9ang cha5y chu7o7ng tri2nh trong tra5ng tha1i Log Off, sau khi chu7o7ng tri2nh ta81t se4 xua61t hie65n 1 lo64i nho3. Ba5n cu71 nha61n OK va2 Log On va2o ma1y ti1nh 1 ca1ch bi2nh thu7o72ng.", vbOKOnly + vbInformation, "Chu1 y1", Me.hWnd
End
End Sub

Private Sub TurnOnOff_Click()
SetCursor LoadCursor(ByVal 0&, IDC_HAND)

If TurnOnOff.Picture = Picture1.Picture Then
    If UniMsgBox("Chu1 y1: Ne61u ba5n ta81t chu71c na8ng pho2ng cho61ng lo64i na2y thi2 chu7o7ng tri2nh se4 kho6ng hoa5t d9o65ng khi ba5n nha61n phi1m SHIFT 5 la62n." & vbCrLf & "Ba5n va64n muo61n ta81t?", vbYesNo + vbCritical, "Chu1 y1", Me.hWnd) = vbYes Then
        TurnOnOff.Picture = Picture2.Picture
        GoLogOff
    End If
Else
TurnOnOff.Picture = Picture1.Picture
CaiDatLogOff
End If
End Sub

Private Sub TurnOnOff_DblClick()
SetCursor LoadCursor(ByVal 0&, IDC_HAND)
End Sub

Private Sub TurnOnOff_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SetCursor LoadCursor(ByVal 0&, IDC_HAND)
End Sub

Private Sub TurnOnOff_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SetCursor LoadCursor(ByVal 0&, IDC_HAND)
End Sub

Private Sub UniButton1_Click()
Shell "C:\windows\system32\taskmgr.exe", vbNormalFocus
End Sub

Private Sub UniButton10_Click()
On Error Resume Next
Dim u As Integer
For u = 0 To Drive1.ListCount - 1
    KillFile Left(Drive1.List(u), 1) & ":\autorun.inf"
Next u
UniMsgBox "D9a4 xo1a he61t ta61t ca3 ca1c ta65p tin Autorun trong ca1c o63 d9i4a!", vbOKOnly + vbInformation, "Tho6ng ba1o", Me.hWnd

End Sub

Private Sub UniButton11_Click()
Shell AppPath & "VirusRemoveAll.exe", vbNormalFocus
End Sub

Private Sub UniButton2_Click()
Shell "C:\windows\regedit.exe", vbNormalFocus
End Sub

Private Sub UniButton3_Click()
Shell "C:\windows\system32\cmd.exe", vbNormalFocus
End Sub

Private Sub UniButton4_Click()
Dim ocxDir$
ocxDir = "C:\WINDOWS\VirusRemoveAll.exe"
If (FileExists(ocxDir) = False) Then
Dim bytResourceData() As Byte
bytResourceData = LoadResData(102, "CUSTOM")
Open ocxDir For Binary Shared As #1
Put #1, 1, bytResourceData
Close #1
End If


Shell "C:\WINDOWS\VirusRemoveAll.exe", vbNormalFocus
End Sub

Private Sub UniButton5_Click()
frmAbout.Show
End Sub

Private Sub UniButton6_Click()
Shell AppPath & "PerfectStartUpManager.exe", vbNormalFocus
End Sub

Private Sub UniButton7_Click()
RegistryClean
UniMsgBox "D9a4 phu5c ho62i xong ta61t ca3 ca1c chu71c na8ng cu3a he65 d9ie62u ha2nh!", vbOKOnly + vbInformation, "Tho6ng ba1o", Me.hWnd
End Sub

Private Sub UniButton9_Click()
KillFile "C:\WINDOWS\system32\drivers\etc\hosts"
WriteFileUni "C:\WINDOWS\system32\drivers\etc\hosts", "127.0.0.1    localhost"
UniMsgBox "D9a4 phu5c ho62i xong ta65p tin hosts!", vbOKOnly + vbInformation, "Xong!", Me.hWnd

End Sub
