VERSION 5.00
Object = "{D1ECD258-D329-4388-AB83-DEC261A66B86}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
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
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin UniControls.UniButton UniButton2 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Icon            =   "frmAbout.frx":57E2
      Style           =   2
      Caption         =   "Trang chu3"
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
      Left            =   3360
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Icon            =   "frmAbout.frx":57FE
      Style           =   2
      Caption         =   "D9o1ng"
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
   Begin UniControls.UniLabel UniLabel5 
      Height          =   495
      Left            =   120
      Top             =   1920
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   873
      Caption         =   $"frmAbout.frx":581A
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
      Left            =   600
      Top             =   1560
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   450
      Caption         =   "Website: http://phanmemtiengviet.co.cc"
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
      Left            =   600
      Top             =   1320
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   450
      Caption         =   "Email: dinhquangtrung90@yahoo.com"
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
   Begin UniControls.UniLabel UniLabel2 
      Height          =   255
      Left            =   600
      Top             =   1080
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   450
      Caption         =   "Ta1c gia3: D9inh Quang Trung (Perfect Software)"
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
   Begin UniControls.UniLabel UniLabel1 
      Height          =   255
      Left            =   0
      Top             =   480
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   450
      Alignment       =   1
      Caption         =   "Chu7o7ng tri2nh pho2ng cho61ng xa3y ra lo64i kho6ng the63 Log On"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   33023
   End
   Begin VB.Image Image1 
      Height          =   330
      Left            =   1320
      Picture         =   "frmAbout.frx":58A8
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UniButton1_Click()
Unload Me
End Sub

Private Sub UniButton2_Click()
Shell "explorer http://phanmemtiengviet.co.cc"
End Sub
