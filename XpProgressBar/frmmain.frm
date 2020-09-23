VERSION 5.00
Begin VB.Form frmmain 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Xp-Progress Bar...NaGy"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   4425
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picProg 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1320
      ScaleHeight     =   195
      ScaleWidth      =   1620
      TabIndex        =   1
      Top             =   480
      Width           =   1620
      Begin VB.PictureBox picImage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   0
         Picture         =   "frmmain.frx":0000
         ScaleHeight     =   150
         ScaleWidth      =   390
         TabIndex        =   2
         Top             =   5
         Width           =   390
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   1560
      Top             =   1680
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   2040
      Top             =   1680
   End
   Begin VB.Timer Timer4 
      Interval        =   1000
      Left            =   1080
      Top             =   1680
   End
   Begin VB.Image i20 
      Height          =   165
      Left            =   3240
      Picture         =   "frmmain.frx":055A
      Top             =   1200
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image i19 
      Height          =   165
      Left            =   3120
      Picture         =   "frmmain.frx":0678
      Top             =   1200
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image i18 
      Height          =   165
      Left            =   3000
      Picture         =   "frmmain.frx":0796
      Top             =   1200
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image i17 
      Height          =   165
      Left            =   2880
      Picture         =   "frmmain.frx":08B4
      Top             =   1200
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image i16 
      Height          =   165
      Left            =   2760
      Picture         =   "frmmain.frx":09D2
      Top             =   1200
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image i15 
      Height          =   165
      Left            =   2640
      Picture         =   "frmmain.frx":0AF0
      Top             =   1200
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image i14 
      Height          =   165
      Left            =   2520
      Picture         =   "frmmain.frx":0C0E
      Top             =   1200
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image i13 
      Height          =   165
      Left            =   2400
      Picture         =   "frmmain.frx":0D2C
      Top             =   1200
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image i12 
      Height          =   165
      Left            =   2280
      Picture         =   "frmmain.frx":0E4A
      Top             =   1200
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image i11 
      Height          =   165
      Left            =   2160
      Picture         =   "frmmain.frx":0F68
      Top             =   1200
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000003&
      BackStyle       =   0  'Transparent
      Caption         =   "Over All Progress"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   840
      TabIndex        =   4
      Top             =   840
      Width           =   2595
   End
   Begin VB.Label porgress 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "Starting to Detect your Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   570
      TabIndex        =   3
      Top             =   120
      Width           =   3225
   End
   Begin VB.Image Image1 
      Height          =   330
      Left            =   1200
      Picture         =   "frmmain.frx":1086
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1890
   End
   Begin VB.Image i1 
      Height          =   165
      Left            =   960
      Picture         =   "frmmain.frx":2930
      Top             =   1200
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image i2 
      Height          =   165
      Left            =   1080
      Picture         =   "frmmain.frx":2A4E
      Top             =   1200
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image i3 
      Height          =   165
      Left            =   1200
      Picture         =   "frmmain.frx":2B6C
      Top             =   1200
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image i4 
      Height          =   165
      Left            =   1320
      Picture         =   "frmmain.frx":2C8A
      Top             =   1200
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image i5 
      Height          =   165
      Left            =   1440
      Picture         =   "frmmain.frx":2DA8
      Top             =   1200
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image i6 
      Height          =   165
      Left            =   1560
      Picture         =   "frmmain.frx":2EC6
      Top             =   1200
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image i7 
      Height          =   165
      Left            =   1680
      Picture         =   "frmmain.frx":2FE4
      Top             =   1200
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image i8 
      Height          =   165
      Left            =   1800
      Picture         =   "frmmain.frx":3102
      Top             =   1200
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image i9 
      Height          =   165
      Left            =   1920
      Picture         =   "frmmain.frx":3220
      Top             =   1200
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image i10 
      Height          =   165
      Left            =   2040
      Picture         =   "frmmain.frx":333E
      Top             =   1200
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image Image2 
      Height          =   330
      Left            =   840
      Picture         =   "frmmain.frx":345C
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   2610
   End
   Begin VB.Label countdown 
      BackStyle       =   0  'Transparent
      Caption         =   "21"
      Height          =   255
      Left            =   2760
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strtime As Integer
Private Sub Form_Load()
Timer1.Enabled = True
strtime = 0
End Sub

Private Sub Timer4_Timer()
If countdown.Caption > 0 Then
countdown.Caption = countdown.Caption - 1
End If
If countdown.Caption = 20 Then i1.Visible = True
If countdown.Caption = 19 Then i2.Visible = True
If countdown.Caption = 18 Then i3.Visible = True
If countdown.Caption = 17 Then i4.Visible = True
If countdown.Caption = 16 Then i5.Visible = True
If countdown.Caption = 15 Then i6.Visible = True
If countdown.Caption = 14 Then i7.Visible = True
If countdown.Caption = 13 Then i8.Visible = True
If countdown.Caption = 12 Then i9.Visible = True
If countdown.Caption = 11 Then i10.Visible = True
If countdown.Caption = 10 Then i11.Visible = True
If countdown.Caption = 9 Then i12.Visible = True
If countdown.Caption = 8 Then i13.Visible = True
If countdown.Caption = 7 Then i14.Visible = True
If countdown.Caption = 6 Then i15.Visible = True
If countdown.Caption = 5 Then i16.Visible = True
If countdown.Caption = 4 Then i17.Visible = True
If countdown.Caption = 3 Then i18.Visible = True
If countdown.Caption = 2 Then i19.Visible = True
If countdown.Caption = 1 Then i20.Visible = True
If i20.Visible = True Then Unload Me
End Sub
Private Sub Timer1_Timer()
If strtime = 5 Then
Call frmClear
Exit Sub
End If
picImage.Left = picImage.Left + 100
If picImage.Left = 3100 Then
Timer1.Enabled = False
Timer2.Enabled = True
End If
End Sub

Private Sub Timer2_Timer()
picImage.Left = picImage.Left - 100
If picImage.Left = 0 Then
Timer2.Enabled = False
Timer1.Enabled = True
End If
End Sub
Public Sub frmClear()
Unload Me
End Sub

