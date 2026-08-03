VERSION 5.00
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.2#0"; "VBCCR18.OCX"
Begin VB.Form FrmStart 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "列车时刻表"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6165
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   6165
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   700
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   6255
      TabIndex        =   3
      Top             =   3030
      Width           =   6255
      Begin VBCCR18.CommandButtonW CommandButtonExit 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   4680
         TabIndex        =   5
         Top             =   200
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "退出"
      End
      Begin VBCCR18.LinkLabel LinkLabelAbout 
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         Caption         =   "FrmStart.frx":0000
         Transparent     =   -1  'True
      End
   End
   Begin VBCCR18.CommandLink CommandLinkTrain 
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1085
      MouseTrack      =   -1  'True
      Caption         =   "FrmStart.frx":0048
      Transparent     =   -1  'True
   End
   Begin VBCCR18.CommandLink CommandLinkStation 
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1085
      MouseTrack      =   -1  'True
      Caption         =   "FrmStart.frx":0074
      Transparent     =   -1  'True
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "请选择列车的查询方式"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2400
   End
End
Attribute VB_Name = "FrmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long _
) As Long

Private Sub CommandButtonExit_Click()
    End
End Sub

Private Sub CommandLinkTrain_Click()
    FrmTrain.Show
    Unload Me
End Sub

Private Sub LinkLabelAbout_Click()
    ShellExecute Me.hwnd, "open", "https://crazyurus.com/", vbNullString, vbNullString, 1
End Sub
