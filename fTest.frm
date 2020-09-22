VERSION 5.00
Begin VB.Form fTest 
   Caption         =   "ucCaptionButton 1.0 test"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image2 
      Height          =   120
      Left            =   315
      Picture         =   "fTest.frx":0000
      Top             =   45
      Visible         =   0   'False
      Width           =   90
   End
   Begin Test.ucCaptionButton ucCaptionButton1 
      Left            =   4080
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Image Image1 
      Height          =   120
      Left            =   105
      Picture         =   "fTest.frx":006A
      Top             =   45
      Visible         =   0   'False
      Width           =   60
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    Call ucCaptionButton1.Caption_AddButton(Me.hWnd, 2, Image1, vbWhite)
    Call ucCaptionButton1.Caption_AddButton(Me.hWnd, 0, Image2, vbWhite, False)
End Sub

Private Sub ucCaptionButton1_ButtonClick(ByVal lhWnd As Long, ByVal lIndex As Long)

    Call MsgBox("Caption button #" & lIndex & " clicked!" & vbCrLf & _
                "(hWnd: " & lhWnd & ")")
End Sub
