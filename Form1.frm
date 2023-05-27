VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ReMOD"
   ClientHeight    =   8865
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7200
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8865
   ScaleWidth      =   7200
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command5 
      Caption         =   "結束並開啟"
      Height          =   495
      Left            =   2040
      TabIndex        =   8
      Top             =   7800
      Width           =   2895
   End
   Begin VB.CommandButton Command4 
      Caption         =   "結束"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   8280
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "取消"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   8280
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "接受"
      Height          =   375
      Left            =   5640
      TabIndex        =   2
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  '平面
      Caption         =   "開始安裝"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2520
      TabIndex        =   0
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Left            =   6840
      Top             =   240
   End
   Begin VB.Label Label4 
      Caption         =   "安裝完成"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   27.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   7
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   3960
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000004&
      Caption         =   "安裝中...."
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   26.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   615
      Left            =   360
      TabIndex        =   5
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   0
      Picture         =   "Form1.frx":680A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '單線固定
      Caption         =   $"Form1.frx":C020
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.Image Image1 
      Height          =   1935
      Left            =   840
      Picture         =   "Form1.frx":C558
      Stretch         =   -1  'True
      Top             =   840
      Width           =   5415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SleepEx Lib "kernel32" _
    (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
Dim oldHeight, oldWidth As Integer

Private Sub Command1_Click()
Sleep 1000
Image1.Visible = False
Command1.Visible = False
Label1.Visible = True
Command2.Visible = True
Command3.Visible = True
Command4.Visible = True
Image2.Visible = True
End Sub

Private Sub Command2_Click()
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
Label1.Visible = False
Me.Visible = False
Sleep 1000
Me.Visible = True
Label2.Visible = True
Label3.Visible = True

End Sub

Private Sub Command3_Click()
MsgBox "ERROR:0931", vbCritical
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Command5_Click()
Me.Visible = False
Sleep 1500
MsgBox "您的系統不支援此應用", vbCritical
End
End Sub

Private Sub Form_Load()

Label2.Visible = False
Image2.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
Command5.Visible = False
Label2.Visible = False
Label4.Visible = False
Label3.Visible = False
' ??定?器
    Timer1.Interval = 50 ' 定?器?隔（以毫秒??位）
    Timer1.Enabled = True
   
oldHeight = Me.Height
oldWidth = Me.Width
End Sub
Private Sub Form_Resize()
Me.Height = oldHeight
Me.Width = oldWidth
End Sub




Private Sub Timer1_Timer()
If Not Label3.Caption = "100%" Then
If Label3.Visible = True Then
Sleep 2000
Label3.Caption = "100%"
End If
End If
If Label3.Visible = True Then
If Label3.Caption = "100%" Then
Label3.Visible = False
Label2.Visible = False
Sleep 1000
Label4.Visible = True
Command5.Visible = True
End If
End If
End Sub
