VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\WinXPC Engine.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   7710
   StartUpPosition =   3  'Windows Default
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   2280
      Top             =   3960
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   435
      Left            =   2640
      TabIndex        =   3
      Top             =   960
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   767
      _Version        =   393216
      Max             =   300
      TickStyle       =   3
   End
   Begin VB.CommandButton Command2 
      Caption         =   "S&TOP"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&GO"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   2280
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   6480
      Top             =   2520
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Max             =   50
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim time

Private Sub Command1_Click()
Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
Timer1.Enabled = False
End Sub

Private Sub Form_Load()
Slider1.Min = 0
Slider1.Max = Timer1.Interval
Slider1.Value = Timer1.Interval

WindowsXPC1.InitSubClassing
End Sub

Private Sub Slider1_Change()
Timer1.Interval = Slider1.Value
End Sub



Private Sub Timer1_Timer()
time = time + 1
If time > ProgressBar1.Max Then time = ProgressBar1.Min
If time < ProgressBar1.Min Then time = ProgressBar1.Min

 ProgressBar1.Value = time
End Sub
