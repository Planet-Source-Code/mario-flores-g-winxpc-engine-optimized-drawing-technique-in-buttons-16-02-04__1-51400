VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "*\A..\Step1\WinXPC Engine ActiveX Code\WinXPC Engine.vbp"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12690
   LinkTopic       =   "Form1"
   ScaleHeight     =   574
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   846
   StartUpPosition =   2  'CenterScreen
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   4440
      Top             =   3960
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   3480
      Top             =   7080
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   8880
      TabIndex        =   23
      Top             =   240
      Width           =   3375
      Begin VB.CommandButton CmdApply 
         Caption         =   "Apply"
         Height          =   265
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   230
         Width           =   615
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "Form1.frx":0000
         Left            =   120
         List            =   "Form1.frx":0010
         TabIndex        =   24
         Text            =   "WindowsXP_Blue"
         Top             =   200
         Width           =   2415
      End
   End
   Begin VB.CommandButton Command10 
      Caption         =   "DBase"
      Height          =   855
      Left            =   9720
      Picture         =   "Form1.frx":005A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Excel"
      Height          =   855
      Left            =   9720
      Picture         =   "Form1.frx":05E4
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Access"
      Height          =   855
      Left            =   9720
      Picture         =   "Form1.frx":0B6E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4440
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   240
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command8 
      Caption         =   "ShowColor"
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "InputBox"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "MsgBox2"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Custom Font"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   2
      Top             =   2160
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&MsgBox1"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   855
      Left            =   9720
      Picture         =   "Form1.frx":10F8
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3240
      Width           =   1695
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   4320
      TabIndex        =   26
      Top             =   7200
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "(New)"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   7200
      TabIndex        =   22
      Top             =   2400
      Width           =   435
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color dialog box + OfficeXP Style"
      ForeColor       =   &H00865724&
      Height          =   195
      Left            =   4800
      TabIndex        =   21
      Top             =   2400
      Width           =   2355
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Buttons Are Drawn when user uses the space bar to push button (Added KeyDown-KeyUp Support) "
      Height          =   195
      Left            =   1680
      TabIndex        =   20
      Top             =   5280
      Width           =   7080
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fixed:"
      ForeColor       =   &H00800080&
      Height          =   195
      Left            =   360
      TabIndex        =   19
      Top             =   5280
      Width           =   420
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":1682
      Height          =   435
      Left            =   1680
      TabIndex        =   18
      Top             =   4680
      Width           =   7200
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Optimization:"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   360
      TabIndex        =   17
      Top             =   4680
      Width           =   900
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Results: Less Resources, Less Time to Draw"
      ForeColor       =   &H00686868&
      Height          =   195
      Left            =   1680
      TabIndex        =   16
      Top             =   6600
      Width           =   3210
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<Caption && Bitmap && Focus Rect> are drawn by the System..                 "
      Height          =   195
      Left            =   1680
      TabIndex        =   15
      Top             =   6000
      Width           =   5025
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Change:"
      ForeColor       =   &H000040C0&
      Height          =   195
      Left            =   360
      TabIndex        =   14
      Top             =   6000
      Width           =   600
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CHANGES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0020A222&
      Height          =   555
      Left            =   360
      TabIndex        =   13
      Top             =   3720
      Width           =   2385
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bitmap Support For Buttons"
      Height          =   195
      Left            =   1680
      TabIndex        =   12
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Added:"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   360
      TabIndex        =   11
      Top             =   5640
      Width           =   510
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Buttons Now Support Images!!"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   9480
      TabIndex        =   10
      Top             =   8040
      Width           =   2160
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MessageBox && InputBox Now Support Scheme Colors (Themes)"
      ForeColor       =   &H00865724&
      Height          =   195
      Left            =   480
      TabIndex        =   6
      Top             =   840
      Width           =   4515
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long


Dim time


Private Sub CmdApply_Click()

If WindowsXPC1.EngineStarted = False Then
   
   WindowsXPC1.ColorScheme = Combo4.ListIndex + 1
   WindowsXPC1.InitSubClassing
   Me.Hide
   Me.Show
Else
   
  If MsgBox("The Engine Needs To Restart!!", vbInformation + vbOKCancel, "WinXPCEngine") = vbOK Then
       WindowsXPC1.ColorScheme = Combo4.ListIndex + 1
       WindowsXPC1.EndWinXPCSubClassing
       
       WindowsXPC1.InitSubClassing
       
       'Instead of Hide + Show
       RedrawWindow Me.hwnd, ByVal 0&, ByVal 0&, &H1
       'Me.Hide
       'Me.Show
  End If

End If

End Sub

Private Sub Timer1_Timer()
time = time + 20
If time > ProgressBar1.Max Then time = ProgressBar1.Min
If time < ProgressBar1.Min Then time = ProgressBar1.Min
ProgressBar1.Value = time
End Sub


Private Sub Combo4_Click()

WindowsXPC1.ColorScheme = Combo4.ListIndex + 1
End Sub

Private Sub Command3_Click()
MsgBox "WinXPC Engine", vbInformation, "Mario Flores"

End Sub

Private Sub Command6_Click()
MsgBox "WinXPC Engine By Mario Alberto Flores Gonzalez", vbYesNoCancel + vbInformation, "WinXPC"
End Sub

Private Sub Command7_Click()
InputBox "CD JUAREZ CHIHUAHUA....", "Example", "Mario Flores"
End Sub

Private Sub Command8_Click()
cd.ShowColor
End Sub

Private Sub Form_Load()
WindowsXPC1.InitSubClassing
End Sub

