VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\WinXPC Engine.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   9135
   StartUpPosition =   3  'Windows Default
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   2760
      Top             =   3600
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   4350
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Text            =   "Simple Example "
            TextSave        =   "Simple Example "
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Mute All"
      Height          =   375
      Left            =   6600
      TabIndex        =   14
      Top             =   3120
      Width           =   975
   End
   Begin MSComctlLib.Slider Slider8 
      Height          =   1455
      Left            =   7440
      TabIndex        =   13
      Top             =   1560
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   2566
      _Version        =   393216
      Orientation     =   1
      Max             =   6
      SelStart        =   3
      TickStyle       =   2
      Value           =   3
   End
   Begin MSComctlLib.Slider Slider7 
      Height          =   495
      Left            =   7320
      TabIndex        =   12
      Top             =   720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      _Version        =   393216
      Enabled         =   0   'False
      Max             =   2
      SelStart        =   1
      Value           =   1
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Mute All"
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   3120
      Width           =   975
   End
   Begin MSComctlLib.Slider Slider6 
      Height          =   1455
      Left            =   5280
      TabIndex        =   9
      Top             =   1560
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   2566
      _Version        =   393216
      Orientation     =   1
      Max             =   6
      SelStart        =   3
      TickStyle       =   2
      Value           =   3
   End
   Begin MSComctlLib.Slider Slider5 
      Height          =   495
      Left            =   5160
      TabIndex        =   8
      Top             =   720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      _Version        =   393216
      Max             =   2
      SelStart        =   1
      Value           =   1
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Mute All"
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   3120
      Width           =   975
   End
   Begin MSComctlLib.Slider Slider4 
      Height          =   1455
      Left            =   3240
      TabIndex        =   5
      Top             =   1560
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   2566
      _Version        =   393216
      Orientation     =   1
      Max             =   6
      SelStart        =   3
      TickStyle       =   2
      Value           =   3
   End
   Begin MSComctlLib.Slider Slider3 
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      _Version        =   393216
      Max             =   2
      SelStart        =   1
      Value           =   1
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Mute All"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   3120
      Width           =   975
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   1455
      Left            =   1200
      TabIndex        =   1
      Top             =   1560
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   2566
      _Version        =   393216
      Orientation     =   1
      Max             =   6
      SelStart        =   3
      TickStyle       =   2
      Value           =   3
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      _Version        =   393216
      Max             =   2
      SelStart        =   1
      Value           =   1
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Volume:"
      Height          =   195
      Left            =   6600
      TabIndex        =   15
      Top             =   1320
      Width           =   570
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   7
      Left            =   6960
      Top             =   840
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   6
      Left            =   8280
      Top             =   840
      Width           =   240
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Volume:"
      Height          =   195
      Left            =   4440
      TabIndex        =   11
      Top             =   1320
      Width           =   570
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   5
      Left            =   4800
      Top             =   840
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   4
      Left            =   6120
      Top             =   840
      Width           =   240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Volume:"
      Height          =   195
      Left            =   2400
      TabIndex        =   7
      Top             =   1320
      Width           =   570
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   3
      Left            =   2760
      Top             =   840
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   2
      Left            =   4080
      Top             =   840
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Volume:"
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   570
   End
   Begin VB.Image Pic1 
      Height          =   480
      Left            =   240
      Picture         =   "Form1.frx":0000
      Top             =   360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Pic2 
      Height          =   480
      Left            =   1560
      Picture         =   "Form1.frx":030A
      Top             =   360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   1
      Left            =   720
      Top             =   840
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   0
      Left            =   2040
      Top             =   840
      Width           =   240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim I As Integer

Private Sub Form_Load()
For I = Image1.LBound To Image1.UBound
 Image1(I) = IIf(I = 0 Or I = 2 Or I = 4 Or I = 6, Pic2, Pic1)
Next I

WindowsXPC1.InitSubClassing
End Sub
