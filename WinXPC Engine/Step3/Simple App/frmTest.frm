VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTest 
   Caption         =   " WindowsXPC"
   ClientHeight    =   8775
   ClientLeft      =   2865
   ClientTop       =   2175
   ClientWidth     =   12885
   FillColor       =   &H00E9F1F2&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000C0&
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   12885
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      Height          =   615
      Left            =   9360
      TabIndex        =   44
      Top             =   0
      Width           =   3375
      Begin VB.CommandButton CmdApply 
         Caption         =   "Apply"
         Height          =   255
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   230
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmTest.frx":000C
         Left            =   120
         List            =   "frmTest.frx":0016
         TabIndex        =   45
         Text            =   "WindowsXP_Blue"
         Top             =   200
         Width           =   2415
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   8280
      Width           =   12885
      _ExtentX        =   22728
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   3387
            Text            =   " WindowsXPC Engine 1.0 "
            TextSave        =   " WindowsXPC Engine 1.0 "
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   4048
            Text            =   "      Author: Mario Flores G       "
            TextSave        =   "      Author: Mario Flores G       "
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   5503
            Text            =   "   email-me sistec_de_juarez@hotmail.com"
            TextSave        =   "   email-me sistec_de_juarez@hotmail.com"
         EndProperty
      EndProperty
   End
   Begin VB.Frame FrameAll 
      Height          =   7450
      Index           =   5
      Left            =   240
      TabIndex        =   48
      Top             =   720
      Visible         =   0   'False
      Width           =   12375
      Begin VB.CommandButton Command6 
         Caption         =   "InputBox"
         Height          =   495
         Left            =   8280
         TabIndex        =   53
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CommandButton Command5 
         Caption         =   "MsgBox"
         Height          =   495
         Left            =   5280
         TabIndex        =   50
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CommandButton Command4 
         Caption         =   "MsgBox"
         Height          =   495
         Left            =   2040
         TabIndex        =   49
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Turn On The Engine Fisrt!!!"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   4560
         TabIndex        =   55
         Top             =   5640
         Width           =   2970
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Regular InputBox"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A85E31&
         Height          =   195
         Left            =   8400
         TabIndex        =   54
         Top             =   960
         Width           =   1500
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Regular Yes-No-Cancel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A85E31&
         Height          =   195
         Left            =   5280
         TabIndex        =   52
         Top             =   960
         Width           =   1980
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Regular MsgBox"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A85E31&
         Height          =   195
         Left            =   2280
         TabIndex        =   51
         Top             =   960
         Width           =   1380
      End
   End
   Begin VB.Frame FrameAll 
      BackColor       =   &H00FFFFFF&
      Caption         =   "About"
      Height          =   6015
      Index           =   2
      Left            =   2040
      TabIndex        =   41
      Top             =   1320
      Width           =   8895
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "web.1asphost.com\marioflores\WinXPc.htm"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   2520
         TabIndex        =   73
         Top             =   2160
         Width           =   3735
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OCX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   4440
         TabIndex        =   72
         Top             =   3360
         Width           =   330
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "sistec_de_juarez@hotmail.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   3360
         TabIndex        =   71
         Top             =   5040
         Width           =   2655
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version 1.0.0"
         ForeColor       =   &H00825623&
         Height          =   195
         Left            =   4080
         TabIndex        =   70
         Top             =   720
         Width           =   960
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mario Alberto Flores Gonzalez"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3360
         TabIndex        =   69
         Top             =   1800
         Width           =   2520
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmTest.frx":0038
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00825623&
         Height          =   1035
         Left            =   3360
         TabIndex        =   68
         Top             =   480
         Width           =   2490
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00825623&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0088C393&
         Height          =   495
         Left            =   -240
         Top             =   3240
         Width           =   9135
      End
   End
   Begin VB.Frame FrameAll 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A85E31&
      Height          =   7455
      Index           =   1
      Left            =   240
      TabIndex        =   63
      Top             =   720
      Width           =   12375
      Begin VB.Image Image2 
         Height          =   6150
         Left            =   2520
         Top             =   480
         Width           =   7350
      End
   End
   Begin VB.Frame FrameAll 
      BorderStyle     =   0  'None
      Height          =   7370
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   12255
      Begin VB.Frame Frame6 
         Height          =   380
         Left            =   5280
         TabIndex        =   74
         Top             =   6930
         Width           =   2055
         Begin VB.Image Image1 
            Height          =   225
            Left            =   1560
            MouseIcon       =   "frmTest.frx":00C4
            MousePointer    =   99  'Custom
            Picture         =   "frmTest.frx":0216
            Top             =   120
            Width           =   225
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H00B99D7F&
            BackStyle       =   0  'Transparent
            Caption         =   "More Controls:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00DD5508&
            Height          =   210
            Left            =   240
            TabIndex        =   75
            Top             =   120
            Width           =   1065
         End
      End
      Begin VB.CommandButton CommandEmulate 
         Caption         =   "&Test Me"
         Height          =   495
         Left            =   10440
         TabIndex        =   64
         Top             =   600
         Width           =   1095
      End
      Begin VB.Frame Frame8 
         Caption         =   "PictureBox"
         Height          =   1575
         Left            =   4920
         TabIndex        =   56
         Top             =   240
         Width           =   4455
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   735
            Left            =   3000
            ScaleHeight     =   705
            ScaleWidth      =   1185
            TabIndex        =   59
            Top             =   360
            Width           =   1215
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00FFFFFF&
            Height          =   735
            Left            =   1680
            ScaleHeight     =   675
            ScaleWidth      =   1155
            TabIndex        =   58
            Top             =   360
            Width           =   1215
         End
         Begin VB.PictureBox Picture1 
            Enabled         =   0   'False
            Height          =   735
            Left            =   240
            ScaleHeight     =   675
            ScaleWidth      =   1275
            TabIndex        =   57
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Flat"
            Height          =   195
            Left            =   3480
            TabIndex        =   62
            Top             =   1200
            Width           =   270
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Disabled"
            Height          =   195
            Left            =   600
            TabIndex        =   61
            Top             =   1200
            Width           =   600
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Normal "
            Height          =   255
            Left            =   2040
            TabIndex        =   60
            Top             =   1200
            Width           =   615
         End
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Enable All Controls"
         Height          =   495
         Left            =   10200
         TabIndex        =   47
         Top             =   1440
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.Frame Frame5 
         Caption         =   "Slider"
         Height          =   2055
         Left            =   120
         TabIndex        =   3
         Top             =   4800
         Width           =   12015
         Begin MSComctlLib.Slider Slider1 
            Height          =   495
            Left            =   960
            TabIndex        =   4
            Top             =   720
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   873
            _Version        =   393216
            Max             =   5
            SelStart        =   5
            Value           =   5
         End
         Begin MSComctlLib.Slider Slider4 
            Height          =   1455
            Left            =   11280
            TabIndex        =   5
            Top             =   360
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   2566
            _Version        =   393216
            Orientation     =   1
            Max             =   3
            SelStart        =   3
            Value           =   3
            TextPosition    =   1
         End
         Begin MSComctlLib.Slider Slider5 
            Height          =   1455
            Left            =   9480
            TabIndex        =   6
            Top             =   360
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   2566
            _Version        =   393216
            Orientation     =   1
            SelStart        =   10
            TickStyle       =   1
            Value           =   10
         End
         Begin MSComctlLib.Slider Slider6 
            Height          =   1455
            Left            =   10320
            TabIndex        =   7
            Top             =   360
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   2566
            _Version        =   393216
            Orientation     =   1
            SelStart        =   10
            TickStyle       =   2
            Value           =   10
         End
         Begin MSComctlLib.Slider Slider7 
            Height          =   375
            Left            =   4440
            TabIndex        =   8
            Top             =   1560
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Max             =   100
            SelStart        =   90
            TickStyle       =   3
            Value           =   90
         End
         Begin MSComctlLib.Slider Slider8 
            Height          =   1455
            Left            =   8760
            TabIndex        =   9
            Top             =   360
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   2566
            _Version        =   393216
            Orientation     =   1
            SelStart        =   10
            TickStyle       =   3
            Value           =   10
            TextPosition    =   1
         End
         Begin MSComctlLib.Slider Slider9 
            Height          =   495
            Left            =   960
            TabIndex        =   10
            Top             =   1440
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   873
            _Version        =   393216
            TickStyle       =   1
            TextPosition    =   1
         End
         Begin MSComctlLib.Slider Slider10 
            Height          =   615
            Left            =   4440
            TabIndex        =   11
            Top             =   600
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   1085
            _Version        =   393216
            SelStart        =   6
            TickStyle       =   2
            Value           =   6
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "No Ticks"
            Height          =   195
            Left            =   8640
            TabIndex        =   21
            Top             =   1800
            Width           =   585
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Left"
            Height          =   195
            Left            =   9600
            TabIndex        =   20
            Top             =   1800
            Width           =   285
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Both"
            Height          =   195
            Left            =   10440
            TabIndex        =   19
            Top             =   1800
            Width           =   330
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Right"
            Height          =   195
            Left            =   11280
            TabIndex        =   18
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Vertical:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   7800
            TabIndex        =   17
            Top             =   360
            Width           =   690
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Horizontal:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   480
            TabIndex        =   16
            Top             =   360
            Width           =   915
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Disabled"
            Height          =   195
            Left            =   3720
            TabIndex        =   15
            Top             =   1560
            Width           =   600
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Both"
            Height          =   195
            Left            =   3840
            TabIndex        =   14
            Top             =   840
            Width           =   330
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Left"
            Height          =   195
            Left            =   480
            TabIndex        =   13
            Top             =   1560
            Width           =   285
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Right"
            Height          =   195
            Left            =   480
            TabIndex        =   12
            Top             =   840
            Width           =   375
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "OptionButton"
         Height          =   2535
         Left            =   3120
         TabIndex        =   27
         Top             =   2160
         Width           =   2295
         Begin VB.OptionButton Option7 
            Caption         =   "Normal State"
            Height          =   255
            Left            =   360
            TabIndex        =   67
            Top             =   480
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Same Font"
            BeginProperty Font 
               Name            =   "Lucida Handwriting"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   66
            Top             =   960
            Width           =   1575
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Disabled State"
            Enabled         =   0   'False
            Height          =   255
            Left            =   360
            TabIndex        =   29
            Top             =   1440
            Width           =   1575
         End
         Begin VB.OptionButton Option3 
            Alignment       =   1  'Right Justify
            Caption         =   "Different Alignment"
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   1920
            Width           =   1815
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "TextBox"
         Height          =   2535
         Left            =   120
         TabIndex        =   22
         Top             =   2160
         Width           =   2895
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            Height          =   495
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   42
            Text            =   "frmTest.frx":02F5
            Top             =   1560
            Width           =   2415
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   960
            TabIndex        =   24
            Text            =   "Text1"
            Top             =   450
            Width           =   1695
         End
         Begin VB.TextBox Text2 
            Enabled         =   0   'False
            Height          =   285
            Left            =   960
            TabIndex        =   23
            Text            =   "Text2"
            Top             =   1050
            Width           =   1695
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "MultiLine"
            Height          =   195
            Left            =   1200
            TabIndex        =   43
            Top             =   2160
            Width           =   615
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Disabled"
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Normal "
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   480
            Width           =   615
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Command Buttons:"
         Height          =   2535
         Left            =   5640
         TabIndex        =   34
         Top             =   2160
         Width           =   6135
         Begin VB.CommandButton Command2 
            Caption         =   "Command2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   3960
            TabIndex        =   37
            Top             =   480
            Width           =   1935
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Command1"
            Enabled         =   0   'False
            Height          =   495
            Left            =   600
            TabIndex        =   36
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CommandButton Command3 
            Caption         =   "MArio FLores"
            BeginProperty Font 
               Name            =   "Viner Hand ITC"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2160
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Big Size"
            Height          =   195
            Left            =   4560
            TabIndex        =   40
            Top             =   2160
            Width           =   540
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Disabled State"
            Height          =   195
            Left            =   720
            TabIndex        =   39
            Top             =   2160
            Width           =   1035
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Different Fonts"
            Height          =   195
            Left            =   2400
            TabIndex        =   38
            Top             =   2160
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "CheckBox"
         Height          =   1575
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   4335
         Begin VB.CheckBox Check4 
            Alignment       =   1  'Right Justify
            Caption         =   "Different Alignment"
            Height          =   375
            Left            =   2280
            TabIndex        =   65
            Top             =   960
            Width           =   1815
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Normal State"
            Height          =   375
            Left            =   480
            TabIndex        =   33
            Top             =   360
            Width           =   1335
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Grayed State"
            Height          =   375
            Left            =   480
            TabIndex        =   32
            Top             =   960
            Value           =   2  'Grayed
            Width           =   1455
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Disabled State"
            Enabled         =   0   'False
            Height          =   375
            Left            =   2280
            TabIndex        =   31
            Top             =   360
            Width           =   1335
         End
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   7900
      Left            =   120
      TabIndex        =   0
      Top             =   340
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   13944
      TabWidthStyle   =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   " Test Engine "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Compatibility"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "About WindowsXPC Engine "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Revisions"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Blank"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Extra"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit




Private Sub Check5_Click()
Dim TempCont As Control
On Error Resume Next

For Each TempCont In Me.Controls
    If TempCont.Name <> Check5.Name And TempCont.Name <> FrameAll(0).Name Then TempCont.Enabled = IIf(Check5.Value = 0, False, True)
Next TempCont


End Sub


Private Sub CmdApply_Click()
If WindowsXPC1.EngineStarted = False Then
   
   WindowsXPC1.ColorScheme = Combo1.ListIndex + 1
   WindowsXPC1.InitSubClassing
   Me.Hide
   Me.Show
Else
   
  If MsgBox("The Engine Needs To Restart!!", vbInformation + vbOKCancel, "WinXPC") = vbOK Then
       WindowsXPC1.EndWinXPCSubClassing
       WindowsXPC1.ColorScheme = Combo1.ListIndex + 1
       WindowsXPC1.InitSubClassing
       Me.Hide
       Me.Show
  End If

End If

End Sub

Private Sub Combo1_Click()
WindowsXPC1.ColorScheme = Combo1.ListIndex + 1
End Sub

  
    
Private Sub Command4_Click()
MsgBox "Button in XP Style", vbInformation, "Test"
End Sub

Private Sub Command5_Click()
MsgBox "XP Buttons Style Very Cool!!!", vbYesNoCancel, "Test"
End Sub

Private Sub Command6_Click()
InputBox "WinXPC Engine V1.0 ", "Test ", "By Mario Flores G"
End Sub



Private Sub CommandEmulate_Click()

If WindowsXPC1.EngineStarted = False Then
   WindowsXPC1.ColorScheme = Combo1.ListIndex + 1
   WindowsXPC1.InitSubClassing
   Me.Hide
   Me.Show
  
Else
   MsgBox "Engine Already Started!", vbInformation, "WinXPC 1.0"
End If
   
   
End Sub




Private Sub Form_Initialize()
    
    Combo1.ListIndex = 1
    FrameAll(0).Visible = True
    FrameAll(1).Visible = False
    FrameAll(2).Visible = False
End Sub

Private Sub lblColorBit_Click(Index As Integer)
MsgBox TabStrip1.Tabs(TabStrip1.SelectedItem.Index).Height
   
End Sub



Private Sub Image1_Click()
Form1.Show
End Sub

Private Sub Slider1_Click()
'Command4.Left = Slider1.Left + ((Slider1.Width - 400) * (Slider1.Value / Slider1.Max))

Debug.Print (Slider1.Width - 15) * (Slider1.Value / Slider1.Max)
End Sub



Private Sub TabStrip1_Click()
'tabstrip1.Font.
FrameAll(0).Visible = False
FrameAll(1).Visible = False
FrameAll(2).Visible = False
FrameAll(5).Visible = False

Select Case TabStrip1.SelectedItem.Index - 1

Case 0, 1, 2, 5
FrameAll(TabStrip1.SelectedItem.Index - 1).Visible = True


End Select

End Sub

