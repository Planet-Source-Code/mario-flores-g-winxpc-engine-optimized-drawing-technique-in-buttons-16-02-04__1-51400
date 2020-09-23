VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Demostration WindowsXPC 1.0"
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   555
   ClientWidth     =   11865
   FillColor       =   &H0097A5A5&
   ForeColor       =   &H00C4E2EF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   592
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   791
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      Height          =   615
      Left            =   8280
      TabIndex        =   30
      Top             =   0
      Width           =   3375
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "Form1.frx":0000
         Left            =   120
         List            =   "Form1.frx":000A
         TabIndex        =   32
         Text            =   "WindowsXP_Blue"
         Top             =   200
         Width           =   2415
      End
      Begin VB.CommandButton CmdApply 
         Caption         =   "Apply"
         Height          =   255
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   230
         Width           =   615
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Step"
      Height          =   375
      Left            =   10440
      TabIndex        =   24
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Inverse"
      Height          =   375
      Left            =   10440
      TabIndex        =   23
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stop Clock"
      Height          =   375
      Left            =   10440
      TabIndex        =   22
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start Clock"
      Height          =   375
      Left            =   10440
      TabIndex        =   20
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "ComboBox && ImageCombo && DriveListBox"
      Height          =   3495
      Left            =   240
      TabIndex        =   11
      Top             =   5280
      Width           =   4455
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   360
         TabIndex        =   28
         Top             =   480
         Width           =   1815
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   360
         TabIndex        =   15
         Text            =   "Combo3"
         Top             =   2520
         Width           =   2535
      End
      Begin VB.ComboBox Combo2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   360
         TabIndex        =   13
         Text            =   "Combo2"
         Top             =   1560
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   360
         TabIndex        =   12
         Text            =   "Combo1"
         Top             =   1080
         Width           =   1815
      End
      Begin MSComctlLib.ImageCombo ImageCombo1 
         Height          =   330
         Left            =   360
         TabIndex        =   14
         Top             =   2040
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Text            =   "ImageCombo1"
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "DriveList"
         Height          =   195
         Left            =   2400
         TabIndex        =   29
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Normal "
         Height          =   195
         Left            =   2400
         TabIndex        =   19
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Image Combo"
         Height          =   195
         Left            =   2400
         TabIndex        =   18
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Disabled"
         Height          =   195
         Left            =   2400
         TabIndex        =   17
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00825623&
         BackStyle       =   0  'Transparent
         Caption         =   "Different Size"
         Height          =   255
         Left            =   3240
         TabIndex        =   16
         Top             =   2880
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "ListView XP Style"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   4575
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   11535
      Begin MSComctlLib.ListView ListView1 
         Height          =   2595
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   10995
         _ExtentX        =   19394
         _ExtentY        =   4577
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Column1"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Column2"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Column3"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   1155
         Left            =   240
         TabIndex        =   10
         Top             =   3240
         Width           =   10995
         _ExtentX        =   19394
         _ExtentY        =   2037
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483630
         BackColor       =   -2147483643
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Handwriting"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Mario"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Alberto"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Flores"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Gonzalez"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ProgressBar:"
      Height          =   3495
      Left            =   4920
      TabIndex        =   1
      Top             =   5280
      Width           =   5055
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   1935
         Left            =   4200
         TabIndex        =   2
         Top             =   600
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   3413
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   1
         Max             =   25
         Orientation     =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar3 
         Height          =   615
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1085
         _Version        =   393216
         Appearance      =   1
         Max             =   25
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   1800
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Max             =   25
      End
      Begin MSComctlLib.ProgressBar ProgressBar4 
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   2640
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Max             =   25
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar5 
         Height          =   1935
         Left            =   3360
         TabIndex        =   27
         Top             =   600
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   3413
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   1
         Max             =   25
         Orientation     =   1
         Scrolling       =   1
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Smooth"
         Height          =   195
         Left            =   3210
         TabIndex        =   21
         Top             =   2640
         Width           =   555
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Smooth"
         Height          =   195
         Left            =   1080
         TabIndex        =   26
         Top             =   2880
         Width           =   540
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Vertical"
         Height          =   255
         Left            =   3960
         TabIndex        =   7
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00CED9D5&
         BackStyle       =   0  'Transparent
         Caption         =   "Different Size"
         Height          =   195
         Left            =   960
         TabIndex        =   6
         Top             =   1440
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Horizontal"
         Height          =   195
         Left            =   1080
         TabIndex        =   5
         Top             =   2160
         Width           =   705
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9480
      Top             =   5400
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test"
      Height          =   495
      Left            =   10200
      TabIndex        =   0
      Top             =   8280
      Width           =   1575
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "MAFG  WinXPC Open Source Engine 1.0.0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00825623&
      Height          =   195
      Left            =   570
      TabIndex        =   33
      Top             =   120
      Width           =   3585
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Option Explicit
 
Private timev As Integer
Private direction As Boolean '//---For Test


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
       Me.Hide
       Me.Show
  End If

End If
End Sub

Private Sub Combo4_Click()
WindowsXPC1.ColorScheme = Combo4.ListIndex + 1
End Sub

Private Sub Command1_Click()
If WindowsXPC1.EngineStarted = False Then
   WindowsXPC1.ColorScheme = Combo4.ListIndex + 1
   WindowsXPC1.InitSubClassing
   Me.Hide
   Me.Show
  
Else
   MsgBox "Engine Already Started!", vbInformation, "WinXPC Engine 1.0"
End If
End Sub



Private Sub Command2_Click()
Command6.Enabled = False
Timer1.Enabled = True
End Sub

Private Sub Command3_Click()
Command6.Enabled = True
Timer1.Enabled = False
End Sub

Private Sub Command5_Click()
direction = Not direction
End Sub

Private Sub Command6_Click()


If direction = False Then
    timev = timev + 1
Else
     timev = timev - 1
End If

If timev > ProgressBar1.Max Then timev = ProgressBar1.Min
If timev < ProgressBar1.Min Then timev = ProgressBar1.Max

ProgressBar1.Value = timev
ProgressBar2.Value = timev
ProgressBar3.Value = timev
ProgressBar4.Value = timev
ProgressBar5.Value = timev


End Sub

Private Sub Form_Initialize()
 Combo4.ListIndex = 1
End Sub

Private Sub Form_Load()
 Dim I As Integer
 Dim mRow As ListItem
    
  '---------------------------
  'initialize Listview Control
  '--------------------------
    ListView1.View = lvwReport
    ListView1.FullRowSelect = True
    
    For I = 1 To 12
     Combo1.AddItem VBA.MonthName(I)
     ImageCombo1.ComboItems.Add , , "Item " & I
    Next I
    
    For I = 0 To 40
      Set mRow = ListView1.ListItems.Add(, , CStr(I))
      Combo3.AddItem "Item " & I
      mRow.SubItems(1) = "This is Item " & I
      
    Next


 
End Sub


Private Sub Timer1_Timer()
ProgressBar1.Value = timev
ProgressBar2.Value = timev
ProgressBar3.Value = timev
ProgressBar4.Value = timev
ProgressBar5.Value = timev

If direction = False Then
    timev = timev + 1
Else
     timev = timev - 1
End If

If timev > ProgressBar1.Max Then timev = ProgressBar1.Min
If timev < ProgressBar1.Min Then timev = ProgressBar1.Max

End Sub

