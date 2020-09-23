VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\WinXPC Engine.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   1560
      Top             =   5760
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   5400
      TabIndex        =   4
      Top             =   4560
      Width           =   1215
   End
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   360
      TabIndex        =   3
      Top             =   3720
      Width           =   4335
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   5040
      TabIndex        =   2
      Top             =   2400
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   5040
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3135
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   5530
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Mario"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Flores"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Dim I As Integer
Dim mRow As ListItem
    
  '---------------------------
  'initialize Listview Control
  '--------------------------
    ListView1.View = lvwReport
    ListView1.FullRowSelect = True
    
       
    For I = 0 To 40
      List1.AddItem I & ".- Item"
      Set mRow = ListView1.ListItems.Add(, , CStr(I))
      mRow.SubItems(1) = "This is Item " & I
    Next


WindowsXPC1.InitSubClassing


End Sub
