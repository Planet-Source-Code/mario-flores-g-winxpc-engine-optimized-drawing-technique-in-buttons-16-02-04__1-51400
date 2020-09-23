VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7455
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00F4E9E8&
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00F4E9E8&
      Caption         =   "&Install"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6600
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F4E9E8&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   0
      ScaleHeight     =   3375
      ScaleWidth      =   7455
      TabIndex        =   5
      Top             =   2880
      Visible         =   0   'False
      Width           =   7455
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Installation Completed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2760
         TabIndex        =   6
         Top             =   1560
         Width           =   2325
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000080&
      Caption         =   "DEMO "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   0
      TabIndex        =   8
      Top             =   1920
      Width           =   7455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WinXPC Engine"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   615
      Left            =   1800
      TabIndex        =   7
      Top             =   360
      Width           =   4140
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "By Mario Flores"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   180
      Left            =   3240
      TabIndex        =   3
      Top             =   960
      Width           =   1140
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":72FA
      Height          =   1935
      Left            =   600
      TabIndex        =   1
      Top             =   4320
      Width           =   6735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "End User License Agreement (EULA)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2400
      TabIndex        =   0
      Top             =   3600
      Width           =   3135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E3E3E3&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   0
      Top             =   3480
      Width           =   7455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------------------------'
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
'                                        A SIMPLE REGISTER SHELL                                   '
'                                            Version 1.00                                          '
'                                                                                                  '
'                           AUTHOR:    MARIO ALBERTO FLORES GONZALEZ                               '
'                                                                                                  '
'                                     CD.JUAREZ CHIHUAHUA MEXICO                                   '
'                                                                                                  '
'                                    sistec_de_juarez@hotmail.com                                  '
'                                                                                                  '
'--------------------------------------------------------------------------------------------------'
'--------------------------------------------------------------------------------------------------'


'//--A SIMPLE REGISTER SHELL...AVOIDING USING regsvr32 MANUALLY...    ;)
                                                                              'Mario Flores G

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lParameter As Long, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Private Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Function RegDll(ByVal inFileSpec As String) As Boolean
On Error Resume Next

    Dim lLib As Long
    Dim lpDLLEntryPoint As Long
    Dim lpThreadID As Long
    Dim lpExitCode As Long
    Dim mThread
    
    lLib = LoadLibrary(inFileSpec)
    
    If lLib = 0 Then
        RegDll = False
        Exit Function
    End If
    
    lpDLLEntryPoint = GetProcAddress(lLib, "DllRegisterServer")
    
    
    If lpDLLEntryPoint = vbNull Then
        RegDll = False
        GoTo earlyExit1
    End If
    
        
    mThread = CreateThread(ByVal 0, 0, ByVal lpDLLEntryPoint, ByVal 0, 0, lpThreadID)
    
    If mThread = 0 Then
        RegDll = False
        GoTo earlyExit1
    End If
    
    mresult = WaitForSingleObject(mThread, 10000)
    
    If mresult <> 0 Then
        GoTo earlyExit2
    End If
    
    CloseHandle mThread
    FreeLibrary lLib
    
    RegDll = True
    Exit Function
    
    
earlyExit1:
    FreeLibrary lLib
    Exit Function
    
earlyExit2:
    FreeLibrary lLib
    lpExitCode = GetExitCodeThread(mThread, lpExitCode)
    ExitThread lpExitCode

End Function

Private Sub Command1_Click()
 Dim Path As String, strSave As String

 strSave = String(200, Chr$(0))
 Path = Left$(strSave, GetWindowsDirectory(strSave, Len(strSave))) + "\System32\"
 
 'Call VBA.FileCopy(App.Path & "\Dll\WinXPC Engine.ocx", Path & "WinXPC Engine.ocx")
 
 Call RegDll(Path & "WinXPC Engine.ocx")
 
 


 
 
 Command2.Caption = "&Exit"
 Picture1.Visible = True
 Command1.Enabled = False
End Sub

Private Sub Command2_Click()
Unload Me
End Sub



