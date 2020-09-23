VERSION 5.00
Begin VB.UserControl WindowsXPC 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3735
   InvisibleAtRuntime=   -1  'True
   MaskColor       =   &H00FFFFFF&
   PropertyPages   =   "WindowsXPC.ctx":0000
   ScaleHeight     =   615
   ScaleWidth      =   3735
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WinXPC Engine 1.0"
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
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   1680
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00825623&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   3735
   End
End
Attribute VB_Name = "WindowsXPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private m_iCount As Long
Private m_SubClassedItem() As cWinXPCEngine

Private m_IDECount As Long
Private m_IDEItem() As cWinXPCEngine

Public Enum SchemeWindowColors
       System = 1
       XP_Blue = 2
       XP_OliveGreen = 3
       XP_Silver = 4
End Enum

Private m_EngineStarted As Boolean
Private m_IDEStarted As Boolean


Private m_IDE As Boolean
Private StartedEngine As Boolean
Private RunningApp As Boolean
Private i As Long

Private m_ColorScheme         As CWindowColors

Private m_FileListBoxControl  As Boolean
Private m_DirListBoxControl   As Boolean
Private m_ListBoxControl      As Boolean
Private m_ListViewControl     As Boolean
Private m_ImageComboControl   As Boolean
Private m_SliderControl       As Boolean
Private m_ProgressBarControl  As Boolean
Private m_StatusBarControl    As Boolean
Private m_TabStripControl     As Boolean
Private m_DriveListBoxControl As Boolean
Private m_ComboBoxControl     As Boolean
Private m_OptionControl       As Boolean
Private m_CheckControl        As Boolean
Private m_ButtonControl       As Boolean
Private m_FrameControl        As Boolean
Private m_PictureControl      As Boolean
Private m_TextControl         As Boolean
Private m_MsgBox_InputBox     As Boolean
Private m_CommonDialog        As Boolean


Private Sub Timer1_Timer()
If UserControl.Ambient.UserMode And Me.EngineStarted = False Then InitSubClassing

Exit Sub
End Sub

Private Sub UserControl_InitProperties()
    m_ColorScheme = WindowsXP_Blue '//--DefaultColors
    m_ListBoxControl = True
    m_DirListBoxControl = True
    m_FileListBoxControl = True
    m_ListViewControl = True
    m_ImageComboControl = True
    m_SliderControl = True
    m_ProgressBarControl = True
    m_StatusBarControl = True
    m_TabStripControl = True
    m_DriveListBoxControl = True
    m_ComboBoxControl = True
    m_OptionControl = True
    m_CheckControl = True
    m_ButtonControl = True
    m_FrameControl = True
    m_PictureControl = True
    m_TextControl = True
    m_MsgBox_InputBox = True
    m_CommonDialog = False
    
End Sub


Public Sub InitSubClassing()
Dim SubclassThis As Boolean
Dim aControl As Control
  
'If Not UserControl.Ambient.UserMode Then Exit Sub

For Each aControl In UserControl.Parent.Controls
     
  SubclassThis = False

  On Error Resume Next
  If Err.Number = 0 Then

     
     Select Case ThisWindowClassName(aControl.hwnd)
            
            Case "ThunderListBox", "ThunderRT6ListBox"
                  If m_ListBoxControl Then SubclassThis = True
            Case "ThunderCommandButton", "ThunderRT6CommandButton", "Button"
                  If m_ButtonControl Then SubclassThis = True
            Case "ThunderFrame", "ThunderRT6Frame"
                  If m_FrameControl Then SubclassThis = True
            Case "ThunderOptionButton", "ThunderRT6OptionButton"
                  If m_OptionControl Then SubclassThis = True
            Case "Slider20WndClass"
                  If m_SliderControl Then SubclassThis = True
            Case "ThunderTextBox", "ThunderRT6TextBox", "Edit"
                  If m_TextControl Then SubclassThis = True
            Case "ThunderCheckBox", "ThunderRT6CheckBox"
                  If m_CheckControl Then SubclassThis = True
            Case "TabStrip20WndClass", "TabStripWndClass"
                  If m_TabStripControl Then SubclassThis = True
            Case "ThunderComboBox", "ThunderRT6ComboBox", "ComboBox"
                  If m_ComboBoxControl Then SubclassThis = True
            Case "ImageCombo20WndClass"
                  If m_ImageComboControl Then SubclassThis = True
            Case "ProgressBar20WndClass"
                  If m_ProgressBarControl Then SubclassThis = True
            Case "ListView20WndClass"
                  If m_ListViewControl Then SubclassThis = True
            Case "StatusBar20WndClass"
                  If m_StatusBarControl Then SubclassThis = True
            Case "ThunderDirListBox", "ThunderRT6DirListBox"
                  If m_DirListBoxControl Then SubclassThis = True
            Case "ThunderDriveListBox", "ThunderRT6DriveListBox"
                  If m_DriveListBoxControl Then SubclassThis = True
            Case "ThunderFileListBox", "ThunderRT6FileListBox"
                  If m_FileListBoxControl Then SubclassThis = True
            Case "ThunderPictureBox", "ThunderPictureBoxDC", "ThunderRT6PictureBoxDC"
                  If m_PictureControl Then SubclassThis = True
            Case Else
                  'Nothing
                   Debug.Print ThisWindowClassName(aControl.hwnd)
     End Select
     
       If TypeName(aControl) = "Adodc" Then SubclassThis = True
    

     If (SubclassThis) Then
         
         m_EngineStarted = True
         m_iCount = m_iCount + 1
         ReDim Preserve m_SubClassedItem(1 To m_iCount) As cWinXPCEngine
         Set m_SubClassedItem(m_iCount) = New cWinXPCEngine
         m_SubClassedItem(m_iCount).SchemeColor = m_ColorScheme
         m_SubClassedItem(m_iCount).ActiveScaleMode = UserControl.Parent.ScaleMode
         m_SubClassedItem(m_iCount).IdeSubClass = m_IDE
         m_SubClassedItem(m_iCount).Attach aControl
         
         
         
     End If
 
 
 End If
 
Next aControl '//-- Each Control
       
       If m_MsgBox_InputBox Then
         m_EngineStarted = True
         m_iCount = m_iCount + 1
         ReDim Preserve m_SubClassedItem(1 To m_iCount) As cWinXPCEngine
         Set m_SubClassedItem(m_iCount) = New cWinXPCEngine
         m_SubClassedItem(m_iCount).SchemeColor = m_ColorScheme
         m_SubClassedItem(m_iCount).IdeSubClass = m_IDE
         m_SubClassedItem(m_iCount).BeforeAttachMessageBox UserControl.Parent.hwnd
        
       End If
       
       If m_CommonDialog Then
         m_EngineStarted = True
         m_iCount = m_iCount + 1
         ReDim Preserve m_SubClassedItem(1 To m_iCount) As cWinXPCEngine
         Set m_SubClassedItem(m_iCount) = New cWinXPCEngine
         m_SubClassedItem(m_iCount).SchemeColor = m_ColorScheme
         m_SubClassedItem(m_iCount).IdeSubClass = m_IDE
         m_SubClassedItem(m_iCount).BeforeAttachCommonDialog UserControl.Parent.hwnd
       End If

      
       SetProp UserControl.Parent.hwnd, "ColorScheme", m_ColorScheme
       
End Sub


Public Sub InitIDESubClassing()
Dim SubclassThis As Boolean
Dim aControl As Control
 
For Each aControl In UserControl.Parent.Controls
     
  SubclassThis = False

  On Error Resume Next
  If Err.Number = 0 Then

     
     Select Case ThisWindowClassName(aControl.hwnd)
            
            Case "ThunderListBox"
                  SubclassThis = True
            Case "ThunderCommandButton", "Button"
                  SubclassThis = True
            Case "ThunderFrame"
                  SubclassThis = True
            Case "ThunderOptionButton"
                  SubclassThis = True
            Case "Slider20WndClass"
                  SubclassThis = True
            Case "ThunderTextBox", "Edit"
                  SubclassThis = True
            Case "ThunderCheckBox"
                  SubclassThis = True
            Case "ThunderComboBox", "ComboBox"
                  SubclassThis = True
            Case "ImageCombo20WndClass"
                  SubclassThis = True
            Case "ProgressBar20WndClass"
                  SubclassThis = True
            Case "ThunderDirListBox"
                  SubclassThis = True
            Case "ThunderDriveListBox"
                  SubclassThis = True
            Case "ThunderFileListBox"
                  SubclassThis = True
            Case "ThunderPictureBox", "ThunderPictureBoxDC"
                  SubclassThis = True
            Case Else
                  'Nothing
                  
     End Select
     

     If (SubclassThis) Then
         
         m_IDEStarted = True
         m_IDECount = m_IDECount + 1
         ReDim Preserve m_IDEItem(1 To m_IDECount) As cWinXPCEngine
         Set m_IDEItem(m_IDECount) = New cWinXPCEngine
         m_IDEItem(m_IDECount).SchemeColor = m_ColorScheme
         m_IDEItem(m_IDECount).Attach aControl
         
     End If
 
 
 End If
 
Next aControl '//-- Each Control


   
End Sub


Public Sub EndWinXPCSubClassing()

      If m_EngineStarted = True Then
      
            For i = 1 To m_iCount
                m_SubClassedItem(i).UnloadEngine
                Set m_SubClassedItem(i) = Nothing
                
            Next i
            m_iCount = 0
            
         EngineStarted = False
         
       End If
       

End Sub

Public Property Get ColorScheme() As SchemeWindowColors
   ColorScheme = m_ColorScheme
End Property

Public Property Let ColorScheme(ByVal cColorScheme As SchemeWindowColors)
   m_ColorScheme = cColorScheme
   PropertyChanged "ColorScheme"
End Property

Public Property Get DirListBoxControl() As Boolean
   DirListBoxControl = m_DirListBoxControl
End Property

Public Property Let DirListBoxControl(ByVal cDirListBoxControl As Boolean)
   m_DirListBoxControl = cDirListBoxControl
   PropertyChanged "DirListBoxControl"
End Property

Public Property Get FileListBoxControl() As Boolean
   FileListBoxControl = m_FileListBoxControl
End Property

Public Property Let FileListBoxControl(ByVal cFileListBoxControl As Boolean)
   m_FileListBoxControl = cFileListBoxControl
   PropertyChanged "FileListBoxControl"
End Property

Public Property Get ListBoxControl() As Boolean
   ListBoxControl = m_ListBoxControl
End Property

Public Property Let ListBoxControl(ByVal cListBoxControl As Boolean)
   m_ListBoxControl = cListBoxControl
   PropertyChanged "ListBoxControl"
End Property

Public Property Get ListViewControl() As Boolean
   ListViewControl = m_ListViewControl
End Property

Public Property Let ListViewControl(ByVal cListViewControl As Boolean)
   m_ListViewControl = cListViewControl
   PropertyChanged "ListViewControl"
End Property

Public Property Get ImageComboControl() As Boolean
   ImageComboControl = m_ImageComboControl
End Property

Public Property Let ImageComboControl(ByVal cImageComboControl As Boolean)
   m_ImageComboControl = cImageComboControl
   PropertyChanged "ImageComboControl"
End Property

Public Property Get SliderControl() As Boolean
   SliderControl = m_SliderControl
End Property

Public Property Let SliderControl(ByVal cSliderControl As Boolean)
   m_SliderControl = cSliderControl
   PropertyChanged "SliderControl"
End Property

Public Property Get ProgressBarControl() As Boolean
   ProgressBarControl = m_ProgressBarControl
End Property

Public Property Let ProgressBarControl(ByVal cProgressBarControl As Boolean)
   m_ProgressBarControl = cProgressBarControl
   PropertyChanged "ProgressBarControl"
End Property

Public Property Get StatusBarControl() As Boolean
   StatusBarControl = m_StatusBarControl
End Property

Public Property Let StatusBarControl(ByVal cStatusBarControl As Boolean)
   m_StatusBarControl = cStatusBarControl
   PropertyChanged "StatusBarControl"
End Property

Public Property Get TabStripControl() As Boolean
   TabStripControl = m_TabStripControl
End Property

Public Property Let TabStripControl(ByVal cTabStripControl As Boolean)
   m_TabStripControl = cTabStripControl
   PropertyChanged "TabStripControl"
End Property

Public Property Get DriveListBoxControl() As Boolean
   DriveListBoxControl = m_DriveListBoxControl
End Property

Public Property Let DriveListBoxControl(ByVal cDriveListBoxControl As Boolean)
   m_DriveListBoxControl = cDriveListBoxControl
   PropertyChanged "DriveListBoxControl"
End Property

Public Property Get ComboBoxControl() As Boolean
   ComboBoxControl = m_ComboBoxControl
End Property

Public Property Let ComboBoxControl(ByVal cComboBoxControl As Boolean)
   m_ComboBoxControl = cComboBoxControl
   PropertyChanged "ComboBoxControl"
End Property

Public Property Get OptionControl() As Boolean
   OptionControl = m_OptionControl
End Property

Public Property Let OptionControl(ByVal cOptionControl As Boolean)
   m_OptionControl = cOptionControl
   PropertyChanged "OptionControl"
End Property

Public Property Get CheckControl() As Boolean
   CheckControl = m_CheckControl
End Property

Public Property Let CheckControl(ByVal cCheckControl As Boolean)
   m_CheckControl = cCheckControl
   PropertyChanged "CheckControl"
End Property

Public Property Get ButtonControl() As Boolean
   ButtonControl = m_ButtonControl
End Property

Public Property Let ButtonControl(ByVal cButtonControl As Boolean)
   m_ButtonControl = cButtonControl
   PropertyChanged "ButtonControl"
End Property

Public Property Get FrameControl() As Boolean
   FrameControl = m_FrameControl
End Property

Public Property Let FrameControl(ByVal cFrameControl As Boolean)
   m_FrameControl = cFrameControl
   PropertyChanged "FrameControl"
End Property

Public Property Get PictureControl() As Boolean
   PictureControl = m_PictureControl
End Property

Public Property Let PictureControl(ByVal cPictureControl As Boolean)
   m_PictureControl = cPictureControl
   PropertyChanged "PictureControl"
End Property

Public Property Get TextControl() As Boolean
   TextControl = m_TextControl
End Property

Public Property Let TextControl(ByVal cTextControl As Boolean)
   m_TextControl = cTextControl
   PropertyChanged "TextControl"
End Property

Public Property Get Common_Dialog() As Boolean
   Common_Dialog = m_CommonDialog
End Property

Public Property Let Common_Dialog(ByVal cCommon_Dialog As Boolean)
   m_CommonDialog = cCommon_Dialog
   PropertyChanged "Common_Dialog"
End Property

Public Property Get MsgBox_InputBox() As Boolean
   MsgBox_InputBox = m_MsgBox_InputBox
End Property

Public Property Let MsgBox_InputBox(ByVal cMsgBox_InputBox As Boolean)
   m_MsgBox_InputBox = cMsgBox_InputBox
   PropertyChanged "MsgBox_InputBox"
End Property

Public Property Get EngineStarted() As Boolean
   EngineStarted = m_EngineStarted
End Property

Public Property Let EngineStarted(ByVal cEngineStarted As Boolean)
   m_EngineStarted = cEngineStarted
   PropertyChanged "EngineStarted"
End Property

Public Property Get IDE() As Boolean
   IDE = m_IDE
End Property

Public Property Let IDE(ByVal cIDE As Boolean)
   m_IDE = cIDE
   PropertyChanged "IDE"
 End Property

Private Sub UserControl_Paint()
If IDE Then
    If Not m_IDEStarted Then InitIDESubClassing
End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 ColorScheme = PropBag.ReadProperty("ColorScheme", WindowsXP_Blue)
 IDE = PropBag.ReadProperty("IDE", False)
 MsgBox_InputBox = PropBag.ReadProperty("MsgBox_InputBox", True)
 Common_Dialog = PropBag.ReadProperty("Common_Dialog", True)
 TextControl = PropBag.ReadProperty("TextControl", True)
 PictureControl = PropBag.ReadProperty("PictureControl", True)
 FrameControl = PropBag.ReadProperty("FrameControl", True)
 ButtonControl = PropBag.ReadProperty("ButtonControl", True)
 CheckControl = PropBag.ReadProperty("CheckControl", True)
 OptionControl = PropBag.ReadProperty("OptionControl", True)
 ComboBoxControl = PropBag.ReadProperty("ComboBoxControl", True)
 DriveListBoxControl = PropBag.ReadProperty("DriveListBoxControl", True)
 TabStripControl = PropBag.ReadProperty("TabStripControl", True)
 StatusBarControl = PropBag.ReadProperty("StatusBarControl", True)
 ProgressBarControl = PropBag.ReadProperty("ProgressBarControl", True)
 SliderControl = PropBag.ReadProperty("SliderControl", True)
 ImageComboControl = PropBag.ReadProperty("ImageComboControl", True)
 ListBoxControl = PropBag.ReadProperty("ListBoxControl", True)
 DirListBoxControl = PropBag.ReadProperty("DirListBoxControl", True)
 FileListBoxControl = PropBag.ReadProperty("FileListBoxControl", True)
 ListViewControl = PropBag.ReadProperty("ListViewControl", True)
 EngineStarted = PropBag.ReadProperty("EngineStarted", False)
 
End Sub


Private Sub UserControl_Resize()
UserControl.Width = 3735
UserControl.Height = 615
End Sub

Private Sub UserControl_Show()
'
End Sub

Private Sub UserControl_Terminate()
'
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 
 
 Call PropBag.WriteProperty("ColorScheme", m_ColorScheme, False)
 Call PropBag.WriteProperty("IDE", m_IDE, False)
 Call PropBag.WriteProperty("EngineStarted", m_EngineStarted, False)
 Call PropBag.WriteProperty("MsgBox_InputBox", m_MsgBox_InputBox, True)
 Call PropBag.WriteProperty("Common_Dialog", m_CommonDialog, True)
 Call PropBag.WriteProperty("TextControl", m_TextControl, True)
 Call PropBag.WriteProperty("ListBoxControl", m_ListBoxControl, True)
 Call PropBag.WriteProperty("PictureControl", m_PictureControl, True)
 Call PropBag.WriteProperty("FrameControl", m_FrameControl, True)
 Call PropBag.WriteProperty("ButtonControl", m_ButtonControl, True)
 Call PropBag.WriteProperty("CheckControl", m_CheckControl, True)
 Call PropBag.WriteProperty("OptionControl", m_OptionControl, True)
 Call PropBag.WriteProperty("ComboBoxControl", m_ComboBoxControl, True)
 Call PropBag.WriteProperty("DriveListBoxControl", m_DriveListBoxControl, True)
 Call PropBag.WriteProperty("TabStripControl", m_TabStripControl, True)
 Call PropBag.WriteProperty("StatusBarControl", m_StatusBarControl, True)
 Call PropBag.WriteProperty("ProgressBarControl", m_ProgressBarControl, True)
 Call PropBag.WriteProperty("SliderControl", m_SliderControl, True)
 Call PropBag.WriteProperty("ImageComboControl", m_ImageComboControl, True)
 Call PropBag.WriteProperty("ListViewControl", m_ListViewControl, True)
 Call PropBag.WriteProperty("FileListBoxControl", m_FileListBoxControl, True)
 Call PropBag.WriteProperty("DirListBoxControl", m_DirListBoxControl, True)

 If m_IDE = True Then
    UserControl.Parent.Hide
    UserControl.Parent.Visible = True
    UserControl.Parent.Show
    UserControl.Parent.Refresh
 End If
 


End Sub

