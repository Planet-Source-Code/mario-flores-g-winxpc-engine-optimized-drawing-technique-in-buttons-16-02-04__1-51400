VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWinXPCEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'==========================================================================================================
'==========================================================================================================
'
       '                   <<<<<<<<<<<<<<<     WINXPC ENGINE    >>>>>>>>>>>>>>>
'                                              VERSION 1.0.0
'
'                   AUTHOR:       .... MARIO ALBERTO FLORES GONZALEZ....
'
'                                       CD.JUAREZ CHIHUAHUA MEXICO
'
'                                      sistec_de_juarez@hotmail.com
'
'
'
'          ALL CODE WAS WRITTEN FROM SCRATCH ..WITH EXCEPTION OF STEVE MCMAHON SUBCLASS LIBRARY
'          THAT WAS UNMODIFIED..(steve@vbaccelerator.com), I ALSO USED SOME CODE IDEAS FROM THE
'          CHAMELEON BUTTON (gonchuki@ yahoo.es) TO MAKE THE CBUTTONXP CLASS !!! NOT COPIED !!!
'          RECODED-FIXED SOME MEMORY LEAKS ANG CHANGED LOGIC TO DRAW BUTTON SOME OF THE ROUTINES
'          FOR XP COLORS WHERE USED..AND SLIGHTLY MODIFIED.
'          USED APIVIEWER 2003 TO GET THE API DECLARATIONS AND CONST VALUES
'          ALSO A LOT OF VISITS TO THE MSDN LIBRARY...
'
'                 ANY OTHER SIMILAR CODE ON OTHER PEOPLE CODE OR LIBRARY ARE COINCIDENCE
'
'                      <<<<  YOU CAN'T USE THIS CODE FOR COMMERCIAL PURPOUSES   >>>>
'
'
'          NOTE:
'                      This control is NOT intended to be compiled into your application,
'                      it is intended to be compiled as a separate OCX and packaged with your app.
'
'                   This notice may not be removed or altered from any source distribution.
'
'==========================================================================================================
'==========================================================================================================
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Option Explicit

Implements ISubclass

Private m_bSubclass        As Boolean
Private m_hWnd             As Long


Private m_ActiveScaleMode As ScaleModeConstants
Private m_SchemeColor As CWindowColors
Private m_IdeSubClass As Boolean
Private m_CBug As Boolean
 
'//-----------------------------------------------

Private m_Adodc       As Boolean
Private m_Button      As Boolean
Private m_CheckBox    As Boolean
Private m_Combo       As Boolean '+ DriveList
Private m_FileListBox As Boolean '+ DirListBox
Private m_Frame       As Boolean
Private m_ICombo      As Boolean
Private m_ListBox     As Boolean
Private m_ListView    As Boolean           '//---Supported Subclassed Controls
Private m_MsgBox      As Boolean
Private m_CommonD     As Boolean
Private m_Option      As Boolean
Private m_Picture     As Boolean
Private m_ProgressBar As Boolean
Private m_Slider      As Boolean
Private m_StatusBar   As Boolean
Private m_TabStrip    As Boolean
Private m_Text        As Boolean
Private m_SysListView32 As Boolean
Private m_Static      As Boolean
'//-----------------------------------------------


Private MyClassObject As Object         '//--- Object Sender  (Contains Subclassed Control Information) VB 6



Private ButtonState As ControlState
Private MovementFlag As ControlState



'//-----------------------------------------------
    Private isFocused As Boolean                   '//---  Controls General Properties
    Private bOver As Boolean
    Private bDown As Boolean
'//-----------------------------------------------

'//-----------------------------------------------
    Private m_CurrentButton As Integer          '//---  Adoedc Control Variables
'//-----------------------------------------------

'//-----------------------------------------------
    Private m_CurrentItem As Long
    Private m_FlagItem As Long                     '//---  ListView Control Variables
    Private m_LVIRect As RECT ' ListViewItemRect
'//-----------------------------------------------

'//-----------------------------------------------
    Private m_TSINumber  As Integer  'TabStripItemNumber
    Private m_TSILNumber As Integer  'TabStripItemLastNumber
    Private m_TSCTab     As Boolean
'//-----------------------------------------------
    
'------------------------------------
Private DestDC      As Long
Private MaskDC      As Long
Private MemDC       As Long
Private OrigDC      As Long
Private MaskPic     As Long              'Temporary DC
Private MemPic      As Long
Private TempPic     As Long
Private OrigPic     As Long
Private TempDC      As Long
'-------------------------------------

Private origBrush As Long
Private TempBrush As Long
Private origColor As Long 'BackColor
Dim m_ItemRect As RECT
Dim m_Width    As Long
Dim m_Height   As Long


Private m_PreDraw As Boolean
Private iHw As Integer, iLW As Integer


'==========================================================================================
'==========================================================================================
'                        THIS WILL SUBCLASS MESSAGESBOXES & INPUTBOXES
                      

Public Sub BeforeAttachMessageBox(ByVal lhWnd As Long)
 If lhWnd = 0 Then Exit Sub
 
 pRelease
    m_MsgBox = True      ' PRE-DEFINE "m_MsgBox" VALUE BEFORE ENGINE IS STARTED
 pAttach lhWnd
 
End Sub

'==========================================================================================
'==========================================================================================
'                        THIS WILL SUBCLASS COMMONDIALOG
        
Public Sub BeforeAttachCommonDialog(ByVal lhWnd As Long)
 If lhWnd = 0 Then Exit Sub
 
 pRelease
    m_CommonD = True
    pAttach lhWnd
End Sub

'==========================================================================================
'==========================================================================================
'                        THIS WILL SUBCLASS MESSAGESBOXES & INPUTBOXES

Public Sub AfterAttachMessageBox(ByVal lhWnd As Long)

 If lhWnd = 0 Then Exit Sub
 
 pRelease

     Select Case ThisWindowClassName(lhWnd)
      
       Case "Edit"
         
         If InStr(ThisWindowClassName(GetParent(lhWnd)), "ComboBox") = 0 Then
            MakeWindowFlat lhWnd
            m_Text = True
         End If            '//---MessagesBoxes and InputBoxes only have this kind of controls
       
       Case "Button"
        If Left$(GetObjectText(lhWnd), InStr(GetObjectText(lhWnd), vbNullChar) - 1) <> "Open as &read-only" Then
          m_Button = True
        Else
          m_CheckBox = True
        End If
        
       Case "ComboBox"
          m_Combo = True
      
       Case Else:
            Exit Sub '(Nothing) UnSupported Control
          
     End Select
 
 pAttach lhWnd


       
End Sub



'==========================================================================================
'==========================================================================================
'                         THIS WILL SUBCLASS ALL THE SUPPORTED CONTROLS

Public Sub Attach(ByVal ObjThis As Object)
Dim lhWnd As Long
    
    
    pRelease '//--- First UnSubclass All Controls.. ;)
   
    On Error Resume Next
   
    If TypeName(ObjThis) = "Adodc" Then
       lhWnd = ChildWindowFromPoint(ObjThis.Parent.hwnd, ObjThis.Left, ObjThis.Top)
    Else
       lhWnd = ObjThis.hwnd '//---Short Reference
    End If
       
    If Err.Number <> 0 Then Exit Sub  'Something happen :(
  

    Set MyClassObject = ObjThis '//--- Control's Properties Container
   

    
    
    '//======================================================================================
    '                             Select the Type Of Control to Subclass
    '//======================================================================================
    

    
''
    
    Select Case ThisWindowClassName(lhWnd)
   
        Case "AdodcWndClass"
             m_Adodc = True
             m_CurrentButton = -1
             MakeWindowFlat lhWnd
        
        Case "ThunderCheckBox", "ThunderRT6CheckBox"
           ' If MyClassObject.Style = 1 Then Exit Sub              '//--- Only Standard Style CheckBox (No Graphical)
            m_CheckBox = True

        Case "ThunderComboBox", "ThunderRT6ComboBox", "ThunderDriveListBox", "ThunderRT6DriveListBox"
            m_Combo = True

        Case "ThunderCommandButton", "ThunderRT6CommandButton", "Button"
           m_Button = True

        Case "ThunderFrame", "ThunderRT6Frame"
            If MyClassObject.BorderStyle = 0 Then Exit Sub        '//--- If Frame <> Fixed Single Then Don't Sublcass ;)
             m_Frame = True

        Case "ComboBox", "ImageCombo20WndClass"
            lhWnd = FindWindowEx(lhWnd, 0&, "ComboBox", ByVal 0&) '//--- Find Combo Inside control
            m_ICombo = True
  
        Case "ListView20WndClass"
            If m_IdeSubClass = True Then Exit Sub
            MakeWindowFlat lhWnd
            lhWnd = FindWindowEx(lhWnd, 0, "msvb_lib_header", 0&) '//--- Find Header Inside control
            m_ListView = True

        Case "ThunderOptionButton", "ThunderRT6OptionButton"
           'If MyClassObject.Style = 1 Then Exit Sub              '//--- Only Standard Style OptionButton (No Graphical)
            m_Option = True

        Case "ProgressBar20WndClass"
             MakeWindowFlat lhWnd
             SendMessageLong lhWnd, CCM_SETBKCOLOR, 0&, ByVal &HFFFFFF  '//--Give The XP Look BackGroundColor = White
             SendMessageLong lhWnd, PBM_SETBARCOLOR, 0&, ByVal &HFFFFFF             '//--Give The XP Look BackGroundColor = White
             m_ProgressBar = True

        Case "Slider20WndClass"
            m_Slider = True

        Case "StatusBar20WndClass"
            m_StatusBar = True

        Case "ThunderListBox", "ThunderRT6ListBox", "ThunderDirListBox", "ThunderRT6DirListBox"
            MakeWindowFlat lhWnd
            m_ListBox = True
        
        Case "ThunderFileListBox", "ThunderRT6FileListBox"
             MakeWindowFlat lhWnd
             m_FileListBox = True
        
        Case "TabStrip20WndClass", "TabStripWndClass"
            Dim TSOWNERDRAW As Long
            
            TSOWNERDRAW = GetWindowLong(lhWnd, GWL_STYLE)
            Call SetWindowLong(lhWnd, GWL_STYLE, TSOWNERDRAW Or TCS_OWNERDRAWFIXED)
            m_TabStrip = True
        
         Case "ThunderTextBox", "ThunderRT6TextBox", "Edit"
             '//--- Set TextBox Appareance= FLAT Style
                    
                                     
                    MakeWindowFlat lhWnd
                    m_Text = True
        Case "ThunderPictureBox", "ThunderPictureBoxDC", "ThunderRT6PictureBoxDC"
             '//--- Set PictureBox Appareance= FLAT Style
                    MakeWindowFlat lhWnd
                    m_Picture = True
        Case Else:
        
            Exit Sub '(Nothing) UnSupported Control <---Actually This Line is never Fired.
            
   
   
    End Select


   pAttach lhWnd '//--- Finally Subclass
   
End Sub

Private Sub pAttach(ByRef hWndA As Long)
  
  m_hWnd = hWndA '//-- Reference Used in All Class Subs & Functions Point's Control hWnd
  
  SubclassMessage m_hWnd, True '//--- Subclass
      
  m_bSubclass = True '//--- Success Control is Subclassed !!  ;)
   
End Sub

Private Sub pRelease()
   
   If m_bSubclass = False Then Exit Sub '//--- No Subclass before...Don't Unsubclass (Exit)
       
   SubclassMessage m_hWnd, False  '//--- Unsubclass
   RedrawWindow m_hWnd, ByVal 0&, ByVal 0&, &H1 '//---(invoke a Paint-event) turn Normal ;)
   KillTimer m_hWnd, 1         '//--  Should never happen that we have a timer left over (but just in case)
   Set MyClassObject = Nothing '//--  Cleaning ClassObject just in case
    
End Sub





Private Sub SubclassMessage(ByVal Sender As Long, ByVal Action As Boolean)

    
     
     If Action = True Then
        
         If Not m_MsgBox Or Not m_CommonD Or Not m_TabStrip Or Not m_ListView Then AttachMessage Me, Sender, WM_PAINT
         If m_MsgBox Or m_CommonD Then AttachMessage Me, Sender, WM_ACTIVATE
        
         If m_Button Then AttachMessage Me, Sender, WM_SETFOCUS
         If m_Button Then AttachMessage Me, Sender, WM_KILLFOCUS
         If m_Button Or m_CheckBox Or m_Option Then AttachMessage Me, Sender, WM_KEYDOWN
         If m_Button Or m_CheckBox Or m_Option Then AttachMessage Me, Sender, WM_KEYUP
         
         If m_TabStrip Then AttachMessage Me, GetParent(Sender), WM_DRAWITEM
         If m_TabStrip Then AttachMessage Me, Sender, WM_PRINTCLIENT
        
         If m_Adodc Or m_Button Or m_CheckBox Or m_Combo Or m_ICombo Or m_ListView Or m_Option Or m_Slider Or m_TabStrip _
         Then AttachMessage Me, Sender, WM_MOUSEMOVE
         
         If m_Adodc Or m_Button Or m_CheckBox Or m_Combo Or m_ICombo Or m_ListView Or m_Option Or m_Slider Or m_TabStrip _
         Then AttachMessage Me, Sender, WM_TIMER
         
         If m_Adodc Or m_Button Or m_CheckBox Or m_Combo Or m_ICombo Or m_ListView Or m_Option Or m_Slider Or m_TabStrip _
         Then AttachMessage Me, Sender, WM_LBUTTONDOWN
         
         If m_Adodc Or m_Button Or m_CheckBox Or m_Combo Or m_ICombo Or m_ListView Or m_Option Or m_Slider Or m_TabStrip _
         Then AttachMessage Me, Sender, WM_LBUTTONUP
         
         If m_Picture Or m_Button Or m_CheckBox Or m_ICombo Or m_Option _
         Then AttachMessage Me, Sender, WM_ENABLE
         
         If m_Combo Then AttachMessage Me, GetParent(Sender), WM_COMMAND
         If m_ICombo Then AttachMessage Me, GetWindow(Sender, GW_CHILD), WM_COMMAND
  
      Else
         
         If Not m_MsgBox Or Not Not m_CommonD Or Not m_TabStrip Or Not m_ListView Then DetachMessage Me, Sender, WM_PAINT
         If m_MsgBox Or m_CommonD Then DetachMessage Me, Sender, WM_ACTIVATE
                  
         If m_Button Then DetachMessage Me, Sender, WM_SETFOCUS
         If m_Button Then DetachMessage Me, Sender, WM_KILLFOCUS
         If m_Button Or m_CheckBox Or m_Option Then DetachMessage Me, Sender, WM_KEYDOWN
         If m_Button Or m_CheckBox Or m_Option Then DetachMessage Me, Sender, WM_KEYUP
         
         If m_TabStrip Then DetachMessage Me, GetParent(Sender), WM_DRAWITEM
         If m_TabStrip Then DetachMessage Me, Sender, WM_PRINTCLIENT
                  
         If m_Adodc Or m_Button Or m_CheckBox Or m_Combo Or m_ICombo Or m_ListView Or m_Option Or m_Slider Or m_TabStrip _
         Then DetachMessage Me, Sender, WM_MOUSEMOVE
         
         If m_Adodc Or m_Button Or m_CheckBox Or m_Combo Or m_ICombo Or m_ListView Or m_Option Or m_Slider Or m_TabStrip _
         Then DetachMessage Me, Sender, WM_TIMER
         
         If m_Adodc Or m_Button Or m_CheckBox Or m_Combo Or m_ICombo Or m_ListView Or m_Option Or m_Slider Or m_TabStrip _
         Then DetachMessage Me, Sender, WM_LBUTTONDOWN
         
         If m_Adodc Or m_Button Or m_CheckBox Or m_Combo Or m_ICombo Or m_ListView Or m_Option Or m_Slider Or m_TabStrip _
         Then DetachMessage Me, Sender, WM_LBUTTONUP
         
         If m_Picture Or m_Button Or m_CheckBox Or m_ICombo Or m_Option _
         Then DetachMessage Me, Sender, WM_ENABLE
         
         If m_Combo Then DetachMessage Me, GetParent(Sender), WM_COMMAND
         If m_ICombo Then DetachMessage Me, GetWindow(Sender, GW_CHILD), WM_COMMAND
         
      
      End If
    
    
      
End Sub

Private Sub DrawControlClass()
   
   If m_hWnd = 0 Then Exit Sub
   
   If m_Combo Or m_ICombo Then
      DrawCombo
   ElseIf m_Adodc Then
      DrawAdodc
   ElseIf m_Button Then
      DrawButton
   ElseIf m_Text Or m_Picture Or m_ListBox Or m_FileListBox Or m_SysListView32 Then
      DrawText
   ElseIf m_ProgressBar Then
      DrawProgressBar
   ElseIf m_TabStrip Then
      DrawTabStrip
   ElseIf m_CheckBox Then
      DrawCheckBox
   ElseIf m_Option Then
      DrawOptionButton
   ElseIf m_StatusBar Then
      DrawStatusBar
   ElseIf m_Slider Then
      DrawSlider
   ElseIf m_ListView Then
      DrawText
      DrawListView
   ElseIf m_Frame Then
      DrawFrame
   Else
      
   End If
   
              
End Sub



Private Sub DrawButton()
Dim NewButton As cButtonXP
Dim hdc As Long
Dim lState As Long

lState = SendMessageLong(m_hWnd, BM_GETSTATE, 0, 0)

   
hdc = GetDC(m_hWnd)

   Set NewButton = New cButtonXP
     NewButton.hdc = IIf(IsWindowEnabled(m_hWnd) = False, hdc, IIf(m_PreDraw = True, TempDC, hdc))
     NewButton.hwnd = m_hWnd
     NewButton.Enabled = IsWindowEnabled(m_hWnd)
     NewButton.Focused = isFocused
     NewButton.Over = bOver
     NewButton.Down = ((lState And BST_CHECKED) = BST_CHECKED) Or ((lState And BST_PUSHED) = BST_PUSHED)
     NewButton.ColorScheme = m_SchemeColor
     NewButton.PreDraw = m_PreDraw
     NewButton.DrawButtonXP
   Set NewButton = Nothing
   
   ReleaseDC m_hWnd, hdc
   
End Sub

Private Sub DrawText()
Dim RcItem As RECT
Dim T_hWnd As Long
Dim hdc As Long
Dim Point As POINTAPI
Dim NewTextXP As cTextBoxXP

Set NewTextXP = New cTextBoxXP

 
 If m_ListView Or m_SysListView32 Then
    T_hWnd = m_hWnd
    m_hWnd = GetParent(m_hWnd)
 End If
  
  
  hdc = GetParent(m_hWnd)
  
  GetWindowRect m_hWnd, RcItem
    
  Point.X = RcItem.Left
  Point.Y = RcItem.Top

  ScreenToClient hdc, Point
  RcItem.Left = Point.X
  RcItem.Top = Point.Y

  Point.X = RcItem.Right
  Point.Y = RcItem.Bottom

  ScreenToClient hdc, Point
  RcItem.Right = Point.X
  RcItem.Bottom = Point.Y

  InflateRect RcItem, 1, 1
  hdc = GetDC(hdc)


   NewTextXP.hdc = hdc
   NewTextXP.hwnd = m_hWnd
   NewTextXP.IBottom = RcItem.Bottom
   NewTextXP.ILeft = RcItem.Left
   NewTextXP.IRight = RcItem.Right
   NewTextXP.ITop = RcItem.Top
   NewTextXP.State = IIf(IsWindowEnabled(m_hWnd) = 0, C_Disabled, ButtonState)
   NewTextXP.ColorScheme = m_SchemeColor
   NewTextXP.DrawTextBoxXP

   Set NewTextXP = Nothing
   ReleaseDC m_hWnd, hdc
 
 
   If m_ListView Or m_SysListView32 Then m_hWnd = T_hWnd
      
 
End Sub

Private Sub DrawListView()
 Dim NewListViewXP As cListViewXP
 Dim hdc As Long
 Dim RcItem As RECT
 Dim i As Long
 
 
 GetClientRect m_hWnd, RcItem
 hdc = GetDC(m_hWnd)
 
  Set NewListViewXP = New cListViewXP
    
    NewListViewXP.hdc = hdc
    NewListViewXP.hwnd = m_hWnd
    NewListViewXP.ItemFont = MyClassObject.Font.Name
    NewListViewXP.ItemItalic = MyClassObject.Font.Italic
    NewListViewXP.ItemUnderline = MyClassObject.Font.Underline
    NewListViewXP.ItemSize = MyClassObject.Font.Size
    NewListViewXP.Width = RcItem.Right
    NewListViewXP.IBottom = m_LVIRect.Bottom
    NewListViewXP.ILeft = m_LVIRect.Left
    NewListViewXP.IRight = m_LVIRect.Right
    NewListViewXP.ITop = m_LVIRect.Top
    NewListViewXP.State = ButtonState
    If m_CurrentItem = -1 Then SendMessageLong m_hWnd, &H1200 + 7, 1, RcItem
    NewListViewXP.CleanHeaders RcItem.Bottom
    NewListViewXP.DrawDownBorder RcItem.Bottom
    If m_CurrentItem <> -1 Then NewListViewXP.DrawListView
    GoSub DrawCleanText
  
  Set NewListViewXP = Nothing
  ReleaseDC m_hWnd, hdc
 
Exit Sub

DrawCleanText:
               
      For i = 0 To MyClassObject.ColumnHeaders.count - 1
        
        SendMessageLong m_hWnd, &H1200 + 7, i, RcItem
        NewListViewXP.IBottom = RcItem.Bottom
        NewListViewXP.ILeft = RcItem.Left
        NewListViewXP.IRight = RcItem.Right
        NewListViewXP.ITop = RcItem.Top
        NewListViewXP.ItemCaption = MyClassObject.ColumnHeaders(i + 1).Text
        NewListViewXP.SetTextParams
    
    Next i

Return

End Sub

Private Sub DrawFrame()
 Dim NewFrame As cFrameXP
 Dim hdc As Long
 
 hdc = GetDC(m_hWnd)

 Set NewFrame = New cFrameXP
   NewFrame.hwnd = m_hWnd
   NewFrame.hdc = hdc
   NewFrame.FontItalic = MyClassObject.FontItalic
   NewFrame.FontSize = MyClassObject.FontSize
   NewFrame.FontUnderline = MyClassObject.FontUnderline
   NewFrame.FrameCaption = MyClassObject.Caption
   NewFrame.FrameFont = MyClassObject.FontName
   NewFrame.Backcolor = MyClassObject.Backcolor
   NewFrame.Forecolor = MyClassObject.Forecolor
   NewFrame.DrawFrame
 Set NewFrame = Nothing
   
 ReleaseDC m_hWnd, hdc

End Sub

Private Sub DrawProgressBar()
 Dim NewProgressBarXP As cProgressBarXP
 Dim hdc As Long
   
 hdc = GetDC(m_hWnd)
  

 Set NewProgressBarXP = New cProgressBarXP
     With NewProgressBarXP
        .hdc = hdc
        .hwnd = m_hWnd
        .Max = MyClassObject.Max
        .Min = MyClassObject.Min
        .Value = MyClassObject.Value
        .Orientation = MyClassObject.Orientation
        .Scrolling = MyClassObject.Scrolling
        .ColorScheme = m_SchemeColor
        .DrawProgressBar
     End With
 Set NewProgressBarXP = Nothing
 
 
 
 
   ReleaseDC m_hWnd, hdc

End Sub

Private Sub DrawAdodc()
Dim NewAdodc As cAdodc
Dim hdc As Long

 hdc = GetDC(m_hWnd)

 Set NewAdodc = New cAdodc
     NewAdodc.hdc = hdc
     NewAdodc.hwnd = m_hWnd
     NewAdodc.Enabled = MyClassObject.Enabled
     NewAdodc.cButton = m_CurrentButton
     NewAdodc.ColorScheme = m_SchemeColor
    
     NewAdodc.Down = IIf(ButtonState = C_Down, True, False)
     NewAdodc.DrawAdodc
 
 
 Set NewAdodc = Nothing

 ReleaseDC m_hWnd, hdc


End Sub

Private Sub DrawCheckBox()
Dim NewCheckBoxXP As cCheckBoxXP
Dim hdc As Long
Dim lState As Long

lState = SendMessageLong(m_hWnd, BM_GETSTATE, 0, 0)
      
hdc = GetDC(m_hWnd)
 
  Set NewCheckBoxXP = New cCheckBoxXP
    NewCheckBoxXP.hdc = hdc
    NewCheckBoxXP.hwnd = m_hWnd
    NewCheckBoxXP.Enabled = IsWindowEnabled(m_hWnd)
    NewCheckBoxXP.Over = bOver
    NewCheckBoxXP.Down = ((lState And BST_PUSHED) = BST_PUSHED)
    NewCheckBoxXP.Value = SendMessageLong(m_hWnd, BM_GETCHECK, 0&, 0&)
    NewCheckBoxXP.Alignment = WindowAlignment(m_hWnd)
    NewCheckBoxXP.DrawCheckBox
 Set NewCheckBoxXP = Nothing
  
 ReleaseDC m_hWnd, hdc
 
End Sub

Private Sub DrawOptionButton()
 Dim NewOptionButtonXP As cOptionButtonXP
 Dim hdc As Long
 Dim lState As Long

 lState = SendMessageLong(m_hWnd, BM_GETSTATE, 0, 0)
 
 hdc = GetDC(m_hWnd)
 
 Set NewOptionButtonXP = New cOptionButtonXP
    NewOptionButtonXP.hdc = hdc
    NewOptionButtonXP.hwnd = m_hWnd
    NewOptionButtonXP.Enabled = IsWindowEnabled(m_hWnd)
    NewOptionButtonXP.Over = bOver
    NewOptionButtonXP.Down = ((lState And BST_PUSHED) = BST_PUSHED)
    NewOptionButtonXP.Value = SendMessageLong(m_hWnd, BM_GETCHECK, 0&, 0&)
    NewOptionButtonXP.Alignment = WindowAlignment(m_hWnd)
    NewOptionButtonXP.DrawOptionButton
 Set NewOptionButtonXP = Nothing
 
 ReleaseDC m_hWnd, hdc

End Sub

Private Sub DrawSlider()
 Dim NewSliderXP As cSliderXP
 Dim hdc As Long
 
 hdc = GetDC(m_hWnd)

 Set NewSliderXP = New cSliderXP
    With NewSliderXP
        .hdc = hdc
        .hwnd = m_hWnd
        .Max = MyClassObject.Max
        .Min = MyClassObject.Min
        .Value = MyClassObject.Value
        .Orientation = MyClassObject.Orientation
        .TickStyle = MyClassObject.TickStyle
        .State = IIf(IsWindowEnabled(m_hWnd) = 0, C_Disabled, IIf(isFocused And ButtonState <> C_Down, C_Focus, ButtonState))
        .DrawSlider
     End With
 Set NewSliderXP = Nothing
 
 ReleaseDC m_hWnd, hdc

End Sub
Private Sub DrawStatusBar()
 Dim NewStatusBarXP As cStatusBarXP
 Dim hdc As Long
 Dim count As Single
 
 hdc = GetDC(m_hWnd)
 
 Set NewStatusBarXP = New cStatusBarXP
    NewStatusBarXP.hdc = hdc
    NewStatusBarXP.hwnd = m_hWnd
    NewStatusBarXP.DrawBar
    For count = 1 To MyClassObject.Panels.count
         If m_IdeSubClass = False Then MyClassObject.Panels(count).Bevel = 0  '//--- Read Only Property (Disabled if IDE Subclassed = TRUE)
         NewStatusBarXP.DrawPanel (MyClassObject.Panels(count).Left / 15) + (MyClassObject.Panels(count).Width / 15)
    Next count
 Set NewStatusBarXP = Nothing
 
 ReleaseDC m_hWnd, hdc
 
End Sub



Private Sub DrawTabStrip()
Dim NewTabXP As cTabStripXP
Dim hdc As Long
  
 
hdc = GetDC(m_hWnd)
 
 
 Set NewTabXP = New cTabStripXP
     NewTabXP.hdc = hdc
     NewTabXP.hwnd = m_hWnd
     NewTabXP.CurrentItem = m_TSINumber
     NewTabXP.LastItem = m_TSILNumber
     NewTabXP.ItemFont = MyClassObject.Font.Name
     NewTabXP.ItemItalic = MyClassObject.Font.Italic
     NewTabXP.ItemUnderline = MyClassObject.Font.Underline
     NewTabXP.ItemSize = MyClassObject.Font.Size
     NewTabXP.ColorScheme = m_SchemeColor
     NewTabXP.CleanItem = m_TSCTab
     NewTabXP.DrawTabStrip
     
   
   Set NewTabXP = Nothing
   
   ReleaseDC m_hWnd, hdc

End Sub

Private Sub DrawCombo()
Dim NewComboXP As cComboBoxXP
Dim hdc As Long

hdc = GetDC(m_hWnd)
 
    Set NewComboXP = New cComboBoxXP
        NewComboXP.hdc = hdc
        NewComboXP.hwnd = m_hWnd
        NewComboXP.State = IIf(IsWindowEnabled(m_hWnd) = 0, C_Disabled, ButtonState)
        NewComboXP.ColorScheme = m_SchemeColor
        NewComboXP.DrawComboBox
    Set NewComboXP = Nothing

    ReleaseDC m_hWnd, hdc

End Sub


Private Function DroppedDown() As Boolean
  DroppedDown = (SendMessageLong(m_hWnd, CB_GETDROPPEDSTATE, 0, 0) <> 0)
End Function

Private Sub Class_Terminate()
   pRelease
   If hHook Then UnhookWindowsHookEx hHook
End Sub

Public Sub UnloadEngine()
        Class_Terminate
End Sub
Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
   ' intentionally blank
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
   
   If (CurrentMessage = WM_PAINT) Then
      ISubclass_MsgResponse = emrPreprocess
   Else
      ISubclass_MsgResponse = emrPostProcess
   End If
   
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
          
       
     'm_CBug = IIf(ButtonState = C_Down, True, False)

  
          
'======================================================================================
     'Line Below Paints the ProgressBar a lot of times ...when ever
     'the progress fires a message in WindowProc..
     'This eliminates the Flicker on the control  + - 50% ;)
     If m_ProgressBar Then DrawControlClass
         
          
          
          
'=============================================================================================
'                                       SUBCLASS MESSAGES
                                        
 Select Case iMsg   '//--- SELECT THE SUBCLASSED WINDOW MESSAGE.... CLEAR ENOUGH?
   

'==========================================================================================
'==========================================================================================
'The WM_COMMAND message is sent when the user selects a command item from a menu, when a control
'sends a notification message to its parent window, or when an accelerator keystroke is translated.

   
   Case WM_COMMAND
      
     'The CBN_CLOSEUP notification message is sent
     'when the list box of a combo box has been closed.
     'The parent window of the combo box receives this
     'notification message through the WM_COMMAND message
      
      
    
      If (m_hWnd = lParam) Then ' lParam handle to Control (combo box)
        
         Select Case wParam \ &H10000  'WParam identifier of control (combo box)
         
            Case CBN_CLOSEUP
                 ButtonState = C_Normal
                 MovementFlag = C_Disabled
                 bOver = False
                 DrawControlClass
                 KillTimer m_hWnd, 1
            Case Else
                 'Nothing
             
         End Select
            
           
      End If
    
 '==========================================================================================
'==========================================================================================
'The WM_PRINTCLIENT message is sent to a window to request that it draw its client area in
'the specified device context.
 
 Case WM_PRINTCLIENT
      
      PrintTabClientColor

'==========================================================================================
'==========================================================================================
'The WM_DRAWITEM message is sent to the owner window of an owner-drawn button, combo box,
'list box, or menu when a visual aspect of the button, combo box, list box, or menu has changed.<-Tabstrip in this case
   
   
   Case WM_DRAWITEM
      
      Dim tDis As DRAWITEMSTRUCT
      
      If wParam = 0 Then
         CopyMemory tDis, ByVal lParam, Len(tDis)
         If tDis.CtlType = ODT_TAB And tDis.hWndItem = m_hWnd Then   '//--Check For StripTab
            m_TSILNumber = m_TSINumber
            m_TSINumber = tDis.itemID
            DrawControlClass
         End If
      End If
 
  
'==========================================================================================
'==========================================================================================
'The WM_ACTIVATE message is sent to both the window being activated and the window being
'deactivated.If the windows use the same input queue, the message is sent synchronously,
'first to the window procedure of the top-level window being deactivated, then to the window
'procedure of the top-level window being activated. If the windows use different input queues,
'the message is sent asynchronously, so the window is activated immediately.

                     
   Case WM_ACTIVATE
                
                             
              If m_MsgBox And InStr(ThisWindowClassName(lParam), "#32770") > 0 Then
                  Icount = 0 'm_ActiveClassesNumber
                  EnumChildWindows lParam, AddressOf EnumChildProc, ByVal 0&
              End If
                              '//--Support For MessageBox
              
              If m_CommonD And InStr(ThisWindowClassName(lParam), "NewHelpClass") > 0 Then
                If hHook Then UnhookWindowsHookEx hHook
                hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf AppHook, App.hInstance, App.ThreadID)
              End If
                              '//--Support For CommonDialog




'==========================================================================================
'==========================================================================================
'An application sends the WM_PAINT message when the system or another application makes a request
'to paint a portion of an application's window. The message is sent when the UpdateWindow or
'RedrawWindow function is called, or by theDispatchMessage function when the application obtains
'a WM_PAINT message by using theGetMessage orPeekMessage function
   
   
   Case WM_PAINT
        
    If m_Button Then
    
        If IsWindowEnabled(m_hWnd) Then
            PrePaintControl
            ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, OrigDC, lParam)
        End If
    
        PostPaintControl
  
    ElseIf Not m_ProgressBar Then
  
        DrawControlClass
    
    End If
  
'==========================================================================================
'==========================================================================================
'The WM_LBUTTONDOWN message is posted when the user presses the left mouse button while the
'cursor is in the client area of a window. If the mouse is not captured, the message is posted
'to the window beneath the cursor. Otherwise, the message is posted to the window that has captured the mouse.

    Case WM_LBUTTONDOWN
     m_CBug = True
     bDown = True
     ButtonState = C_Down
 
     RedrawWindow m_hWnd, ByVal 0&, ByVal 0&, &H1 '//---(invoke a Paint-event) ..See WM_PAINT For Details
      
          
     
'==========================================================================================
'==========================================================================================
'The WM_LBUTTONUP message is posted when the user releases the left mouse button while the
'cursor is in the client area of a window. If the mouse is not captured, the message is posted
'to the window beneath the cursor. Otherwise, the message is posted to the window that has captured the mouse.
   
   Case WM_LBUTTONUP
       
   
   m_CBug = True
   bDown = False
   ButtonState = C_Up
   
              
        
   '//---Slider XP Controls Hover Area (Hover Area only inside of THUMB)
   
   If m_Slider Then
           If CheckSliderArea = True Then ButtonState = C_Over
           If CheckSliderArea = False Then ButtonState = C_Normal
   
   ElseIf m_Adodc Then
      
           MovementFlag = C_Disabled
           SetTimer m_hWnd, 1, 1, 0
        
        
   '//---Ugly but must repaint every optionbutton control in actual container every time it changes value
   '     because can 't seem to get right values with ownerdraw buttons except for graphic styled one's
                
        'Solution:
   ElseIf m_Option Then
           If InsideArea(m_hWnd) = True Then
                EnumChildWindows GetParent(m_hWnd), AddressOf CleanOptionButtomArray, ByVal 0&
           Else
                DrawControlClass        '//--This Works for None-Graphic & None-OwnedDraw OptionButtons
           End If
   ElseIf m_TabStrip Then
   
   
   Else
          
          RedrawWindow m_hWnd, ByVal 0&, ByVal 0&, &H1 '//---(invoke a Paint-event) ..See WM_PAINT For Details
   
   End If
                         
        
   m_CurrentItem = -1
       
        
        
'==========================================================================================
'==========================================================================================
'The WM_ENABLE message is sent when an application changes the enabled state of a window.
'It is sent to the window whose enabled state is changing. This message is sent before the
'EnableWindow function returns, but after the enabled state (WS_DISABLED style bit) of the window has changed.
   
   
   Case WM_ENABLE

       If IsWindowEnabled(m_hWnd) = 0 Then
            ButtonState = C_Disabled
       Else
            ButtonState = C_Normal
       End If
    
       RedrawWindow m_hWnd, ByVal 0&, ByVal 0&, &H1 '//---(invoke a Paint-event) ..See WM_PAINT For Details
         
'==========================================================================================
'==========================================================================================
'The WM_SETFOCUS message is sent to a window after it has gained the keyboard focus.


    Case WM_SETFOCUS
                
'         ButtonState = C_Disabled
          
          isFocused = True
          
          If m_Button Then
             PostPaintControl

          Else
             RedrawWindow m_hWnd, ByVal 0&, ByVal 0&, &H1 '//---(invoke a Paint-event) ..See WM_PAINT For Details
    
          End If

'==========================================================================================
'==========================================================================================
'The WM_KILLFOCUS message is sent to a window immediately before it loses the keyboard focus.

    Case WM_KILLFOCUS
        
'        ButtonState = C_Disabled
        isFocused = False
        PostPaintControl


'==========================================================================================
'==========================================================================================
'The WM_KEYDOWN message is posted to the window with the keyboard focus when a nonsystem key is pressed.
'A nonsystem key is a key that is pressed when the alt key is not pressed.

       
    Case WM_KEYDOWN
        
            If Not bDown Then
                
                bDown = True
                
                LongInt2Int wParam, iHw, iLW
            
                Select Case (iLW)
                
                    Case vbKeySpace
         
                     RedrawWindow m_hWnd, ByVal 0&, ByVal 0&, &H1 '//---(invoke a Paint-event) ..See WM_PAINT For Details
             
                End Select
                 
                 
             End If

'==========================================================================================
'==========================================================================================
'The WM_KEYUP message is posted to the window with the keyboard focus when a nonsystem key is released.
'A nonsystem key is a key that is pressed when the alt key is not pressed, or a keyboard key
'that is pressed when a window has the keyboard focus.
     
     Case WM_KEYUP
            
             
             If bDown Then
                
                bDown = False
                
                LongInt2Int wParam, iHw, iLW
            
                Select Case (iLW)
                
                    Case vbKeySpace
         
                     RedrawWindow m_hWnd, ByVal 0&, ByVal 0&, &H1 '//---(invoke a Paint-event) ..See WM_PAINT For Details
             
                 End Select
                 
                 
             End If
             
     
'==========================================================================================
'==========================================================================================
'The WM_MOUSEMOVE message is posted to a window when the cursor moves. If the mouse is not
'captured,the message is posted to the window that contains the cursor. Otherwise, the message
'is posted to the window that has captured the mouse.

   
   Case WM_MOUSEMOVE
  
   If m_Button Or m_CheckBox Or m_Option Then
                        
      If InsideArea(m_hWnd) = True Then
           If bOver = False Then
                 bOver = True
                 RedrawWindow m_hWnd, ByVal 0&, ByVal 0&, &H1 '//---(invoke a Paint-event) ..See WM_PAINT For Details
                 SetTimer m_hWnd, 1, 1, 0
            End If
            
      Else
           If bOver Then
                  RedrawWindow m_hWnd, ByVal 0&, ByVal 0&, &H1 '//---(invoke a Paint-event) ..See WM_PAINT For Details
                  bOver = False
           End If
      End If
            
    ElseIf m_Adodc Then
        
        If MyClassObject.Enabled Then
          If CheckAdodcArea = True Then
             SetTimer m_hWnd, 1, 1, 0
             DrawControlClass
          End If
        End If
        
     ElseIf m_Slider Then
        
         If CheckSliderArea = True And bOver = False And bDown = False Then
            ButtonState = C_Over
            bOver = True
            DrawControlClass
         End If
           
         If CheckSliderArea = False And bOver = True And bDown = False Then
            ButtonState = C_Normal
            bOver = False
            DrawControlClass
         End If
     
     ElseIf m_ListView Then
     
          If ButtonState <> C_Down Then
             If CheckListViewArea(lParam) = True Then
                ButtonState = C_Over
                DrawControlClass
                SetTimer m_hWnd, 1, 1, 0
             End If
          End If
           
     ElseIf m_TabStrip Then
           
           If CheckTabStripArea(lParam) = True Then DrawControlClass
           SetTimer m_hWnd, 1, 1, 0
     
      ElseIf m_Combo Or m_ICombo Then

            If InsideArea(m_hWnd) = True Then
                    If bOver = False Then
                       bOver = True
                       ButtonState = C_Over
                       DrawControlClass
                       SetTimer m_hWnd, 1, 1, 0
                    End If
            Else
                    ButtonState = C_Down
            End If
            
       End If

'==========================================================================================
'==========================================================================================
'The WM_TIMER message is posted to the installing thread's message queue when a timer expires.
'You can process the message by providing a WM_TIMER case in the window procedure. Otherwise,
'the default window procedure will call the TimerProc callback function specified in the call
'to the SetTimer function used to install the timer.

    Case WM_TIMER
    
        
             If InsideArea(m_hWnd) = False Then
                 
                 If Not m_ListView Or m_ListView And ButtonState <> C_Down Then
                     If (m_Combo And ButtonState <> C_Down) Or (m_ICombo And ButtonState <> C_Down) Or (Not m_ICombo) Then ButtonState = C_Normal
                 End If
   
                 If m_TabStrip Then m_TSCTab = True
                 If m_Adodc Then m_CurrentButton = -1
   
                  
                  KillTimer m_hWnd, 1
                  bOver = False
                  
                  ButtonState = C_Normal
                  
                  If m_Button Then
                    PostPaintControl
                  Else
                    DrawControlClass
                  End If
                  
                  m_TSILNumber = SendMessageLong(m_hWnd, TCM_GETCURSEL, 0&, 0&)
                  m_TSINumber = m_TSILNumber
                  m_CurrentItem = -1
                  m_TSCTab = False
                        
            Else
            
                  bOver = True
                           
                  If ButtonState <> C_Down Then

                       If m_Slider Then
                           If CheckSliderArea = True Then ButtonState = C_Over
                           If CheckSliderArea = False Then ButtonState = C_Normal
                       End If
                     
                       '//--- Some Controls have Different Behavior when Mouse is Over
                       If Not m_Slider And Not m_Combo And Not m_ICombo Then ButtonState = C_Over
                     
                       If m_Combo Or m_ICombo Then
                            If DroppedDown = False Then
                               ButtonState = C_Over 'ComboBoxes Don't have C_Over (Mouse Move) When Dropped ;)
                            End If
                       End If

                   End If
         
                   If m_Adodc And ButtonState <> MovementFlag Then
                      MovementFlag = ButtonState
                      DrawControlClass
                   End If
              
              End If
          
   
   End Select
   
   
      
   
       
   
End Function

Private Sub PrePaintControl()
         DestDC = GetDC(m_hWnd)
         GetWindowRect m_hWnd, m_ItemRect
                m_Width = m_ItemRect.Right - m_ItemRect.Left
                m_Height = m_ItemRect.Bottom - m_ItemRect.Top
         CreateNewDCWorkArea m_Width, m_Height
         SelectBitmap
         m_PreDraw = True
         DrawControlClass
End Sub

Private Sub PostPaintControl()
         CreateBackMask m_Width, m_Height
         origBrush = SelectObject(TempDC, TempBrush)
         SelectObject TempDC, origBrush
        '------------------------------------------------------------------------------------------------------
        DOBitBlt m_Width, m_Height '//--- Do RasterOperations
        CleanDCs                    '//--- Free Memory <--Prevent Leaks
        m_PreDraw = False
        DrawControlClass
End Sub



'=======================================================================================================================
' CREATE A MASK COLOR BACKGROUND
'=======================================================================================================================

Private Sub CreateBackMask(ByVal m_Width As Long, ByVal m_Height As Long)
        
        origColor = SetBkColor(DestDC, GetSysColor(15))
        SetBkColor OrigDC, GetSysColor(15)
        BitBlt MaskDC, 0, 0, m_Width, m_Height, OrigDC, 0, 0, vbSrcCopy
       
End Sub


'=======================================================================================================================
' CREATE THE NEW TEMP WORK AREA
'=======================================================================================================================

Private Sub CreateNewDCWorkArea(ByVal m_Width As Long, ByVal m_Height As Long)
        
        MaskDC = CreateCompatibleDC(DestDC)
        MaskPic = CreateBitmap(m_Width, m_Height, 1, 1, ByVal 0&)
        MemDC = CreateCompatibleDC(DestDC)
        MemPic = CreateCompatibleBitmap(DestDC, m_Width, m_Height)
        TempDC = CreateCompatibleDC(DestDC)
        TempPic = CreateCompatibleBitmap(DestDC, m_Width, m_Height)
        OrigDC = CreateCompatibleDC(DestDC)
        OrigPic = CreateCompatibleBitmap(DestDC, m_Width, m_Height)

End Sub


'=======================================================================================================================
' BITBLT  RasterOperations
'=======================================================================================================================

Private Sub DOBitBlt(ByVal m_Width As Long, ByVal m_Height As Long)
        
        BitBlt MemDC, 0, 0, m_Width, m_Height, MaskDC, 0, 0, vbSrcCopy
        BitBlt MemDC, 0, 0, m_Width, m_Height, OrigDC, 0, 0, vbSrcPaint
        BitBlt TempDC, 0, 0, m_Width, m_Height, MaskDC, 0, 0, vbMergePaint
        BitBlt TempDC, 0, 0, m_Width, m_Height, MemDC, 0, 0, vbSrcAnd
        BitBlt DestDC, 0, 0, m_Width, m_Height, TempDC, 0, 0, vbSrcCopy

End Sub

'=======================================================================================================================
' CLEAN UP MEMORY
'=======================================================================================================================

Private Sub CleanDCs()
        
        DeleteDC TempDC
        DeleteObject TempPic
        DeleteDC MaskDC
        DeleteObject MaskPic
        DeleteDC MemDC
        DeleteObject MemPic
        DeleteDC OrigDC
        DeleteObject OrigPic
        DeleteObject TempBrush

End Sub


'=======================================================================================================================
' SELECT THE CURRENT IMAGE
'=======================================================================================================================

Private Sub SelectBitmap()
Dim cHandle As Long

       cHandle = SelectObject(MaskDC, MaskPic)
       DeleteObject cHandle
       cHandle = SelectObject(MemDC, MemPic)
       DeleteObject cHandle
       cHandle = SelectObject(TempDC, TempPic)
       DeleteObject cHandle
       cHandle = SelectObject(OrigDC, OrigPic)
       DeleteObject cHandle
       
End Sub


Private Function CheckAdodcArea() As Boolean
  Dim WinItem As RECT
  Dim Point As POINTAPI
  Dim hRgn As Long
  Dim NButton As Integer
   
    GetCursorPos Point
    
    
    GetWindowRect m_hWnd, WinItem
    WinItem.Left = WinItem.Left + 41
    WinItem.Right = WinItem.Right - 42
    NButton = -1
    GoSub CheckThisButton
    If CheckAdodcArea = True Then Exit Function
        
    GetWindowRect m_hWnd, WinItem
    WinItem.Right = WinItem.Left + 20
    NButton = 0
    GoSub CheckThisButton
    If CheckAdodcArea = True Then Exit Function
    
    GetWindowRect m_hWnd, WinItem
    WinItem.Right = WinItem.Left + 40
    WinItem.Left = WinItem.Left + 21
    NButton = 1
    GoSub CheckThisButton
    If CheckAdodcArea = True Then Exit Function
    
    GetWindowRect m_hWnd, WinItem
    WinItem.Left = WinItem.Right - 41
    WinItem.Right = WinItem.Right - 22
    NButton = 2
    GoSub CheckThisButton
    If CheckAdodcArea = True Then Exit Function
    
    GetWindowRect m_hWnd, WinItem
    WinItem.Left = WinItem.Right - 21
    NButton = 3
    GoSub CheckThisButton
    
           
   
   Exit Function

CheckThisButton:
    
    hRgn = CreateRectRgnIndirect(WinItem)
       
    If PtInRegion(hRgn, CLng(Point.X), CLng(Point.Y)) Then
        If NButton <> m_CurrentButton Then
            m_CurrentButton = NButton
            CheckAdodcArea = True
        End If
    Else
        CheckAdodcArea = False
    End If
        
        DeleteObject hRgn
      
    Return
    
End Function


Private Function CheckSliderArea() As Boolean
        
   Dim SliderRegion As RECT
   Dim RcItem As RECT
   Dim Point As POINTAPI
   Dim hRgn As Long
   
   
   SendMessageLong m_hWnd, TBM_GETTHUMBRECT, 0, SliderRegion
   
   GetCursorPos Point
   GetWindowRect m_hWnd, RcItem

   SliderRegion.Left = SliderRegion.Left + RcItem.Left
   SliderRegion.Right = SliderRegion.Right + RcItem.Left
   SliderRegion.Top = SliderRegion.Top + RcItem.Top
   SliderRegion.Bottom = SliderRegion.Bottom + RcItem.Top
   
    hRgn = CreateRectRgnIndirect(SliderRegion)
       
    If PtInRegion(hRgn, CLng(Point.X), CLng(Point.Y)) Then
        CheckSliderArea = True
    Else
        CheckSliderArea = False
    End If
        
    DeleteObject hRgn
        
End Function
'==========================================================================================
'==========================================================================================
'ListView Check Area
Private Function CheckListViewArea(ByVal lParam As Long) As Boolean
 
 Dim hti As HDHITTESTINFO
 Dim mLH As TLoHiLong, mAL As TAllLong
 
      mAL.All = lParam
      LSet mLH = mAL
      hti.pt.X = mLH.Lo
      hti.pt.Y = mLH.Hi
            
      SendMessageLong m_hWnd, HDM_HITTEST, 0&, hti
      SendMessageLong m_hWnd, HDM_GETITEMRECT, hti.iItem, m_LVIRect
     
      If hti.iItem <> m_CurrentItem Then
        m_CurrentItem = hti.iItem
        CheckListViewArea = True
      End If
      
End Function

'==========================================================================================
'==========================================================================================
'Tab Strip Check Area
Private Function CheckTabStripArea(ByVal lParam As Long) As Boolean
 
 Dim tTI As TCHITTESTINFO
 Dim iItem As Long
 Dim mLH As TLoHiLong, mAL As TAllLong
 
      mAL.All = lParam
      LSet mLH = mAL
      tTI.pt.X = mLH.Lo
      tTI.pt.Y = mLH.Hi
            
           
      iItem = SendMessageLong(m_hWnd, TCM_HITTEST, 0&, tTI)
             
     
       If iItem <> m_TSINumber Then
           m_TSILNumber = m_TSINumber
           m_TSINumber = iItem
           CheckTabStripArea = True
       End If
      
End Function

Private Sub PrintTabClientColor()
Dim TRec As RECT
Dim cHdc As Long
Dim Valh As Long

cHdc = GetDC(m_hWnd)

GetClientRect m_hWnd, TRec

Select Case m_ActiveScaleMode

Case vbTwips
     TRec.Top = Abs((MyClassObject.ClientTop - MyClassObject.Top) / Screen.TwipsPerPixelY)

Case vbPixels
     TRec.Top = Abs(MyClassObject.ClientTop - MyClassObject.Top)
     
Case vbPoints
     TRec.Top = (Abs(MyClassObject.ClientTop - MyClassObject.Top) * 100) / 75
     
End Select



DrawFillRectangle TRec, vbWhite, cHdc
   
ReleaseDC m_hWnd, cHdc


End Sub


Public Property Let ActiveScaleMode(ByRef cActiveScaleMode As ScaleModeConstants)
   m_ActiveScaleMode = cActiveScaleMode
End Property

Public Property Let IdeSubClass(ByVal cIdeSubClass As Boolean)
   m_IdeSubClass = cIdeSubClass
End Property

Public Property Get SchemeColor() As CWindowColors
   SchemeColor = m_SchemeColor
End Property

Public Property Let SchemeColor(ByRef cSchemeColor As CWindowColors)
   m_SchemeColor = cSchemeColor
End Property
