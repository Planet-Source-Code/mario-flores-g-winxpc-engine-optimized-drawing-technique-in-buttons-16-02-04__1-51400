VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTabStripXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_hWnd As Long
Private m_Hdc  As Long
Private m_ItemCaption As String
Private m_ItemFont As String
Private m_ItemSize As Integer
Private m_ItemItalic As Boolean
Private m_ItemUnderline As Boolean
Private m_CurrentItem As Integer
Private m_LastItem As Integer
Private m_SelectedItem As Integer
Private m_CountItems As Integer
Private m_CleanItem As Boolean
Private RcItem As RECT

Private m_ColorScheme As CWindowColors
Private ActualHighlight As Long

'//-----------------------------------------------
    Private m_TSIRect As RECT 'TabStripItemRect
'//-----------------------------------------------


Dim i As Long

Private Sub DrawPanelArea()
 
 
 SendMessageLong m_hWnd, TCM_GETITEMRECT, 0, m_TSIRect '//--Retrieve the bounding rectangle x of tab

 
     GetClientRect m_hWnd, RcItem
      
     DrawLine 0, 0, 0, RcItem.Bottom, m_Hdc, vbButtonFace
     DrawLine 1, m_TSIRect.Bottom + 1, 1, RcItem.Bottom - 3, m_Hdc, &H9C9B91
     
     DrawLine 1, RcItem.Bottom - 3, RcItem.Right, RcItem.Bottom - 3, m_Hdc, &H9C9B91
     DrawLine 1, RcItem.Bottom - 2, RcItem.Right, RcItem.Bottom - 2, m_Hdc, &HBFCED0
     DrawLine 1, RcItem.Bottom - 1, RcItem.Right, RcItem.Bottom - 1, m_Hdc, &HD0E0E3
     
     DrawLine RcItem.Right - 3, m_TSIRect.Bottom + 1, RcItem.Right - 3, RcItem.Bottom - 3, m_Hdc, &H9C9B91
     DrawLine RcItem.Right - 2, m_TSIRect.Bottom, RcItem.Right - 2, RcItem.Bottom - 1, m_Hdc, &HBFCED0
     DrawLine RcItem.Right - 1, m_TSIRect.Bottom, RcItem.Right - 1, RcItem.Bottom - 1, m_Hdc, &HD0E0E3
     
     DrawLine 1, m_TSIRect.Bottom + 1, RcItem.Right - 3, m_TSIRect.Bottom + 1, m_Hdc, &H9C9B91
     
     SendMessageLong m_hWnd, TCM_GETITEMRECT, m_SelectedItem, m_TSIRect
     
     DrawLine m_TSIRect.Left, m_TSIRect.Bottom, m_TSIRect.Right + 1, m_TSIRect.Bottom, m_Hdc, vbWhite
     DrawLine m_TSIRect.Left, m_TSIRect.Bottom + 1, m_TSIRect.Right + 1, m_TSIRect.Bottom + 1, m_Hdc, vbWhite
 
      SendMessageLong m_hWnd, TCM_GETITEMRECT, m_CountItems - 1, m_TSIRect '//--Retrieve the bounding rectangle x of tab
     DrawLine m_TSIRect.Right, m_TSIRect.Bottom, RcItem.Right, m_TSIRect.Bottom, m_Hdc, vbButtonFace
    
End Sub



Private Sub LastTabDraw()
Dim i As Long, StepXP1 As Single, XPFace As Long

If m_ColorScheme <> WindowsXP_Silver Then

 XPFace = ShiftColor(GetLngColor(vbButtonFace), &H30, True)
 StepXP1 = 25 / m_TSIRect.Bottom
 
 For i = m_TSIRect.Top To m_TSIRect.Bottom
        DrawLine m_TSIRect.Left, i + 1, m_TSIRect.Right + 1, i + 1, m_Hdc, ShiftColor(XPFace, -StepXP1 * i, True)
 Next i
 
Else
   
   DrawGradientMenu m_Hdc, m_TSIRect.Left, m_TSIRect.Top, m_TSIRect.Right - m_TSIRect.Left, m_TSIRect.Bottom, GetRGBColors(vbWhite), GetRGBColors(GetLngColor(&HD7C3C6)), GRADIENT_VERTICAL

End If


End Sub

Private Sub DrawCaption(ByVal iItem As Long)
Dim TabCaption As String
    

    TabCaption = GetTabText(iItem)
    If Len(Trim(TabCaption)) <> 0 Then
       SetBkMode m_Hdc, 1
       SelectFont m_Hdc, m_ItemSize, m_ItemItalic, m_ItemFont, m_ItemUnderline
       DrawText m_Hdc, TabCaption, Len(TabCaption), m_TSIRect, DT_SINGLELINE Or DT_VCENTER Or DT_CENTER
    End If
    
    
End Sub

Private Sub DrawImage(ByVal iItem As Long, ByVal bSelected As Boolean)
 
 Dim lImage As Long
 Dim Ihwnd  As Long
 Dim m_hIml As Long
 Dim tTI As TCITEM
 Dim cx As Long, cy As Long
 Dim lX As Long, lY As Long
 
 
   lImage = GetTabImage(iItem)
   If lImage > -1 Then
      m_hIml = SendMessageLong(m_hWnd, TCM_GETIMAGELIST, Ihwnd, tTI)
           
      If Not m_hIml = 0 Then
         ImageList_GetIconSize m_hIml, cx, cy
         lY = m_TSIRect.Top + (m_TSIRect.Bottom - m_TSIRect.Top - cy) \ 2
         lX = m_TSIRect.Left + 6
         If Not bSelected Then lY = m_TSIRect.Top + 3
         ImageList_Draw m_hIml, lImage, m_Hdc, lX, lY, 1
         m_TSIRect.Left = lX + cx + 1
      End If
   End If

End Sub

Public Sub DrawTabStrip()



m_SelectedItem = SendMessageLong(m_hWnd, TCM_GETCURSEL, 0&, 0&) '//--Currently selected tab in a tab control
m_CountItems = SendMessageLong(m_hWnd, TCM_GETITEMCOUNT, 0&, 0&) '//--Retrieve the number of tabs in the tab control

If m_CurrentItem > m_CountItems Or m_CurrentItem < 0 Or m_LastItem > m_CountItems Or m_LastItem < 0 Then Exit Sub
   
 
 
If m_SelectedItem <> m_LastItem Then
    SendMessageLong m_hWnd, TCM_GETITEMRECT, m_LastItem, m_TSIRect '//--Retrieve the bounding rectangle of tab
    LastTabDraw
    DrawFrameTab False
    DrawImage m_LastItem, False
    DrawCaption m_LastItem
End If


SendMessageLong m_hWnd, TCM_GETITEMRECT, m_CurrentItem, m_TSIRect '//--Retrieve the bounding rectangle of tab
    
If m_SelectedItem <> m_CurrentItem Then
   LastTabDraw
   If m_CleanItem = False Then
   DrawFrameTab True
   Else
   DrawFrameTab False
   End If
   
   DrawImage m_CurrentItem, False
   DrawCaption m_CurrentItem
 
Else
   m_TSIRect.Right = m_TSIRect.Right + 1
   DrawFillRectangle m_TSIRect, vbWhite, m_Hdc
   m_TSIRect.Right = m_TSIRect.Right - 1
   DrawFrameTab True, True
   DrawImage m_CurrentItem, True
   DrawCaption m_CurrentItem
End If
 
 DrawPanelArea

          
End Sub

Private Sub DrawFrameTab(ByVal DrawOver As Boolean, Optional Selected As Boolean = False)
Dim Color1 As Long, Color2 As Long, Color3 As Long
Dim SDif As Integer '//--Size Difference in Selected Tab (Raised)

'========================================================================================
'Theme Support Colors
Dim TCol1 As Long, TCol2 As Long

 If Selected Then
    SDif = 2 '//--Raise Tab A Little Bit ;)
 Else
    SDif = 0 '//--Draw Normal Size
 End If

 TCol1 = IIf(m_ColorScheme = WindowsXP_OliveGreen, &H5896E3, &H3CC7FF)
 TCol2 = IIf(m_ColorScheme = WindowsXP_OliveGreen, &H9CAC9B, &HB4A791)

 
 If DrawOver Then
    Color1 = &H8DC2DD: Color2 = &H3C97E4: Color3 = &H53A1DD
    DrawLine m_TSIRect.Left + 1, m_TSIRect.Top - SDif, m_TSIRect.Right - 1 + SDif, m_TSIRect.Top - SDif, m_Hdc, &H288BE6
    DrawLine m_TSIRect.Left + 1, m_TSIRect.Top + 1 - SDif, m_TSIRect.Right - 1 + SDif, m_TSIRect.Top + 1 - SDif, m_Hdc, TCol1
    DrawLine m_TSIRect.Left + 1 - Abs(SDif / 3), m_TSIRect.Top + 2 - SDif, m_TSIRect.Right - 1 + Abs(SDif / 3) + SDif, m_TSIRect.Top + 2 - SDif, m_Hdc, TCol1
 Else
    Color1 = IIf(m_ColorScheme = WindowsXP_OliveGreen, &HC4DBD7, &HD0D6D0)
    Color2 = IIf(m_ColorScheme = WindowsXP_OliveGreen, &HB2B7B9, &HBFB9A9)
    Color3 = IIf(m_ColorScheme = WindowsXP_OliveGreen, &HC1C3C5, &HCDC4B6)
    DrawLine m_TSIRect.Left + 1, m_TSIRect.Top, m_TSIRect.Right - 1, m_TSIRect.Top, m_Hdc, TCol2
 End If
    
       
 DrawLine m_TSIRect.Left + 1 - SDif, m_TSIRect.Top + 1, m_TSIRect.Left + 1 - SDif, m_TSIRect.Bottom + 1, m_Hdc, TCol2
 DrawLine m_TSIRect.Right - 1 + SDif, m_TSIRect.Top + 1, m_TSIRect.Right - 1 + SDif, m_TSIRect.Bottom + 1, m_Hdc, TCol2
 
 
    
 '//=======================================================================================
 '                               Do the Round Tab's Effect
 '                                     (UP Corners)
    
    Call SetPixelV(m_Hdc, m_TSIRect.Left + 1 - SDif, m_TSIRect.Top - SDif, GetSysColor(15))
    Call SetPixelV(m_Hdc, m_TSIRect.Right - 1 + SDif, m_TSIRect.Top - SDif, GetSysColor(15))
    Call SetPixelV(m_Hdc, m_TSIRect.Left + 2 - SDif, m_TSIRect.Top - SDif, Color1)
    Call SetPixelV(m_Hdc, m_TSIRect.Left + 3 - SDif, m_TSIRect.Top - SDif, Color2)
    Call SetPixelV(m_Hdc, m_TSIRect.Left + 1 - SDif, m_TSIRect.Top + 1 - SDif, Color1)
    Call SetPixelV(m_Hdc, m_TSIRect.Left + 2 - SDif, m_TSIRect.Top + 1 - SDif, Color3)
    Call SetPixelV(m_Hdc, m_TSIRect.Left + 1 - SDif, m_TSIRect.Top + 2 - SDif, Color2)
    Call SetPixelV(m_Hdc, m_TSIRect.Right - 2 + SDif, m_TSIRect.Top - SDif, Color1)
    Call SetPixelV(m_Hdc, m_TSIRect.Right - 3 + SDif, m_TSIRect.Top - SDif, Color2)
    Call SetPixelV(m_Hdc, m_TSIRect.Right - 1 + SDif, m_TSIRect.Top + 1 - SDif, Color1)
    Call SetPixelV(m_Hdc, m_TSIRect.Right - 2 + SDif, m_TSIRect.Top + 1 - SDif, Color3)
    Call SetPixelV(m_Hdc, m_TSIRect.Right - 1 + SDif, m_TSIRect.Top + 2 - SDif, Color2)
    
 
 
End Sub

 Private Function GetTabText(ByVal iItem As Long) As String

   Dim tTI As TCITEM
   Dim lR As Long
   Dim sText As String
          
   tTI.cchTextMax = 255
   tTI.pszText = String$(255, 0)
   tTI.mask = TCIF_TEXT
   lR = SendMessageLong(m_hWnd, TCM_GETITEMA, iItem, tTI)
   If (lR <> 0) Then
      sText = tTI.pszText
      lR = InStr(sText, Chr$(0))
      If (lR <> 0) Then
         GetTabText = Left$(sText, lR - 1)
      Else
         GetTabText = sText
      End If
   End If
 
 End Function

 Private Function GetTabImage(ByVal iItem As Long) As Long

   Dim tTI As TCITEM
   Dim lR As Long
        
      tTI.mask = TCIF_IMAGE
      lR = SendMessageLong(m_hWnd, TCM_GETITEMA, iItem, tTI)
      If (lR <> 0) Then GetTabImage = tTI.iImage
   
 End Function
Private Sub SchemeControl()

    Select Case m_ColorScheme
              
        Case SystemColors, WindowsXP_Blue
             ActualHighlight = GetLngColor(XPBlue_Highlight)
        Case WindowsXP_OliveGreen
             ActualHighlight = GetLngColor(XPGreen_Highlight)
    End Select
    
    
End Sub

Public Property Let ColorScheme(ByRef cColorScheme As CWindowColors)
   m_ColorScheme = cColorScheme
End Property
Public Property Let CurrentItem(ByVal cCurrentItem As Integer)
   m_CurrentItem = cCurrentItem
End Property

Public Property Let LastItem(ByVal cLastItem As Integer)
   m_LastItem = cLastItem
End Property

Public Property Let hwnd(ByVal cHwnd As Long)
   m_hWnd = cHwnd
End Property

Public Property Let hdc(ByVal cHdc As Long)
   m_Hdc = cHdc
End Property

Public Property Let ItemCaption(ByVal cItemCaption As String)
   m_ItemCaption = cItemCaption
End Property

Public Property Let ItemFont(ByVal cItemFont As String)
   m_ItemFont = cItemFont
End Property

Public Property Let ItemSize(ByVal cItemSize As Integer)
   m_ItemSize = cItemSize
End Property

Public Property Let ItemUnderline(ByVal cItemUnderline As Boolean)
   m_ItemUnderline = cItemUnderline
End Property

Public Property Let ItemItalic(ByVal cItemItalic As Boolean)
   m_ItemItalic = cItemItalic
End Property

Public Property Let CleanItem(ByVal cCleanItem As Boolean)
   m_CleanItem = cCleanItem
End Property

