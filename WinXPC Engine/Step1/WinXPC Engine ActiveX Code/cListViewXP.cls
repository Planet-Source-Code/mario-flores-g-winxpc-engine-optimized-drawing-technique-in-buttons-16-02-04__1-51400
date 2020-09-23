VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cListViewXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------------------------'
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
'                                           cListViewXP.cls                                        '
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

Option Explicit

Private m_hWnd As Long
Private m_Hdc As Long
Private m_Top As Long
Private m_Bottom As Long
Private m_Left As Long
Private m_Right As Long
Private m_State As ControlState
Private m_ItemCaption As String
Private m_ItemFont As String
Private m_ItemSize As Integer
Private m_Width As Integer
Private m_ItemItalic As Boolean
Private m_ItemUnderline As Boolean
Private RcItem As RECT

Public Sub DrawListView()


 
 RcItem.Left = m_Left
 RcItem.Top = m_Top
 RcItem.Right = m_Right
 RcItem.Bottom = m_Bottom

' Call CleanHeaders
' Call DrawDownBorder
    
    Select Case m_State

        Case C_Over
              
              CenterRectangle RcItem, vbWhite
              SetBkColor m_Hdc, vbBlue
              DrawLine m_Left, m_Bottom - 3, m_Right, m_Bottom - 3, m_Hdc, GetLngColor(&HA9F8&)
              DrawLine m_Left, m_Bottom - 2, m_Right, m_Bottom - 2, m_Hdc, GetLngColor(&H47C2FC)
              DrawLine m_Left, m_Bottom - 1, m_Right, m_Bottom - 1, m_Hdc, GetLngColor(&H1FB3F8)
              Call SetPixelV(m_Hdc, m_Left + 0, m_Bottom - 3, GetLngColor(&H68CCF0))
              Call SetPixelV(m_Hdc, m_Left + 1, m_Bottom - 3, GetLngColor(&H87D5FA))
              Call SetPixelV(m_Hdc, m_Left + 2, m_Bottom - 3, GetLngColor(&H9DDFFC))
              Call SetPixelV(m_Hdc, m_Left + 3, m_Bottom - 3, GetLngColor(&H70CFFB))
              Call SetPixelV(m_Hdc, m_Left + 0, m_Bottom - 2, GetLngColor(&HB5E0F2))
              Call SetPixelV(m_Hdc, m_Left + 1, m_Bottom - 2, GetLngColor(&H2FBAFB))
              Call SetPixelV(m_Hdc, m_Left + 2, m_Bottom - 2, GetLngColor(&HA6E0FC))
              Call SetPixelV(m_Hdc, m_Left + 3, m_Bottom - 2, GetLngColor(&HB2E4FA))
              Call SetPixelV(m_Hdc, m_Left + 4, m_Bottom - 2, GetLngColor(&H9CDCFA))
              Call SetPixelV(m_Hdc, m_Left + 5, m_Bottom - 2, GetLngColor(&H85D6FC))
              Call SetPixelV(m_Hdc, m_Left + 0, m_Bottom - 1, GetLngColor(&H9CDCFA))
              Call SetPixelV(m_Hdc, m_Left + 1, m_Bottom - 1, GetLngColor(&HC4E2EF))
              Call SetPixelV(m_Hdc, m_Right - 0, m_Bottom - 2, GetLngColor(&HC4E2EF))
              Call SetPixelV(m_Hdc, m_Right - 1, m_Bottom - 1, GetLngColor(&HC4E2EF))
              Call SetPixelV(m_Hdc, m_Right - 2, m_Bottom - 1, GetLngColor(&HC4E2EF))
   
        Case C_Down
             
              CenterRectangle RcItem, GetLngColor(&HD8DFDE)
              DrawRectangle RcItem, GetLngColor(&H97A5A5), m_Hdc
        Case Else
            'Nothing

    End Select

End Sub

Public Sub DrawDownBorder(TBottom As Long)
Dim i As Long, StepXP1 As Long, XPFace As Long
   
  XPFace = ShiftColor(GetSysColor(15), -&H2, True)
  StepXP1 = 55 / TBottom
                   
  For i = 0 To 4
       DrawLine 0, TBottom - 4 + i, m_Width, TBottom - 4 + i, m_Hdc, ShiftColor(XPFace, -StepXP1 * ((((i / 4) * 100) * TBottom) / 100), True)
  Next i
             
End Sub

Public Sub CleanHeaders(TBottom As Long)
Dim hBrush As Long
Dim BRect As RECT

BRect.Top = 0
BRect.Left = 0
BRect.Bottom = TBottom
BRect.Right = m_Width

    hBrush = CreateSolidBrush(GetSysColor(15))
    FillRect m_Hdc, BRect, hBrush
    DeleteObject hBrush

End Sub

Private Sub DrawPanel(ByVal X As Integer)
    DrawLine X - 3, 3, X - 3, RcItem.Bottom - 5, m_Hdc, GetLngColor(vbGrayText)
    DrawLine X - 2, 3, X - 2, RcItem.Bottom - 5, m_Hdc, GetLngColor(vbWhite)
End Sub


Public Sub SetTextParams()
   
   RcItem.Left = m_Left
   RcItem.Top = m_Top
   RcItem.Right = m_Right
   RcItem.Bottom = m_Bottom

   SelectFont m_Hdc, m_ItemSize, m_ItemItalic, m_ItemFont, m_ItemUnderline
   SetBkMode m_Hdc, 1
   DrawText m_Hdc, m_ItemCaption, Len(m_ItemCaption), RcItem, &H1
   Call DrawPanel(m_Right)
End Sub

Private Sub CenterRectangle(ByRef BRect As RECT, ByVal Color As Long)
Dim hBrush As Long
    
    hBrush = CreateSolidBrush(Color)
    InflateRect BRect, 1, 0
    FillRect m_Hdc, BRect, hBrush
    DeleteObject hBrush

End Sub


Public Property Let State(ByVal cState As ControlState)
   m_State = cState
End Property

Public Property Let ITop(ByVal cITop As Long)
   m_Top = cITop
End Property

Public Property Let IBottom(ByVal cIBottom As Long)
   m_Bottom = cIBottom
End Property

Public Property Let ILeft(ByVal cILeft As Long)
   m_Left = cILeft
End Property

Public Property Let IRight(ByVal cIRight As Long)
   m_Right = cIRight
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

Public Property Let Width(ByVal cWidth As Integer)
   m_Width = cWidth
End Property
