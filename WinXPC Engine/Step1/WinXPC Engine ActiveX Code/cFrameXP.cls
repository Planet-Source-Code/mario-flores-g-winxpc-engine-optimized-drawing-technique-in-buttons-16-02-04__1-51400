VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFrameXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------------------------------------'
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
'                                            cFrameXP.cls                                          '
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
Private m_hDC As Long
Private m_FontItalic As Boolean
Private m_FontUnderline As Boolean
Private m_FontSize As Integer
Private m_FrameFont As String
Private m_FrameCaption As String
Private m_ForeColor As OLE_COLOR
Private m_BackColor As OLE_COLOR
Private m_RealTextSize As Long
Private RcItem As RECT


Public Sub DrawFrame()

m_RealTextSize = (m_FontSize * 20) / Screen.TwipsPerPixelY

GetClientRect m_hWnd, RcItem: CleanCornerArea RcItem
GetClientRect m_hWnd, RcItem: MakeRegion RcItem, m_hWnd
DrawFrameArea RcItem

If Len(m_FrameCaption) > 0 Then DrawFrameText


End Sub


Private Sub DrawFrameText()

RcItem.Left = 9
RcItem.Top = 0
RcItem.Bottom = m_RealTextSize


SelectFont m_hDC, m_FontSize, m_FontItalic, m_FrameFont, m_FontUnderline
SetTextColor m_hDC, m_ForeColor
SetBkMode m_hDC, 2
SetBkColor m_hDC, GetLngColor(m_BackColor)

DrawText m_hDC, m_FrameCaption, Len(m_FrameCaption), RcItem, 0


End Sub

Private Sub DrawFrameArea(ByRef hRect As RECT)
hRect.Top = m_RealTextSize / 2

DrawRectangle hRect, GetLngColor(vbActiveBorder), m_hDC
Call SetPixelV(m_hDC, 0, hRect.Top, GetLngColor(m_BackColor))
Call SetPixelV(m_hDC, 1, hRect.Top, GetLngColor(m_BackColor))
Call SetPixelV(m_hDC, 0, hRect.Top + 1, GetLngColor(m_BackColor))
Call SetPixelV(m_hDC, hRect.Right - 1, hRect.Top, GetLngColor(m_BackColor))  '//--Clean Up Corners
Call SetPixelV(m_hDC, hRect.Right - 2, hRect.Top, GetLngColor(m_BackColor))
Call SetPixelV(m_hDC, hRect.Right - 1, hRect.Top + 1, GetLngColor(m_BackColor))

Call SetPixelV(m_hDC, 1, hRect.Top + 1, GetLngColor(vbActiveBorder))
Call SetPixelV(m_hDC, 2, hRect.Top + 1, GetLngColor(vbActiveBorder))  '//--UpLeftCorner
Call SetPixelV(m_hDC, 1, hRect.Top + 2, GetLngColor(vbActiveBorder))

Call SetPixelV(m_hDC, 1, hRect.Bottom - 3, GetLngColor(vbActiveBorder))
Call SetPixelV(m_hDC, 1, hRect.Bottom - 2, GetLngColor(vbActiveBorder)) '//--DownLeftCorner
Call SetPixelV(m_hDC, 2, hRect.Bottom - 2, GetLngColor(vbActiveBorder))

Call SetPixelV(m_hDC, hRect.Right - 3, hRect.Top + 1, GetLngColor(vbActiveBorder))
Call SetPixelV(m_hDC, hRect.Right - 2, hRect.Top + 1, GetLngColor(vbActiveBorder))  '//--UpRightCorner
Call SetPixelV(m_hDC, hRect.Right - 2, hRect.Top + 2, GetLngColor(vbActiveBorder))

Call SetPixelV(m_hDC, hRect.Right - 2, hRect.Bottom - 3, GetLngColor(vbActiveBorder))
Call SetPixelV(m_hDC, hRect.Right - 2, hRect.Bottom - 2, GetLngColor(vbActiveBorder)) '//--DownRightCorner
Call SetPixelV(m_hDC, hRect.Right - 3, hRect.Bottom - 2, GetLngColor(vbActiveBorder))

End Sub


Private Sub CleanCornerArea(ByRef hRect As RECT)
     
    DrawLine 0, 6, 0, hRect.Bottom, m_hDC, m_BackColor
    DrawLine 1, 6, 1, hRect.Bottom, m_hDC, m_BackColor
    DrawLine hRect.Right - 1, 6, hRect.Right - 1, hRect.Bottom, m_hDC, m_BackColor
    DrawLine hRect.Right - 2, 6, hRect.Right - 2, hRect.Bottom, m_hDC, m_BackColor
    DrawLine 0, hRect.Bottom - 1, hRect.Right - 1, hRect.Bottom - 1, m_hDC, m_BackColor
    DrawLine 0, hRect.Bottom - 2, hRect.Right - 1, hRect.Bottom - 2, m_hDC, m_BackColor
    
    hRect.Bottom = IIf(Len(m_FrameCaption) > 0, m_RealTextSize, 8)
    DrawFillRectangle hRect, m_BackColor, m_hDC
   
End Sub


Public Property Let Backcolor(ByVal cBackcolor As OLE_COLOR)
   m_BackColor = cBackcolor
End Property

Public Property Let Forecolor(ByVal cForecolor As OLE_COLOR)
   m_ForeColor = cForecolor
End Property

Public Property Let FrameCaption(ByVal cFrameCaption As String)
   m_FrameCaption = cFrameCaption
End Property

Public Property Let FrameFont(ByVal cFrameFont As String)
   m_FrameFont = cFrameFont
End Property

Public Property Let FontSize(ByVal cFontSize As Integer)
   m_FontSize = cFontSize
End Property

Public Property Let FontUnderline(ByVal cFontUnderline As Boolean)
   m_FontUnderline = cFontUnderline
End Property

Public Property Let FontItalic(ByVal cFontItalic As Boolean)
   m_FontItalic = cFontItalic
End Property

Public Property Let hwnd(ByVal cHwnd As Long)
   m_hWnd = cHwnd
End Property

Public Property Let hdc(ByVal cHdc As Long)
   m_hDC = cHdc
End Property


