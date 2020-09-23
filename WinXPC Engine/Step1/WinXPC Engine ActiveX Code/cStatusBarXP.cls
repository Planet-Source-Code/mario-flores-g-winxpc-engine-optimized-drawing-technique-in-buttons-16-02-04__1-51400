VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStatusBarXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------------------------'
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
'                                          cStatusBarXP.cls                                        '
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

Private m_Hwnd As Long
Private m_hDC As Long
Private rcItem As RECT


Public Sub DrawBar()
Dim I As Long, II As Long, StepXP1 As Long, XPFace As Long

GetClientRect m_Hwnd, rcItem
CleanCornerArea rcItem


XPFace = ShiftColor(GetSysColor(15), -&H2, True)

   StepXP1 = 66 / 3
                    
   For I = 0 To 3
     DrawLine 0, Abs(I - 3), rcItem.Right, Abs(I - 3), m_hDC, ShiftColor(XPFace, -StepXP1 * ((((I / 3) * 100) * 3) / 100), True)
   Next I

'XPFace = ShiftColor(GetSysColor(15), -&H2, True)

   StepXP1 = 25 / rcItem.Bottom
                    
   For I = 0 To 6
     DrawLine 0, rcItem.Bottom - 6 + I, rcItem.Right, rcItem.Bottom - 6 + I, m_hDC, ShiftColor(XPFace, -StepXP1 * ((((I / 6) * 100) * rcItem.Bottom) / 100), True)
   Next I


For II = 3 To 1 Step -1
    For I = 2 To 3.3 * II Step 4
        CenterRectangle (rcItem.Right - 1) - I, (rcItem.Bottom - 3 - (Abs(II - 3) * 4)), vbWhite
        CenterRectangle (rcItem.Right - 2) - I, (rcItem.Bottom - 4 - (Abs(II - 3) * 4)), GetLngColor(&HA3B4B8)
    Next I
Next II

End Sub


Public Sub DrawPanel(ByVal X As Integer)
    DrawLine X, 5, X, rcItem.Bottom - 3, m_hDC, GetLngColor(vbGrayText)
    DrawLine X + 1, 5, X + 1, rcItem.Bottom - 3, m_hDC, GetLngColor(vbWhite)
End Sub

Private Sub CenterRectangle(ByVal X As Integer, ByVal Y As Integer, ByVal Color As Long)
Dim hBrush As Long
Dim hRect As RECT
    
    hRect.Top = Y
    hRect.Left = X
    hRect.Bottom = Y + 2
    hRect.Right = X + 2
    
    hBrush = CreateSolidBrush(Color)
    FillRect m_hDC, hRect, hBrush
    DeleteObject hBrush

End Sub

Private Sub CleanCornerArea(ByRef hRect As RECT)
Dim hBrush As Long
    
    hRect.Left = hRect.Right - 15
       
    hBrush = CreateSolidBrush(GetSysColor(15))
    FillRect m_hDC, hRect, hBrush
    DeleteObject hBrush

End Sub

Public Property Get hWnd() As Long
   hWnd = m_Hwnd
End Property

Public Property Let hWnd(ByVal cHwnd As Long)
   m_Hwnd = cHwnd
End Property

Public Property Get hdc() As Long
   hdc = m_hDC
End Property
Public Property Let hdc(ByVal cHdc As Long)
   m_hDC = cHdc
End Property

