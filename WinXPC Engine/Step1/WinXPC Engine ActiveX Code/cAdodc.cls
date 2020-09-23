VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAdodc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------------------------'
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
'                                             cADODC.cls                                           '
'                                            Version 1.01                                          '
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

Private m_hWnd          As Long
Private m_Hdc           As Long
Private m_cButton       As Integer
Private m_ColorScheme   As CWindowColors
Private RcItem          As RECT
Private m_Down          As Boolean
Private m_Enabled       As Boolean
Private ActualHighlight As Long
Private ActualGrayText  As Long

Public Sub DrawAdodc()

Dim TempRect As RECT

GetWindowRect m_hWnd, RcItem

'==========================================================
'//---Clean Old Buttons
Let TempRect = RcItem
TempRect.Right = TempRect.Left + 40
DrawFillRectangle TempRect, vbButtonFace, m_Hdc
'----
Let TempRect = RcItem
TempRect.Left = TempRect.Right - 41
DrawFillRectangle TempRect, vbButtonFace, m_Hdc


     Select Case m_cButton

        Case 0
            TempRect.Left = 2: TempRect.Right = 20: TempRect.Top = 2: TempRect.Bottom = TempRect.Bottom - 2
        Case 1
            TempRect.Left = 21: TempRect.Right = 40: TempRect.Top = 2: TempRect.Bottom = TempRect.Bottom - 2
        Case 2
            TempRect.Left = TempRect.Right - 41: TempRect.Right = TempRect.Right - 22: TempRect.Top = 2: TempRect.Bottom = TempRect.Bottom - 2
        Case 3
            TempRect.Left = TempRect.Right - 21: TempRect.Right = TempRect.Right - 2: TempRect.Top = 2: TempRect.Bottom = TempRect.Bottom - 2
        Case Else
            'Nothing
    End Select

'Debug.Print Now

    If m_cButton <> -1 Then
     
         SchemeControl
        If Not m_Down Then
            DrawFillRectangle TempRect, ShiftColorOXP(GetLngColor(ActualHighlight), 200), m_Hdc
            DrawRectangle TempRect, ShiftColorOXP(GetLngColor(ActualHighlight), 95), m_Hdc
        Else
            DrawFillRectangle TempRect, ShiftColorOXP(GetLngColor(ActualHighlight), 100), m_Hdc
            DrawRectangle TempRect, ShiftColorOXP(GetLngColor(ActualHighlight), 35), m_Hdc
        End If
    End If

'==========================================================
'//---Draw New Arrows in Buttons
Call DrawArrow(RcItem.Left + 10, ((RcItem.Top + RcItem.Bottom) / 2) - 1, 0, IIf(m_Down And m_cButton = 0 And m_Enabled, vbWhite, IIf(m_Enabled, vbBlack, XPBlue_GrayText)))
Call DrawArrow(RcItem.Left + 27, ((RcItem.Top + RcItem.Bottom) / 2) - 1, 1, IIf(m_Down And m_cButton = 1 And m_Enabled, vbWhite, IIf(m_Enabled, vbBlack, XPBlue_GrayText)))
Call DrawArrow(RcItem.Right - 17, ((RcItem.Top + RcItem.Bottom) / 2) - 1, 2, IIf(m_Down And m_cButton = 3 And m_Enabled, vbWhite, IIf(m_Enabled, vbBlack, XPBlue_GrayText)))
Call DrawArrow(RcItem.Right - 34, ((RcItem.Top + RcItem.Bottom) / 2) - 1, 3, IIf(m_Down And m_cButton = 2 And m_Enabled, vbWhite, IIf(m_Enabled, vbBlack, XPBlue_GrayText)))


'============================================================
'//---Draw The New Window Border
DrawRectangle RcItem, GetLngColor(vbActiveBorder), m_Hdc
Let TempRect = RcItem: InflateRect TempRect, -1, -1
DrawRectangle TempRect, vbWhite, m_Hdc

End Sub

Private Sub DrawArrow(X As Long, Y As Long, cType As Byte, Optional cColor As Long = vbBlack)

Dim J1 As Integer, J2 As Integer
                
           
    If cType = 0 Or cType = 1 Then
        
        For J1 = X To X + 6
          DrawLine J1, Y - Abs((J1 + 6) - (X + 6)), X + 6, Y - Abs((J1 + 6) - (X + 6)), m_Hdc, cColor
          DrawLine J1, Y + Abs((J1 + 6) - (X + 6)), X + 6, Y + Abs((J1 + 6) - (X + 6)), m_Hdc, cColor
        Next J1
    End If
           
    If cType = 0 Then
        DrawLine X - 3, Y - 4, X - 3, Y + 5, m_Hdc, cColor
        DrawLine X - 2, Y - 4, X - 2, Y + 5, m_Hdc, cColor
    End If
 
    If cType = 2 Or cType = 3 Then
        For J1 = X To X + 6
          DrawLine X, Y - Abs(J1 - (X + 6)), J1, Y - Abs(J1 - (X + 6)), m_Hdc, cColor
          DrawLine X, Y + Abs(J1 - (X + 6)), J1, Y + Abs(J1 - (X + 6)), m_Hdc, cColor
        Next J1
    End If
 
    If cType = 2 Then
        DrawLine X + 8, Y - 4, X + 8, Y + 5, m_Hdc, cColor
        DrawLine X + 7, Y - 4, X + 7, Y + 5, m_Hdc, cColor
    End If

 
  
End Sub
Private Sub SchemeControl()

    Select Case m_ColorScheme
              
        Case SystemColors
             ActualHighlight = vbHighlight
             ActualGrayText = vbGrayText
        Case WindowsXP_Blue
             ActualHighlight = XPBlue_Highlight
             ActualGrayText = XPBlue_GrayText
        Case WindowsXP_OliveGreen
             ActualHighlight = XPGreen_Highlight
             ActualGrayText = XPBlue_GrayText
    End Select
    
    
End Sub

Public Property Let Down(ByVal cDown As Boolean)
   m_Down = cDown
End Property

Public Property Let Enabled(ByVal cEnabled As Boolean)
   m_Enabled = cEnabled
End Property

Public Property Let ColorScheme(ByRef cColorScheme As CWindowColors)
   m_ColorScheme = cColorScheme
End Property

Public Property Let cButton(ByVal ccButton As Integer)
   m_cButton = ccButton
End Property

Public Property Let hwnd(ByVal cHwnd As Long)
   m_hWnd = cHwnd
End Property


Public Property Let hdc(ByVal cHdc As Long)
   m_Hdc = cHdc
End Property

