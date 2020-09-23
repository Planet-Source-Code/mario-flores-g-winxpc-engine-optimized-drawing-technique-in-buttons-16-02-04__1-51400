VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTextBoxXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------------------------'
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
'                                           cTextBoxXP.cls                                         '
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
Private m_ColorScheme As CWindowColors
Private RcItem As RECT

Private ActualHighlight As Long
Private ActualGrayText As Long


Public Sub DrawTextBoxXP()

Dim Outline As Long
Dim RcItem As RECT

    RcItem.Left = m_Left
    RcItem.Top = m_Top
    RcItem.Right = m_Right
    RcItem.Bottom = m_Bottom

Call SchemeControl '//--Select Colors
 


Select Case m_State
    Case 0, 1, 2, 3
        DrawRectangle RcItem, ShiftColorOXP(GetLngColor(ActualHighlight), 95), m_Hdc '&HB99D7F
    Case 4
        DrawRectangle RcItem, ShiftColorOXP(GetLngColor(ActualGrayText), 80), m_Hdc
    Case Else
         'Nothing
End Select
   
      
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
        Case WindowsXP_Silver
             ActualHighlight = XPSilver_Highlight
             ActualGrayText = XPBlue_GrayText
    End Select
    
    
End Sub

Public Property Let ColorScheme(ByRef cColorScheme As CWindowColors)
   m_ColorScheme = cColorScheme
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

Public Property Let State(ByVal cState As ControlState)
   m_State = cState
End Property

Public Property Let hwnd(ByVal cHwnd As Long)
   m_hWnd = cHwnd
End Property

Public Property Let hdc(ByVal cHdc As Long)
   m_Hdc = cHdc
End Property
