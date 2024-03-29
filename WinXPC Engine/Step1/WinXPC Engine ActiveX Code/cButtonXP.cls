VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cButtonXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_hWnd      As Long
Private m_Hdc       As Long
Private m_Enabled   As Boolean
Private m_Focused   As Boolean
Private m_Down      As Boolean
Private m_Over      As Boolean
Private m_PreDraw   As Boolean
Private RcItem As RECT
Private m_ColorScheme As CWindowColors
Private cFaceO As Long, XPFace As Long, OXPf As Long


'=========================================================================================
'=========================================================================================
'                                    Scheme Color Support
'Color Hover Variables
 Private BColor As Long 'Border Color
 Private HCol1 As Long, HCol2 As Long, HCol3 As Long, HCol4 As Long, HCol5 As Long 'Hover Colors
 Private FCol1 As Long, FCol2 As Long, FCol3 As Long, FCol4 As Long, FCol5 As Long 'Focus Colors



'//=========================================================================================================
'                                    THIS SUB DRAWS THE XP STYLED BUTTON
'//=========================================================================================================

Public Sub DrawButtonXP()

Dim i As Long, StepXP1 As Single, XPFace2 As Long, tempCol As Long
Dim TempRec As RECT
    
    Call SchemeControl
    
    GetClientRect m_hWnd, RcItem
       
    XPFace = ShiftColor(GetLngColor(vbButtonFace), &H30, True)
    OXPf = ShiftColorOXP(GetLngColor(vbHighlight))
    
   
'//=========================================================================================================
'                                         CHECK FOR DISABLED BUTTON

         If m_Enabled = False Then
                  TempRec.Left = 0: TempRec.Top = 0: TempRec.Right = RcItem.Right - 1: TempRec.Bottom = RcItem.Bottom
                  DrawRectangle TempRec, GetLngColor(vbButtonFace), m_Hdc
                  TempRec.Left = 0: TempRec.Top = RcItem.Bottom - 1: TempRec.Right = RcItem.Right: TempRec.Bottom = RcItem.Bottom - 1
                  DrawRectangle TempRec, GetLngColor(vbButtonFace), m_Hdc

                  
                  TempRec.Left = 0: TempRec.Top = 0: TempRec.Right = RcItem.Right: TempRec.Bottom = RcItem.Bottom
                  DrawRectangle TempRec, ShiftColor(XPFace, -&H18, True), m_Hdc
                  TempRec.Left = 0: TempRec.Top = 0: TempRec.Right = RcItem.Right: TempRec.Bottom = RcItem.Bottom
                  DrawRectangle TempRec, ShiftColor(XPFace, -&H54, True), m_Hdc
                  Call SetPixelV(m_Hdc, 1, 1, ShiftColor(XPFace, -&H48, True))
                  Call SetPixelV(m_Hdc, 1, RcItem.Bottom - 2, ShiftColor(XPFace, -&H48, True))
                  Call SetPixelV(m_Hdc, RcItem.Right - 2, 1, ShiftColor(XPFace, -&H48, True))
                  Call SetPixelV(m_Hdc, RcItem.Right - 2, RcItem.Bottom - 2, ShiftColor(XPFace, -&H48, True))
                  GoTo CutButton
                  Exit Sub
         End If
         
'//=========================================================================================================
'                                         CHECK FOR PRESSED BUTTON

         If m_Down = True Then
             
             StepXP1 = 25 / RcItem.Bottom
             XPFace2 = ShiftColor(XPFace, -32, True)
               
            If m_PreDraw Then
                   
                           
                   If m_ColorScheme <> WindowsXP_Silver Then
                      For i = 1 To RcItem.Bottom
                        DrawLine 0, RcItem.Bottom - i, RcItem.Right, RcItem.Bottom - i, m_Hdc, ShiftColor(XPFace2, -StepXP1 * i, True)
                      Next i
                   Else
                      DrawGradientMenu m_Hdc, 0, 0, RcItem.Right, RcItem.Bottom, GetRGBColors(GetLngColor(&HD7C3C6)), GetRGBColors(vbWhite), GRADIENT_VERTICAL
                   End If
                   
            Else
    
                   TempRec.Left = 0: TempRec.Top = 0: TempRec.Right = RcItem.Right: TempRec.Bottom = RcItem.Bottom
                   DrawRectangle TempRec, &H733C00, m_Hdc
                   Call SetPixelV(m_Hdc, 1, 1, GetLngColor(&H7B4D10))
                   Call SetPixelV(m_Hdc, 1, RcItem.Bottom - 2, GetLngColor(&H7B4D10))
                   Call SetPixelV(m_Hdc, RcItem.Right - 2, 1, GetLngColor(&H7B4D10))
                   Call SetPixelV(m_Hdc, RcItem.Right - 2, RcItem.Bottom - 2, GetLngColor(&H7B4D10))
                   DrawLine 2, RcItem.Bottom - 2, RcItem.Right - 2, RcItem.Bottom - 2, m_Hdc, ShiftColor(XPFace2, &H10, True)
                   DrawLine 1, RcItem.Bottom - 3, RcItem.Right - 2, RcItem.Bottom - 3, m_Hdc, ShiftColor(XPFace2, &HA, True)
                   DrawLine RcItem.Right - 2, 2, RcItem.Right - 2, RcItem.Bottom - 2, m_Hdc, ShiftColor(XPFace2, &H5, True)
                   DrawLine RcItem.Right - 3, 3, RcItem.Right - 3, RcItem.Bottom - 3, m_Hdc, XPFace
                   DrawLine 2, 1, RcItem.Right - 2, 1, m_Hdc, ShiftColor(XPFace2, -&H20, True)
                   DrawLine 1, 2, RcItem.Right - 2, 2, m_Hdc, ShiftColor(XPFace2, -&H18, True)
                   DrawLine 1, 2, 1, RcItem.Bottom - 2, m_Hdc, ShiftColor(XPFace2, -&H20, True)
                   DrawLine 2, 2, 2, RcItem.Bottom - 2, m_Hdc, ShiftColor(XPFace2, -&H16, True)
                   GoSub PaintNormalPost
                   GoTo CutButton
            
            End If
            
            Exit Sub
          
          
          
          Else

'//=========================================================================================================
'                                         CHECK FOR FOCUSED STATE
                  
                  If m_Focused And Not m_Over Then
                    If m_PreDraw Then
                        GoSub PaintNormalBefore
                    Else
                        TempRec.Left = 1: TempRec.Top = 2: TempRec.Right = RcItem.Right - 1: TempRec.Bottom = RcItem.Bottom - 2
                        DrawRectangle TempRec, FCol1, m_Hdc
                        DrawLine 2, RcItem.Bottom - 2, RcItem.Right - 2, RcItem.Bottom - 2, m_Hdc, FCol2
                        DrawLine 2, 1, RcItem.Right - 2, 1, m_Hdc, FCol3
                        DrawLine 1, 2, RcItem.Right - 1, 2, m_Hdc, FCol4
                        DrawLine 2, 3, 2, RcItem.Bottom - 3, m_Hdc, FCol5
                        DrawLine RcItem.Right - 3, 3, RcItem.Right - 3, RcItem.Bottom - 3, m_Hdc, FCol5
                        GoSub PaintNormalPost
                        GoTo CutButton
                    End If
                        Exit Sub
                 
'//=========================================================================================================
'                                         CHECK FOR HOVER STATE
                   ElseIf m_Over Then
                       
                            If m_PreDraw Then
                                GoSub PaintNormalBefore
                            Else
                                TempRec.Left = 1: TempRec.Top = 2: TempRec.Right = RcItem.Right - 1: TempRec.Bottom = RcItem.Bottom - 2
                                DrawRectangle TempRec, HCol1, m_Hdc
                                DrawLine 2, RcItem.Bottom - 2, RcItem.Right - 2, RcItem.Bottom - 2, m_Hdc, HCol2
                                DrawLine 2, 1, RcItem.Right - 2, 1, m_Hdc, HCol3
                                DrawLine 1, 2, RcItem.Right - 1, 2, m_Hdc, HCol4
                                DrawLine 2, 3, 2, RcItem.Bottom - 3, m_Hdc, HCol5
                                DrawLine RcItem.Right - 3, 3, RcItem.Right - 3, RcItem.Bottom - 3, m_Hdc, HCol5
                                GoSub PaintNormalPost
                                GoTo CutButton
                            End If
                                Exit Sub
                   Else
'//=========================================================================================================
                                          'NORMAL STATE NONE-ABOVE
                                
                            If m_PreDraw Then
                                GoSub PaintNormalBefore
                            Else
                                DrawLine 2, RcItem.Bottom - 2, RcItem.Right - 2, RcItem.Bottom - 2, m_Hdc, ShiftColor(XPFace, -&H30, True)
                                DrawLine 1, RcItem.Bottom - 3, RcItem.Right - 2, RcItem.Bottom - 3, m_Hdc, ShiftColor(XPFace, -&H20, True)
                                DrawLine RcItem.Right - 2, 2, RcItem.Right - 2, RcItem.Bottom - 2, m_Hdc, ShiftColor(XPFace, -&H24, True)
                                DrawLine RcItem.Right - 3, 3, RcItem.Right - 3, RcItem.Bottom - 3, m_Hdc, ShiftColor(XPFace, -&H18, True)
                                DrawLine 2, 1, RcItem.Right - 2, 1, m_Hdc, ShiftColor(XPFace, &H10, True)
                                DrawLine 1, 2, RcItem.Right - 2, 2, m_Hdc, ShiftColor(XPFace, &HA, True)
                                DrawLine 1, 2, 1, RcItem.Bottom - 2, m_Hdc, ShiftColor(XPFace, -&H5, True)
                                DrawLine 2, 3, 2, RcItem.Bottom - 3, m_Hdc, ShiftColor(XPFace, -&HA, True)
                                GoSub PaintNormalPost
                                GoTo CutButton
                             End If
                                Exit Sub
                        
                        
                    
                  
                    End If
          
          
          End If
            

Exit Sub

PaintNormalBefore:
                    
                    If m_ColorScheme <> WindowsXP_Silver Then
                    
                    StepXP1 = 25 / RcItem.Bottom
                        For i = 1 To RcItem.Bottom
                            DrawLine 0, i, RcItem.Right, i, m_Hdc, ShiftColor(XPFace, -StepXP1 * i, True)
                        Next i
                    Else
                    DrawGradientMenu m_Hdc, 0, 0, RcItem.Right, RcItem.Bottom, GetRGBColors(vbWhite), GetRGBColors(GetLngColor(&HD7C3C6)), GRADIENT_VERTICAL
                    End If

Return

PaintNormalPost:
                           
                    TempRec.Left = 0: TempRec.Top = 0: TempRec.Right = RcItem.Right: TempRec.Bottom = RcItem.Bottom
                    DrawRectangle TempRec, BColor, m_Hdc
                    Call SetPixelV(m_Hdc, 1, 1, GetLngColor(BColor))
                    Call SetPixelV(m_Hdc, 1, RcItem.Bottom - 2, GetLngColor(BColor))
                    Call SetPixelV(m_Hdc, RcItem.Right - 2, 1, GetLngColor(BColor))
                    Call SetPixelV(m_Hdc, RcItem.Right - 2, RcItem.Bottom - 2, GetLngColor(BColor))

Return




CutButton:
  
  GetClientRect m_hWnd, RcItem
   
      
'//=========================================================================================================
'  MSGBOX SEEM TO HAVE A BUG WHEN CUTTING CORNERS WITH THE MAKE REGION SUB... SO IM GOING TO PAINT THE CORNERS
'//=========================================================================================================
   Call SetPixelV(m_Hdc, 0, 0, GetLngColor(vbButtonFace))
   Call SetPixelV(m_Hdc, 1, 0, GetLngColor(vbButtonFace)) '//--Left Top Cut
   Call SetPixelV(m_Hdc, 0, 1, GetLngColor(vbButtonFace))
   Call SetPixelV(m_Hdc, RcItem.Right - 1, 0, GetLngColor(vbButtonFace))
   Call SetPixelV(m_Hdc, RcItem.Right - 2, 0, GetLngColor(vbButtonFace)) '//--Right Top Cut
   Call SetPixelV(m_Hdc, RcItem.Right - 1, 1, GetLngColor(vbButtonFace))
   Call SetPixelV(m_Hdc, 0, RcItem.Bottom - 2, GetLngColor(vbButtonFace))
   Call SetPixelV(m_Hdc, 1, RcItem.Bottom - 1, GetLngColor(vbButtonFace)) '//--Left Bottom Cut
   Call SetPixelV(m_Hdc, 0, RcItem.Bottom - 1, GetLngColor(vbButtonFace))
   Call SetPixelV(m_Hdc, RcItem.Right - 1, RcItem.Bottom - 2, GetLngColor(vbButtonFace))
   Call SetPixelV(m_Hdc, RcItem.Right - 2, RcItem.Bottom - 1, GetLngColor(vbButtonFace)) '//--Right Bottom Cut
   Call SetPixelV(m_Hdc, RcItem.Right - 1, RcItem.Bottom - 1, GetLngColor(vbButtonFace))
 '//=========================================================================================================
   

'//=========================================================================================================
            '       THIS REALLY CUT'S THE CORNERS FOR BUTTONS EXCEPT IN MSGBOX ..
'//=========================================================================================================
  
   
   Call MakeRegion(RcItem, m_hWnd)
 


End Sub


'//=========================================================================================================
'                                      THIS SUB DRAWS THE BUTTONS CAPTION.
'//=========================================================================================================
 


Private Sub SchemeControl()

    Select Case m_ColorScheme
              
        Case SystemColors, WindowsXP_Blue, WindowsXP_Silver
             BColor = &H733C00
             HCol1 = &H31B2FF: HCol2 = &H96E7&: HCol3 = &HCEF3FF: HCol4 = &H8CDBFF: HCol5 = &H6BCBFF
             FCol1 = &HE7AE8C: FCol2 = &HEF826B: FCol3 = &HFFE7CE: FCol4 = &HF7D7BD: FCol5 = &HF0D1B5
        Case WindowsXP_OliveGreen
             BColor = &H66237
             HCol1 = &H4F91E3: HCol2 = &H2572CF: HCol3 = &H95C5FC: HCol4 = &H96BEED: HCol5 = &H8BB8EB
             FCol1 = &H80CBB1: FCol2 = &H66A7A8: FCol3 = &H8FD1C2: FCol4 = &H80CBB1: FCol5 = &H6DC9A6
      
    End Select
    
End Sub

Public Property Let ColorScheme(ByRef cColorScheme As CWindowColors)
   m_ColorScheme = cColorScheme
End Property

Public Property Let Over(ByVal cOver As Boolean)
   m_Over = cOver
End Property

Public Property Let Down(ByVal cDown As Boolean)
   m_Down = cDown
End Property

Public Property Let Focused(ByVal cFocused As Boolean)
   m_Focused = cFocused
End Property

Public Property Let Enabled(ByVal cEnabled As Boolean)
   m_Enabled = cEnabled
End Property

Public Property Let hwnd(ByVal cHwnd As Long)
   m_hWnd = cHwnd
End Property

Public Property Let hdc(ByVal cHdc As Long)
   m_Hdc = cHdc
End Property

Public Property Let PreDraw(ByVal cPreDraw As Boolean)
   m_PreDraw = cPreDraw
End Property
