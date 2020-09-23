VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSliderXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------------------------'
'--------------------------------------------------------------------------------------------------'
'                                                                                                  '
'                                            cSliderXP.cls                                         '
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


' !!!!! This Version of cSliderXP Doesn't Fully Supports Smaller Sliders...The have some Drawing Problems :(



Option Explicit
Private m_hWnd As Long
Private m_Hdc As Long
Private m_Min As Integer
Private m_Max As Integer
Private m_Val As Integer           '//--- Control Value
Private m_Ort As Byte              '//--- Control Orientation
Private m_Tsl As Byte              '//--- Control TickStyle
Private m_Stt As ControlState      '//--- Control State
Private AreaRgn As RECT
Private TrackRect As RECT


Private Enum cDirection
    Left_TO_Right = 0
    Right_TO_Left = 1
    Top_TO_Bottom = 2
    Bottom_TO_Top = 3
End Enum

Private AB As Long
Private AT As Long   '//---Area Regions ..
Private AL As Long
Private AR As Long

Private TT As Long   '//---Area Regions ..
Private TL As Long
Private TR As Long


'//-- Border Colors
Dim Col1 As Long, Col2 As Long, Col3 As Long, Col4 As Long, Col5 As Long, Col6 As Long
'//-- Inner Colors
Dim ICol1 As Long, ICol2 As Long, ICol3 As Long, ICol4 As Long, ICol5 As Long, ICol6 As Long, ICol7 As Long

'====================================================================================================
'====================================================================================================

'                            DRAW A STYLED XP SLIDER ...NOT EXACTLY BUT CLOSE .. ;)

'====================================================================================================
'====================================================================================================

Public Sub DrawSlider()
    
    SendMessageLong m_hWnd, TBM_GETCHANNELRECT, 0, TrackRect
    SendMessageLong m_hWnd, TBM_GETTHUMBRECT, 0, AreaRgn
    
    AB = AreaRgn.Bottom: AT = AreaRgn.Top: AR = AreaRgn.Right: AL = AreaRgn.Left
    TT = TrackRect.Top: TR = TrackRect.Right: TL = TrackRect.Left
    
    Call DrawSlideLine
    Call DrawSliderMod
    Call DrawCenter
    Call SelectSliderPos
   
   
End Sub

'====================================================================================================
'====================================================================================================

'                                SELECT THE SLIDER ROTATION AND STYLE

'====================================================================================================
'====================================================================================================

Private Sub SelectSliderPos()

Select Case m_Tsl

      
      Case 1
                If m_Ort = 0 Then
                    DrawSliderBody (Bottom_TO_Top)
                Else
                    DrawSliderBody (Right_TO_Left)
                End If
                
      Case 0, 3
                If m_Ort = 0 Then
                    DrawSliderBody (Top_TO_Bottom)
                Else
                    DrawSliderBody (Left_TO_Right)
                End If
      Case 2
                Call DrawSqR
      
      Case Else
           'Nothing
   
   End Select

End Sub

'====================================================================================================
'====================================================================================================

'                                DRAW THE XP STYLED CHANNEL TRACK ....

'            MOST OF THE LINES ARE ""SetPixelV Function"" TO DRAW CORNERS OF THE TRACKS

'====================================================================================================
'====================================================================================================

Private Sub DrawSlideLine()
 
    
    If m_Ort = 0 Then
        
        DrawLine TL, TT, TR, TT, m_Hdc, GetLngColor(&H999C9D)
        DrawLine TL, TT + 1, TR, TT + 1, m_Hdc, GetLngColor(&HE9F1F2)
        DrawLine TL, TT + 2, TR, TT + 2, m_Hdc, GetLngColor(&HE0EDF0)
        DrawLine TL, TT + 3, TR, TT + 3, m_Hdc, GetLngColor(vbWhite)
        
        Call SetPixelV(m_Hdc, TL, TT, GetLngColor(&HD5DCDD))
        Call SetPixelV(m_Hdc, TL + 1, TT, GetLngColor(&HBCC0C0))
        Call SetPixelV(m_Hdc, TL, TT + 1, GetLngColor(&HADB1B2))
        Call SetPixelV(m_Hdc, TL + 1, TT + 1, GetLngColor(&HC7CDCE))
        Call SetPixelV(m_Hdc, TL, TT + 2, GetLngColor(&H999C9C))
        Call SetPixelV(m_Hdc, TL + 1, TT + 2, GetLngColor(&HD5D6D6))
        Call SetPixelV(m_Hdc, TL, TT + 3, GetLngColor(&HCFD3D4))
        Call SetPixelV(m_Hdc, TL + 1, TT + 3, GetLngColor(&HBDC2C3))
         
    Else
    
         'TBM_GETCHANNELRECT Returns Horizontal Rects'..Yust Invert them ;)   TOP=LEFT ,LEFT=TOP,.. ETC...
         DrawLine TT, TL, TT, TR, m_Hdc, GetLngColor(&H999C9D)
         DrawLine TT + 1, TL, TT + 1, TR, m_Hdc, GetLngColor(&HE9F1F2)
         DrawLine TT + 2, TL, TT + 2, TR, m_Hdc, GetLngColor(&HE0EDF0)
         DrawLine TT + 3, TL, TT + 3, TR, m_Hdc, GetLngColor(vbWhite)

         Call SetPixelV(m_Hdc, TT, TL, GetLngColor(&HD5DCDD))
         Call SetPixelV(m_Hdc, TT + 1, TL, GetLngColor(&HBCC0C0))
         Call SetPixelV(m_Hdc, TT, TL + 1, GetLngColor(&HADB1B2))
         Call SetPixelV(m_Hdc, TT + 1, TL + 1, GetLngColor(&HC7CDCE))
         Call SetPixelV(m_Hdc, TT, TR - 1, GetLngColor(&HD5DCDD))
         Call SetPixelV(m_Hdc, TT + 1, TR - 1, GetLngColor(&HBCC0C0))
         Call SetPixelV(m_Hdc, TT, TR - 2, GetLngColor(&HADB1B2))
         Call SetPixelV(m_Hdc, TT + 1, TR - 2, GetLngColor(&HC7CDCE))
       
    End If

        DrawTicks
  
End Sub

'====================================================================================================
'====================================================================================================

'                                   DRAW THE SLIDER STATE MODE ....

'                          (SELECT COLORS DEPENDING ON THE CURRENT STATE OF CONTROL)

'====================================================================================================
'====================================================================================================


Private Sub DrawSliderMod()


    Select Case m_Stt

           Case C_Normal
            
                GoSub DefaultCornerColors
                ICol1 = &H76AC63: ICol2 = &H1FB621: ICol3 = &H18911A: ICol4 = &H59824B: ICol5 = &H64CE66: ICol6 = &H46C447: ICol7 = &H23B925
             
           Case C_Disabled
           
                GoSub DisabledCornerColors
                ICol1 = &HBAC7CA: ICol2 = vbWhite: ICol3 = &HD7E0E1: ICol4 = &HA9B6B9: ICol5 = &HC6D0D2: ICol6 = &HC1CCCE: ICol7 = vbWhite
          
           Case C_Over, C_Focus
            
                GoSub DefaultCornerColors
                ICol1 = &H7EABBA: ICol2 = &H35B4F9: ICol3 = &H2A8EC4: ICol4 = &H60828E: ICol5 = &H74CBFB: ICol6 = &H55C0FA: ICol7 = &H39B5F9
     
           Case C_Down
            
                GoSub DefaultCornerColors
                ICol1 = &H709E63: ICol2 = &H129522: ICol3 = &HE751B: ICol4 = &H55784B: ICol5 = &H5BB666: ICol6 = &H3BA748: ICol7 = &H129522
     
           
           Case Else
           'Nothing
    
    End Select


Exit Sub


DefaultCornerColors:
Col1 = &HCDC4B5: Col2 = &HB2A691: Col3 = &H928877: Col4 = &HC0C3C3: Col5 = &HD9DCDC: Col6 = &H887E6E
Return
DisabledCornerColors:
Col1 = &HC6D4D6: Col2 = &HBAC9CC: Col3 = &HADBCBF: Col4 = &HD9E2E3: Col5 = &HE2EBEC: Col6 = &HA9B8BB
Return

End Sub

'====================================================================================================
'====================================================================================================

'                                   DRAW THE SLIDER BODY ....

'====================================================================================================
'====================================================================================================

Private Sub DrawSliderBody(ByRef Dtn As cDirection)
Dim i As Integer

Dim PMX As Long          '//--Prox VAL
Dim FBM As Long          '//--FixedBottom
Dim DIF As Integer       '//--FLAG TO ROTATE X ANGLE THE SLIDER


Select Case Dtn

        
        
     Case Top_TO_Bottom, Bottom_TO_Top
 
        '=========================================================================================================
        'DRAW THE SLIDER.. 'DIF' IS THE FLAG TO ROTATE THE SLIDER UPSIDE DOWN OR UP DEPENDING ON THE TSL
          
        DIF = IIf(Dtn = Top_TO_Bottom, 0, 2)
        FBM = AB - 1 - ((AR - AL) / 2)
        PMX = Round(((AL + AR) / 2) + 0.1)
         
        DrawLine AL, AT + (DIF * 2.5), AL, FBM + (DIF * 2.5), m_Hdc, GetLngColor(Col1)
        DrawLine AR - 1, AT + 1 + (DIF * 2.5), AR - 1, FBM + (DIF * 2.5), m_Hdc, GetLngColor(Col3)
        DrawLine AR - 2, AT + 1 + (DIF * 2.5), AR - 2, FBM + (DIF * 2.5), m_Hdc, GetLngColor(Col4)
        DrawLine AR - 3, AT + 1 + (DIF * 2.5), AR - 3, FBM + (DIF * 2.5), m_Hdc, GetLngColor(Col5)
        DrawLine AL + 1, AT + (DIF * 10.5), AR - 1, AT + (DIF * 10.5), m_Hdc, GetLngColor(Col1)
        DrawLine AL + 1, AT + 1 + (DIF * 9.5), AR - 1, AT + 1 + (DIF * 9.5), m_Hdc, GetLngColor(ICol5)
        DrawLine AL + 1, AT + 2 + (DIF * 8.5), AR - 1, AT + 2 + (DIF * 8.5), m_Hdc, GetLngColor(ICol6)
        DrawLine AL + 1, AT + 3 + (DIF * 7.5), AR - 1, AT + 3 + (DIF * 7.5), m_Hdc, GetLngColor(ICol7)
        DrawLine PMX, AB - 5 - (DIF * 6.5), AR - 1, FBM - 3 - (DIF * 2.5), m_Hdc, GetLngColor(ICol3)
        DrawLine PMX, AB - 3 - (DIF * 8.5), AR, FBM - 2 - (DIF * 3.5), m_Hdc, GetLngColor(Col6)
        DrawLine PMX, AB - 4 - (DIF * 7.5), AR - 1, FBM - 2 - (DIF * 3.5), m_Hdc, GetLngColor(ICol4)
        DrawLine PMX - 1, AB - 5 - (DIF * 6.5), AR - 1, FBM - 4 - (DIF * 1.5), m_Hdc, GetLngColor(Col4)
        DrawLine PMX - 1, AB - 6 - (DIF * 5.5), AR - 2, FBM - 4 - (DIF * 1.5), m_Hdc, GetLngColor(Col5)
        DrawLine AL, FBM - 1 - (DIF * 4.5), PMX, AB - 1 - (DIF * 10.5), m_Hdc, GetLngColor(Col2)
        DrawLine AL + 1, FBM - 2 - (DIF * 3.5), PMX, AB - 3 - (DIF * 8.5), m_Hdc, GetLngColor(ICol2)
        DrawLine AL + 1, FBM - 1 - (DIF * 4.5), PMX, AB - 2 - (DIF * 9.5), m_Hdc, GetLngColor(ICol1)
        
        
        '====================================================================================
        'CLEAN THUMB AREA ..CORNERS...
   
        If Dtn = Top_TO_Bottom Then
            
            For i = 0 To ((AR - AL) / 2)
                DrawLine AL + i, FBM + i, AL + i, AB, m_Hdc, GetLngColor(vbButtonFace)
                DrawLine AR - i - 1, FBM + i, AR - i - 1, AB, m_Hdc, GetLngColor(vbButtonFace)
            Next i

            SetPixelV m_Hdc, AL, AT, GetLngColor(vbButtonFace)
            SetPixelV m_Hdc, AR - 1, AT, GetLngColor(vbButtonFace)
        
        Else
            
            For i = 0 To ((AR - AL) / 2)
                DrawLine AL + i, AT, AL + i, AT + Round(((AR - AL) / 2) + 0.1) - i, m_Hdc, GetLngColor(vbButtonFace)
                DrawLine AR - i - 1, AT, AR - i - 1, AT + Round(((AR - AL) / 2) + 0.1) - i, m_Hdc, GetLngColor(vbButtonFace)
            Next i
            
            SetPixelV m_Hdc, AL, AB - 1, GetLngColor(vbButtonFace)
            SetPixelV m_Hdc, AR - 1, AB - 1, GetLngColor(vbButtonFace)
                
        End If
      '==================================================================================
   
      
       Case Left_TO_Right, Right_TO_Left
       
       '=========================================================================================================
        'DRAW THE SLIDER.. 'DIF' IS THE FLAG TO ROTATE THE SLIDER LEFTSIDE OR RIGHT DEPENDING ON THE TSL
                  
        DIF = IIf(Dtn = Left_TO_Right, 0, 2)
        FBM = AR - 1 - ((AB - AT) / 2)
        PMX = Round(((AT + AB) / 2) - 0.1)
    
        DrawLine AL + 1 + (DIF * 3), AT, FBM + (DIF * 3) - 1, AT, m_Hdc, GetLngColor(Col1)
        DrawLine AL + (DIF * 10.5), AT + 1, AL + (DIF * 10.5), AB - 1, m_Hdc, GetLngColor(Col2)
        DrawLine AL + 1 + (DIF * 3), AB - 1, FBM + (DIF * 3) - 1, AB - 1, m_Hdc, GetLngColor(Col3)
        DrawLine AL + 1 + (DIF * 3), AB - 2, FBM + (DIF * 3) - 1, AB - 2, m_Hdc, GetLngColor(Col4)
        DrawLine AL + 1 + (DIF * 3), AB - 3, FBM + (DIF * 3) - 1, AB - 3, m_Hdc, GetLngColor(Col5)
        DrawLine FBM - 1 - (DIF * 4.5), AB - 1, AR - 1 - (DIF * 10.5), PMX - 1, m_Hdc, GetLngColor(Col6)
        DrawLine FBM - 1 - (DIF * 4.5), AT, AR - 1 - (DIF * 10.5), PMX + 1, m_Hdc, GetLngColor(Col2)
        DrawLine FBM - 2 - (DIF * 3.5), AB - 4, AR - 5 - (DIF * 6.5), PMX - 1, m_Hdc, GetLngColor(Col5)
        DrawLine FBM - 2 - (DIF * 3.5), AB - 3, AR - 4 - (DIF * 7.5), PMX - 1, m_Hdc, GetLngColor(Col4)
        DrawLine FBM - 1 - (DIF * 4.5), AT + 1, AR - 2 - (DIF * 9.5), PMX + 1, m_Hdc, GetLngColor(ICol1)
        DrawLine FBM - 2 - (DIF * 3.5), AT + 1, AR - 3 - (DIF * 8.5), PMX + 1, m_Hdc, GetLngColor(ICol2)
        DrawLine FBM - 1 - (DIF * 4.5), AB - 2, AR - 3 - (DIF * 8.5), PMX, m_Hdc, GetLngColor(ICol4)
        DrawLine FBM - 2 - (DIF * 3.5), AB - 2, AR - 4 - (DIF * 7.5), PMX, m_Hdc, GetLngColor(ICol3)
        DrawLine AL + 1 + (DIF * 8.5), AT + 1, AL + 1 + (DIF * 8.5), AB - 1, m_Hdc, GetLngColor(ICol5)
        DrawLine AL + 2 + (DIF * 8.5), AT + 1, AL + 2 + (DIF * 8.5), AB - 1, m_Hdc, GetLngColor(ICol6)
        DrawLine AL + 3 + (DIF * 8.5), AT + 1, AL + 3 + (DIF * 8.5), AB - 1, m_Hdc, GetLngColor(ICol7)

        '====================================================================================
        'CLEAN THUMB AREA ..CORNERS...
   
         If Dtn = Left_TO_Right Then
            
            For i = 0 To ((AB - AT) / 2)
                DrawLine FBM + i, AT + i, AR, AT + i, m_Hdc, GetLngColor(vbButtonFace)
                DrawLine AR - i, PMX + i - 1, AR, PMX + i - 1, m_Hdc, GetLngColor(vbButtonFace)
            Next i

            SetPixelV m_Hdc, AL, AT, GetLngColor(vbButtonFace)
            SetPixelV m_Hdc, AL, AB - 1, GetLngColor(vbButtonFace)
        
         Else
            
            For i = 0 To ((AB - AT) / 2)
                DrawLine AL, AT + i, AL + Round(((AB - AT) / 2) + 0.1) - i, AT + i, m_Hdc, GetLngColor(vbButtonFace)
                DrawLine AL, PMX + i - 1, AL + i, PMX + i - 1, m_Hdc, GetLngColor(vbButtonFace)
            Next i
            
            SetPixelV m_Hdc, AR - 1, AT, GetLngColor(vbButtonFace)
            SetPixelV m_Hdc, AR - 1, AB - 1, GetLngColor(vbButtonFace)
                
        End If
         
 
     Case Else
        'Nothing

End Select
     

End Sub

'====================================================================================================
'====================================================================================================

'                                   DRAW THE SLIDER BODY 2....
                                    'IN CASE OF NO THICKS
                                
'====================================================================================================
'====================================================================================================

Private Sub DrawSqR()
  
  If m_Ort = 0 Then
   DrawLine AL, AT, AL, AB, m_Hdc, GetLngColor(Col1)
   DrawLine AR - 1, AT, AR - 1, AB, m_Hdc, GetLngColor(Col3)
   DrawLine AR - 2, AT, AR - 2, AB, m_Hdc, GetLngColor(Col4)
   DrawLine AR - 3, AT, AR - 3, AB, m_Hdc, GetLngColor(Col5)
   
   DrawLine AL, AB - 1, AR, AB - 1, m_Hdc, GetLngColor(Col6)
   DrawLine AL + 1, AT, AR, AT, m_Hdc, GetLngColor(Col1)
   DrawLine AL + 1, AB - 2, AR - 1, AB - 2, m_Hdc, GetLngColor(ICol6)
   DrawLine AL + 1, AB - 3, AR - 1, AB - 3, m_Hdc, GetLngColor(ICol5)
   DrawLine AL + 1, AT + 1, AR - 1, AT + 1, m_Hdc, GetLngColor(ICol5)
   DrawLine AL + 1, AT + 2, AR - 1, AT + 2, m_Hdc, GetLngColor(ICol6)
   
  Else

    DrawLine AL, AT, AR, AT, m_Hdc, GetLngColor(Col1)
    DrawLine AL, AB - 1, AR, AB - 1, m_Hdc, GetLngColor(Col3)
    DrawLine AL, AB - 2, AR, AB - 2, m_Hdc, GetLngColor(Col4)
    DrawLine AL, AB - 3, AR, AB - 3, m_Hdc, GetLngColor(Col5)
    DrawLine AR - 1, AT, AR - 1, AB, m_Hdc, GetLngColor(Col6)
    DrawLine AL, AT, AL, AB, m_Hdc, GetLngColor(Col1)
    DrawLine AR - 2, AT + 1, AR - 2, AB - 1, m_Hdc, GetLngColor(ICol6)
    DrawLine AR - 3, AT + 1, AR - 3, AB - 1, m_Hdc, GetLngColor(ICol5)
    DrawLine AL + 1, AT + 1, AL + 1, AB - 1, m_Hdc, GetLngColor(ICol5)
    DrawLine AL + 2, AT + 1, AL + 2, AB - 1, m_Hdc, GetLngColor(ICol6)
   
   
  End If
   
   SetPixelV m_Hdc, AR - 1, AT, GetLngColor(vbButtonFace)
   SetPixelV m_Hdc, AL, AT, GetLngColor(vbButtonFace)
   SetPixelV m_Hdc, AR - 1, AB - 1, GetLngColor(vbButtonFace)
   SetPixelV m_Hdc, AL, AB - 1, GetLngColor(vbButtonFace)
  
End Sub

'====================================================================================================
'====================================================================================================

'                                   DRAW THE SLIDER CENTER
                                    '     WHITE AREA
                                
'====================================================================================================
'====================================================================================================

Private Sub DrawCenter()

Dim hBrush As Long
Dim hRect As RECT
         
    hRect.Left = AL
    hRect.Right = AR
    hRect.Top = AT
    hRect.Bottom = AB
    
    hBrush = CreateSolidBrush(vbWhite)
    FillRect m_Hdc, hRect, hBrush
    DeleteObject hBrush

End Sub

'====================================================================================================
'====================================================================================================

'                                   DRAW THE SLIDER TICKS
                    '     CHANGE THICKS COLOR FROM BLACK TO XP COLOR (&H92A1A1)
                                
'====================================================================================================
'====================================================================================================

Private Sub DrawTicks()
Dim i As Integer
Dim TicPos As Integer
Dim NumTics As Integer
 
 
 NumTics = SendMessageLong(m_hWnd, TBM_GETNUMTICS, 0&, 0&)
   
 If NumTics = 0 Then Exit Sub
 
 If m_Ort = 0 Then
   '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
      For i = 0 To NumTics - 3
            TicPos = SendMessageLong(m_hWnd, TBM_GETTICPOS, i, 0&)
            If m_Tsl = 0 Or m_Tsl = 2 Then DrawLine TicPos, AB + 1, TicPos, AB + 4, m_Hdc, GetLngColor(&H92A1A1)
            If m_Tsl = 1 Or m_Tsl = 2 Then DrawLine TicPos, AT - 4, TicPos, AT - 1, m_Hdc, GetLngColor(&H92A1A1)
      Next i

      If m_Tsl = 0 Or m_Tsl = 2 Then
            DrawLine TL + 5, AB + 1, TL + 5, AB + 5, m_Hdc, GetLngColor(&H92A1A1) 'LEFT LOGICAL
            DrawLine TR - 6, AB + 1, TR - 6, AB + 5, m_Hdc, GetLngColor(&H92A1A1) 'RIGHT LOGICAL
      End If

      If m_Tsl = 1 Or m_Tsl = 2 Then
            DrawLine TL + 5, AT - 5, TL + 5, AT - 1, m_Hdc, GetLngColor(&H92A1A1) 'LEFT LOGICAL
            DrawLine TR - 6, AT - 5, TR - 6, AT - 1, m_Hdc, GetLngColor(&H92A1A1) 'RIGHT LOGICAL
      End If
   '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
 
 Else
   '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
      For i = 0 To NumTics - 3
            TicPos = SendMessageLong(m_hWnd, TBM_GETTICPOS, i, 0&)
            If m_Tsl = 0 Or m_Tsl = 2 Then DrawLine AR + 1, TicPos, AR + 4, TicPos, m_Hdc, GetLngColor(&H92A1A1)
            If m_Tsl = 1 Or m_Tsl = 2 Then DrawLine AL - 4, TicPos, AL - 1, TicPos, m_Hdc, GetLngColor(&H92A1A1)
      Next i

      If m_Tsl = 0 Or m_Tsl = 2 Then
             DrawLine AR + 1, TL + 5, AR + 5, TL + 5, m_Hdc, GetLngColor(&H92A1A1) 'LEFT LOGICAL
             DrawLine AR + 1, TR - 6, AR + 5, TR - 6, m_Hdc, GetLngColor(&H92A1A1) 'RIGHT LOGICAL
      End If

      If m_Tsl = 1 Or m_Tsl = 2 Then
             DrawLine AL - 5, TL + 5, AL - 1, TL + 5, m_Hdc, GetLngColor(&H92A1A1) 'LEFT LOGICAL
             DrawLine AL - 5, TR - 6, AL - 1, TR - 6, m_Hdc, GetLngColor(&H92A1A1) 'RIGHT LOGICAL
      End If
   '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
 
 
 End If


End Sub

Public Property Let hwnd(ByVal cHwnd As Long)
   m_hWnd = cHwnd
End Property

Public Property Let hdc(ByVal cHdc As Long)
   m_Hdc = cHdc
End Property

Public Property Let Min(ByVal cMin As Integer)
   m_Min = cMin
End Property

Public Property Let Max(ByVal cMax As Integer)
   m_Max = cMax
End Property

Public Property Let Orientation(ByVal cOrientation As Byte)
   m_Ort = cOrientation
End Property

Public Property Let Value(ByVal cValue As Integer)
   m_Val = cValue
End Property

Public Property Let TickStyle(ByVal cTickStyle As Byte)
   m_Tsl = cTickStyle
End Property

Public Property Let State(ByVal cState As ControlState)
   m_Stt = cState
End Property
