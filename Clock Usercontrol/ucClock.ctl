VERSION 5.00
Begin VB.UserControl ucClock 
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   1260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1260
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   DrawStyle       =   2  'Dot
   HasDC           =   0   'False
   HitBehavior     =   0  'None
   MaskColor       =   &H80000014&
   PaletteMode     =   4  'None
   PropertyPages   =   "ucClock.ctx":0000
   ScaleHeight     =   84
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   84
   ToolboxBitmap   =   "ucClock.ctx":0024
   Windowless      =   -1  'True
   Begin VB.Timer tmrSystemClockTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   45
      Top             =   780
   End
End
Attribute VB_Name = "ucClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'---------------------------------------------------------------------------------------------------------------------------------
' MODULE     : ucClock
' FILENAME   : C:\Documents and Settings\Bryan\Desktop\Clock Test Project\UserControl1.ctl
' CREATED BY : Bryan Utley
'         ON : Thursday, March 15, 2007 at 3:10:28 PM
' COPYRIGHT  : Copyright 2007 - All Rights Reserved
'              The World Wide Web Programmer's Consortium
'
' DESCRIPTION: A usercontrol for displaying a photo-realistic clock.
'
' Credits/Acknowledgements - Thanks goes to:
'
'   LaVolpe for his generous contributions to all of us here at PSC.  This entire project
'           is based on his c32bppDIB class.  I have used his set of classes here without
'           modification, but for any updates that he may have made to this invalueable PSC
'           submission, please visit the following link:
'
'       http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=67466&lngWId=1
'
' COMMENTS:
'
' WEB SITE   : http://www.thewwwpc.com
' E-MAIL     : bryan@thewwwpc.com or bryanutley2000@yahoo.com
'
' MODIFICATION HISTORY:
'
' 1.0.0   MODIFIED ON   : Thursday, March 15, 2007 at 3:10:28 PM
'         MODIFIED BY   : Bryan Utley
'         MODIFICATIONS : Initial Version
'         ASSISTANCE    : Considerable help and guidence from LaVolpe (Keith)
' ---------------------------------------------------------------------------------------------------------------------------------
'
Option Explicit
Option Base 1

Private Declare Function BitBlt Lib "gdi32.dll" ( _
        ByVal hDestDC As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hSrcDC As Long, _
        ByVal xSrc As Long, _
        ByVal ySrc As Long, _
        ByVal dwRop As Long) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" ( _
        ByVal hdc As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32.dll" ( _
        ByVal hdc As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32.dll" ( _
        ByVal crColor As Long) As Long

Private Declare Function DeleteDC Lib "gdi32.dll" ( _
        ByVal hdc As Long) As Long

Private Declare Function DeleteObject Lib "gdi32.dll" ( _
        ByVal hObject As Long) As Long

Private Declare Function FillRect Lib "user32.dll" ( _
        ByVal hdc As Long, _
        ByRef lpRect As Any, _
        ByVal hBrush As Long) As Long

Private Declare Function GetClipBox Lib "gdi32.dll" ( _
        ByVal hdc As Long, _
        ByRef lpRect As RECT) As Long

Private Declare Function GetDC Lib "user32.dll" ( _
        ByVal hwnd As Long) As Long

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" ( _
        ByVal lpszSoundName As String, _
        ByVal uFlags As Long) As Long

Private Declare Function ReleaseDC Lib "user32.dll" ( _
        ByVal hwnd As Long, _
        ByVal hdc As Long) As Long

Private Declare Function SelectObject Lib "gdi32.dll" ( _
        ByVal hdc As Long, _
        ByVal hObject As Long) As Long

Private Declare Function SetRect Lib "user32" ( _
        lpRect As Any, _
        ByVal X1 As Long, _
        ByVal Y1 As Long, _
        ByVal X2 As Long, _
        ByVal Y2 As Long) As Long

Enum enumStyles
    [Style 1] = 1
    [Style 2] = 2
    [Style 3] = 3
    [Style 4] = 4
    [Style 5] = 5
    [Style 6] = 6
    [Style 7] = 7
    [Style 8] = 8
End Enum

Enum enumOpacityLevel
    [Invisible] = 0
    [5 Percent] = 5
    [10 Percent] = 10
    [15 Percent] = 15
    [20 Percent] = 20
    [25 Percent] = 25
    [30 Percent] = 30
    [35 Percent] = 35
    [40 Percent] = 40
    [45 Percent] = 45
    [50 Percent] = 50
    [55 Percent] = 55
    [60 Percent] = 60
    [65 Percent] = 65
    [70 Percent] = 70
    [75 Percent] = 75
    [80 Percent] = 80
    [85 Percent] = 85
    [90 Percent] = 90
    [95 Percent] = 95
    [Opaque] = 100
End Enum

Private Type RECT
    Left        As Long
    Top         As Long
    Right       As Long
    Bottom      As Long
End Type

Private Type LogoImage
    Left        As Long
    Top         As Long
    Width       As Long
    Height      As Long
End Type

Private Type OffscreenDC
    DC          As Long
    hBmp        As Long
    hOldBmp     As Long
End Type

Private Type ColorSetting
    Hue         As Single
    Saturation  As Single
    Luminosity  As Single
End Type

Private Type AlarmTypes
    Time        As String
    Reason      As String
End Type

Private Const SND_ASYNC       As Long = &H1
Private Const OneSecond       As Long = 1051
Private Const FifteenSeconds  As Long = &H3A97

Private logo                  As LogoImage
Private memDC                 As OffscreenDC
Private cAlarms               As New Collection

Private oBackground           As New c32bppDIB
Private oHand_Hour            As New c32bppDIB
Private oHand_Minute          As New c32bppDIB
Private oHand_Second          As New c32bppDIB
Private oHighlights           As New c32bppDIB
Private oLogoImage            As c32bppDIB

Private mClockSize            As Integer
Private mColor_Highlight      As ColorSetting
Private mHighlightRotation    As Integer
Private mLogoImage            As LogoImage
Private mOpacity_Background   As Integer
Private mOpacity_Hand_Hour    As Integer
Private mOpacity_Hand_Minute  As Integer
Private mOpacity_Hand_Second  As Integer
Private mOpacity_Highlights   As Integer
Private mOpacity_LogoImage    As Integer
Private mShowHighLights       As Boolean
Private mShowLogoImage        As Boolean
Private mShowSecondHand       As Boolean
Private mSoundTickStyle       As String
Private mSoundPlayTick        As Boolean
Private mStyle                As Integer
Private mTimeZoneOffset       As Integer
Private mHands_Offset(3, 2)   As Integer

Public Event Alarm(ByVal sHour As String, ByVal sMinute As String, ByVal sSecond As String, ByVal sComment As String, ByVal sKey As String)
Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


Public Function AlarmAdd(ByVal iHour As Integer, ByVal iMinute As Integer, ByVal iSecond As Integer, Optional ByVal sKey As String = "", Optional ByVal sComment As String = "") As Integer

    cAlarms.Add CStr(iHour) & "|" & CStr(iMinute) & "|" & CStr(iSecond) & "|" & sComment, sKey
    AlarmAdd = cAlarms.Count

End Function

Public Function AlarmDelete(ByVal vIndex As Variant) As Boolean

    On Error GoTo err_Handler

    cAlarms.Remove vIndex
    AlarmDelete = True

    Exit Function

err_Handler:

    AlarmDelete = False

End Function

Public Function AlarmRemoveAll() As Boolean

    On Error GoTo err_Handler

    While cAlarms.Count > 0
        cAlarms.Remove 1
    Wend

    AlarmRemoveAll = True

    Exit Function

err_Handler:

    AlarmRemoveAll = False

End Function

Private Function CalculatedAngle(ByVal AngleToCalculate As enumClockHands) As Long

    Select Case AngleToCalculate

     Case eHour
        CalculatedAngle = 30& * (Hour(Now) + TimeZoneOffset + (Minute(Now) / 60))

     Case eMinute
        CalculatedAngle = Minute(Now) * 6

     Case eSecond
        CalculatedAngle = Second(Now) * 6

    End Select

End Function

Private Sub CheckAlarms(ByVal iHour As Integer, ByVal iMinute As Integer, ByVal iSecond As Integer)

  Dim arrAlarm As Variant
  Dim iIndex   As Integer

    For iIndex = 1 To cAlarms.Count

        arrAlarm = Split(cAlarms.Item(iIndex), "|", , vbTextCompare)

        If UBound(arrAlarm) = 3 Then

            If CInt(arrAlarm(0)) = iHour And CInt(arrAlarm(1)) = iMinute Then

                RaiseEvent Alarm(arrAlarm(0), arrAlarm(1), arrAlarm(2), arrAlarm(3), iIndex)

            End If

        End If

    Next

End Sub

Public Property Get ClockSize() As Integer

    ClockSize = mClockSize

End Property

Public Property Let ClockSize(ByVal Value As Integer)

    mClockSize = Value
    PropertyChanged "ClockSize"
    UserControl_Resize

End Property

Public Property Let Color_Highlight_Hue(ByVal Value As Single)

    mColor_Highlight.Hue = Value
    PropertyChanged "Color_Highlight_Hue"
    ResetHighlightColor

End Property

Public Property Get Color_Highlight_Hue() As Single

    Color_Highlight_Hue = mColor_Highlight.Hue

End Property

Public Property Get Color_Highlight_Luminosity() As Single

    Color_Highlight_Luminosity = mColor_Highlight.Luminosity

End Property

Public Property Let Color_Highlight_Luminosity(ByVal Value As Single)

    mColor_Highlight.Luminosity = Value
    PropertyChanged "Color_Highlight_Luminosity"
    ResetHighlightColor

End Property

Public Property Let Color_Highlight_Saturation(ByVal Value As Single)

    mColor_Highlight.Saturation = Value
    PropertyChanged "Color_Highlight_Saturation"
    ResetHighlightColor

End Property

Public Property Get Color_Highlight_Saturation() As Single

    Color_Highlight_Saturation = mColor_Highlight.Saturation

End Property

Public Property Let HandOffset_Hour_Horz(ByVal Value As Integer)

    mHands_Offset(eHour, eHorz) = Value
    PropertyChanged "HandOffset_Hour_Horz"
    SetImages
    SetLogoMetrics
    UserControl.Refresh

End Property

Public Property Get HandOffset_Hour_Horz() As Integer

    HandOffset_Hour_Horz = mHands_Offset(eHour, eHorz)

End Property

Public Property Let HandOffset_Hour_Vert(ByVal Value As Integer)

    mHands_Offset(eHour, eVert) = Value
    PropertyChanged "HandOffset_Hour_Vert"
    SetImages
    SetLogoMetrics
    UserControl.Refresh

End Property

Public Property Get HandOffset_Hour_Vert() As Integer

    HandOffset_Hour_Vert = mHands_Offset(eHour, eVert)

End Property

Public Property Get HandOffset_Minute_Horz() As Integer

    HandOffset_Minute_Horz = mHands_Offset(eMinute, eHorz)

End Property

Public Property Let HandOffset_Minute_Horz(ByVal Value As Integer)

    mHands_Offset(eMinute, eHorz) = Value
    PropertyChanged "HandOffset_Minute_Horz"
    SetImages
    SetLogoMetrics
    UserControl.Refresh

End Property

Public Property Get HandOffset_Minute_Vert() As Integer

    HandOffset_Minute_Vert = mHands_Offset(eMinute, eVert)

End Property

Public Property Let HandOffset_Minute_Vert(ByVal Value As Integer)

    mHands_Offset(eMinute, eVert) = Value
    PropertyChanged "HandOffset_Minute_Vert"
    SetImages
    SetLogoMetrics
    UserControl.Refresh

End Property

Public Property Let HandOffset_Second_Horz(ByVal Value As Integer)

    mHands_Offset(eSecond, eHorz) = Value
    PropertyChanged "HandOffset_Second_Horz"
    SetImages
    SetLogoMetrics
    UserControl.Refresh

End Property

Public Property Get HandOffset_Second_Horz() As Integer

    HandOffset_Second_Horz = mHands_Offset(eSecond, eHorz)

End Property

Public Property Get HandOffset_Second_Vert() As Integer

    HandOffset_Second_Vert = mHands_Offset(eSecond, eVert)

End Property

Public Property Let HandOffset_Second_Vert(ByVal Value As Integer)

    mHands_Offset(eSecond, eVert) = Value
    PropertyChanged "HandOffset_Second_Vert"
    SetImages
    SetLogoMetrics
    UserControl.Refresh

End Property

Private Function Hands_Offset(oHand As enumClockHands, oOffset As enumHVoffsets) As Integer

    Hands_Offset = mHands_Offset(oHand, oOffset)

End Function

Public Property Get HighlightRotation() As Integer

    HighlightRotation = mHighlightRotation

End Property

Public Property Let HighlightRotation(ByVal Value As Integer)

    mHighlightRotation = Value
    PropertyChanged "HighlightRotation"
    UserControl.Refresh

End Property

Public Property Let LogoImage_Height(ByVal Value As Long)

    mLogoImage.Height = Value
    PropertyChanged "LogoImage_Height"
    SetLogoMetrics
    UserControl.Refresh

End Property

Public Property Get LogoImage_Height() As Long

    LogoImage_Height = mLogoImage.Height

End Property

Public Property Get LogoImage_Left() As Long

    LogoImage_Left = mLogoImage.Left

End Property

Public Property Let LogoImage_Left(ByVal Value As Long)

    mLogoImage.Left = Value
    PropertyChanged "LogoImage_Left"
    SetLogoMetrics
    UserControl.Refresh

End Property

Public Property Let LogoImage_Top(ByVal Value As Long)

    mLogoImage.Top = Value
    PropertyChanged "LogoImage_Top"
    SetLogoMetrics
    UserControl.Refresh

End Property

Public Property Get LogoImage_Top() As Long

    LogoImage_Top = mLogoImage.Top

End Property

Public Property Get LogoImage_Width() As Long

    LogoImage_Width = mLogoImage.Width

End Property

Public Property Let LogoImage_Width(ByVal Value As Long)

    mLogoImage.Width = Value
    PropertyChanged "LogoImage_Width"
    SetLogoMetrics
    UserControl.Refresh

End Property

Public Property Let Opacity_Background(ByVal vOpacity_Background As enumOpacityLevel)

    mOpacity_Background = vOpacity_Background
    PropertyChanged "Opacity_Background"
    SetLogoMetrics
    UserControl.Refresh

End Property

Public Property Get Opacity_Background() As enumOpacityLevel

    Opacity_Background = mOpacity_Background

End Property

Public Property Let Opacity_Hand_Hour(ByVal vOpacity_Hand_Hour As enumOpacityLevel)

    mOpacity_Hand_Hour = vOpacity_Hand_Hour
    PropertyChanged "Opacity_Hand_Hour"
    SetLogoMetrics
    UserControl.Refresh

End Property

Public Property Get Opacity_Hand_Hour() As enumOpacityLevel

    Opacity_Hand_Hour = mOpacity_Hand_Hour

End Property

Public Property Let Opacity_Hand_Minute(ByVal vOpacity_Hand_Minute As enumOpacityLevel)

    mOpacity_Hand_Minute = vOpacity_Hand_Minute
    PropertyChanged "Opacity_Hand_Minute"
    UserControl.Refresh

End Property

Public Property Get Opacity_Hand_Minute() As enumOpacityLevel

    Opacity_Hand_Minute = mOpacity_Hand_Minute

End Property

Public Property Let Opacity_Hand_Second(ByVal vOpacity_Hand_Second As enumOpacityLevel)

    mOpacity_Hand_Second = vOpacity_Hand_Second
    PropertyChanged "Opacity_Hand_Second"
    UserControl.Refresh

End Property

Public Property Get Opacity_Hand_Second() As enumOpacityLevel

    Opacity_Hand_Second = mOpacity_Hand_Second

End Property

Public Property Let Opacity_Highlights(ByVal vOpacity_Highlights As enumOpacityLevel)

    mOpacity_Highlights = vOpacity_Highlights
    PropertyChanged "Opacity_Highlights"
    UserControl.Refresh

End Property

Public Property Get Opacity_Highlights() As enumOpacityLevel

    Opacity_Highlights = mOpacity_Highlights

End Property

Public Property Get Opacity_LogoImage() As enumOpacityLevel

    Opacity_LogoImage = mOpacity_LogoImage

End Property

Public Property Let Opacity_LogoImage(ByVal vOpacity_LogoImage As enumOpacityLevel)

    mOpacity_LogoImage = vOpacity_LogoImage
    PropertyChanged "Opacity_LogoImage"
    SetLogoMetrics
    UserControl.Refresh

End Property

Private Sub PaintIT(ByVal hdc As Long, X As Long, Y As Long, cX As Long, cY As Long)

  Dim iAngle_Hour   As Integer
  Dim iAngle_Minute As Integer
  Dim iAngle_Second As Integer
  Dim lBrush        As Long
  Dim uTileRect     As RECT

    '// Blt copy oBackground to DC
    BitBlt memDC.DC, X, Y, cX, cY, hdc, X, Y, vbSrcCopy

    '// Render Background to DC
    oBackground.Render memDC.DC, X, Y, ClockSize, ClockSize, X, Y, , , Opacity_Background

    '// Rotate and Render Hour Hand to DC
    If (CalculatedAngle(eMinute) Mod 90 = 0) And (CalculatedAngle(eSecond) = 0) Then
        '// Render Logo to DC
        If ShowLogoImage Then
            SetLogoMetrics
        End If
        oBackground.LoadPicture_FromOrignalFormat
        oHand_Hour.RotateAtTopLeft oBackground.LoadDIBinDC(True), CalculatedAngle(eHour), Hands_Offset(eHour, eHorz), Hands_Offset(eHour, eVert), , , , , , , Opacity_Hand_Hour
        oBackground.LoadDIBinDC False
    End If

    '// Rotate and Render Minute Hand to DC
    oHand_Minute.RotateAtTopLeft memDC.DC, CalculatedAngle(eMinute), Hands_Offset(eMinute, eHorz), Hands_Offset(eMinute, eVert), ClockSize, ClockSize, , , , , Opacity_Hand_Minute

    '// Rotate and Render Secnd Hand to DC
    If ShowSecondHand Then
        oHand_Second.RotateAtTopLeft memDC.DC, CalculatedAngle(eSecond), Hands_Offset(eSecond, eHorz), Hands_Offset(eSecond, eVert), ClockSize, ClockSize, , , , , Opacity_Hand_Second
    End If

    '// Render Clock Highlights overlay to DC
    If ShowHighlights Then
        oHighlights.RotateAtTopLeft memDC.DC, HighlightRotation, X, Y, ClockSize, ClockSize, X, Y, , , Opacity_Highlights
    End If

    '// Render complete clock Image DC to Usercontrol DC
    BitBlt UserControl.hdc, X, Y, cX, cY, memDC.DC, X, Y, vbSrcCopy

End Sub

Private Sub ResetHighlightColor()

    oHighlights.LoadPicture_FromOrignalFormat
    oHighlights.Colorize mColor_Highlight.Hue, mColor_Highlight.Saturation, mColor_Highlight.Luminosity
    UserControl.Refresh

End Sub

Private Sub SetImages()

  Dim StyleFolder As String

    StyleFolder = "Style" & CStr(Style)

    oBackground.LoadPicture_File App.Path & "\Images\" & StyleFolder & "\system.png", ClockSize, ClockSize, True
    oHand_Hour.LoadPicture_File App.Path & "\Images\" & StyleFolder & "\system_h.png", ClockSize, ClockSize
    oHand_Minute.LoadPicture_File App.Path & "\Images\" & StyleFolder & "\system_m.png", ClockSize, ClockSize
    oHand_Second.LoadPicture_File App.Path & "\Images\" & StyleFolder & "\system_s.png", ClockSize, ClockSize
    oHighlights.LoadPicture_File App.Path & "\Images\" & StyleFolder & "\System_Highlights.png", ClockSize, ClockSize, True

    If ShowLogoImage Then
        SetLogoMetrics

     Else
        oHand_Hour.RotateAtTopLeft oBackground.LoadDIBinDC(True), CalculatedAngle(eHour), Hands_Offset(eHour, eHorz), Hands_Offset(eHour, eVert), , , , , , , Opacity_Hand_Hour
        oBackground.LoadDIBinDC False

    End If

    UserControl.Refresh

End Sub

Public Sub SetLogoMetrics()

    oBackground.LoadPicture_FromOrignalFormat

    If mShowLogoImage Then

        If oLogoImage Is Nothing Then

            Set oLogoImage = New c32bppDIB
            oLogoImage.LoadPicture_File App.Path & "\Images\" & "Style" & CStr(Style) & "\system_Logo.png"

        End If

        logo.Width = IIf(LogoImage_Width = -1, oLogoImage.Width, LogoImage_Width)
        logo.Height = IIf(LogoImage_Height = -1, oLogoImage.Height, LogoImage_Height)
        logo.Top = IIf(LogoImage_Top = -1, oBackground.Width / 2, LogoImage_Top)
        logo.Left = IIf(LogoImage_Left = -1, (oBackground.Height / 2) - (logo.Width / 2), LogoImage_Left)
        oLogoImage.Render oBackground.LoadDIBinDC(True), logo.Left, logo.Top, logo.Width, logo.Height, , , , , Opacity_LogoImage

     Else

        Set oLogoImage = Nothing

    End If

    oHand_Hour.RotateAtTopLeft oBackground.LoadDIBinDC(True), CalculatedAngle(eHour), Hands_Offset(eHour, eHorz), Hands_Offset(eHour, eVert), , , , , , , Opacity_Hand_Hour
    oBackground.LoadDIBinDC False

End Sub

Public Property Let ShowHighlights(ByVal vShowHighlights As Boolean)

    mShowHighLights = vShowHighlights
    PropertyChanged "ShowHighlights"
    UserControl.Refresh

End Property

Public Property Get ShowHighlights() As Boolean

    ShowHighlights = mShowHighLights

End Property

Public Property Get ShowLogoImage() As Boolean

    ShowLogoImage = mShowLogoImage

End Property

Public Property Let ShowLogoImage(ByVal vShowLogoImage As Boolean)

    mShowLogoImage = vShowLogoImage
    PropertyChanged "ShowLogoImage"
    SetLogoMetrics
    UserControl.Refresh

End Property

Public Property Get ShowSecondHand() As Boolean

    ShowSecondHand = mShowSecondHand

End Property

Public Property Let ShowSecondHand(ByVal vShowSecondHand As Boolean)

    If UserControl.Ambient.UserMode = True Then

        With tmrSystemClockTimer
            .Enabled = False

            If vShowSecondHand Then
                .Interval = OneSecond
             Else
                .Interval = FifteenSeconds
            End If

            .Enabled = True
        End With

    End If

    mShowSecondHand = vShowSecondHand
    PropertyChanged "ShowSecondHand"
    UserControl.Refresh

End Property

Public Property Get SoundPlayTick() As Boolean

    SoundPlayTick = mSoundPlayTick

End Property

Public Property Let SoundPlayTick(ByVal bvalue As Boolean)

    mSoundPlayTick = bvalue
    PropertyChanged "SoundPlayTick"

End Property

Public Property Get SoundTickStyle() As String

    SoundTickStyle = mSoundTickStyle

End Property

Public Property Let SoundTickStyle(ByVal sTickFileName As String)

    mSoundTickStyle = sTickFileName
    PropertyChanged "SoundTickStyle"

End Property

Public Property Let Style(ByVal vStyle As enumStyles)

    If vStyle >= [Style 1] And vStyle <= [Style 8] Then
        mStyle = vStyle
        PropertyChanged "Style"
        SetImages
        UserControl_Resize
        UserControl.Refresh
    End If

End Property

Public Property Get Style() As enumStyles

    Style = mStyle

End Property

Public Property Get TimeZoneOffset() As Integer

    TimeZoneOffset = mTimeZoneOffset

End Property

Public Property Let TimeZoneOffset(ByVal Value As Integer)

    mTimeZoneOffset = Value
    PropertyChanged "TimeZoneOffset"
    SetLogoMetrics
    UserControl.Refresh

End Property

Private Sub tmrSystemClockTimer_Timer()

  Dim iCurrentSecond As Integer

    iCurrentSecond = Second(Now)

    '// PlaySound if enabled
    If SoundPlayTick And ShowSecondHand Then
        sndPlaySound App.Path & "\sounds\ticks\" & SoundTickStyle & ".wav", SND_ASYNC
    End If

    UserControl.Refresh

    If mShowSecondHand = False Then
        If iCurrentSecond < 45 Then
            tmrSystemClockTimer.Interval = FifteenSeconds
         Else
            tmrSystemClockTimer.Interval = (60 - iCurrentSecond) * 1000
        End If

    End If

    If iCurrentSecond = 0 And cAlarms.Count > 0 Then
        CheckAlarms Hour(Now), Minute(Now), Second(Now)
    End If

End Sub

Private Sub UserControl_Click()

    RaiseEvent Click

End Sub

Private Sub UserControl_DblClick()

    RaiseEvent DblClick

End Sub

Private Sub UserControl_Hide()

    tmrSystemClockTimer.Enabled = False

End Sub

Private Sub UserControl_HitTest(X As Single, Y As Single, HitResult As Integer)

    HitResult = vbHitResultHit

End Sub

Private Sub UserControl_Initialize()

    UserControl_Resize

End Sub

Private Sub UserControl_InitProperties()

    ShowSecondHand = True
    ShowLogoImage = False
    ShowHighlights = True
    Opacity_Background = Opaque
    Opacity_Hand_Hour = Opaque
    Opacity_Hand_Minute = Opaque
    Opacity_Hand_Second = Opaque
    Opacity_Highlights = Opaque
    Opacity_LogoImage = Opaque
    LogoImage_Height = -1
    LogoImage_Left = -1
    LogoImage_Top = -1
    LogoImage_Width = -1
    TimeZoneOffset = 0
    SoundPlayTick = False
    SoundTickStyle = "Tick-1"
    HighlightRotation = 0
    Color_Highlight_Hue = 0
    Color_Highlight_Luminosity = 1
    Color_Highlight_Saturation = 0.5

    ClockSize = 130
    Style = 1

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Sub UserControl_Paint()

  Dim uRect As RECT

    GetClipBox UserControl.hdc, uRect
    PaintIT UserControl.hdc, uRect.Left, uRect.Top, uRect.Right - uRect.Left, uRect.Bottom - uRect.Top

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Style = PropBag.ReadProperty("Style", 1)
    ShowSecondHand = PropBag.ReadProperty("ShowSecondHand", True)
    ShowLogoImage = PropBag.ReadProperty("ShowLogoImage", False)
    ShowHighlights = PropBag.ReadProperty("ShowHighlights", True)
    Opacity_Background = PropBag.ReadProperty("Opacity_Background", 100)
    Opacity_Hand_Hour = PropBag.ReadProperty("Opacity_Hand_Hour", 100)
    Opacity_Hand_Minute = PropBag.ReadProperty("Opacity_Hand_Minute", 100)
    Opacity_Hand_Second = PropBag.ReadProperty("Opacity_Hand_Second", 100)
    Opacity_Highlights = PropBag.ReadProperty("Opacity_Highlights", 100)
    Opacity_LogoImage = PropBag.ReadProperty("Opacity_LogoImage", 100)
    LogoImage_Height = PropBag.ReadProperty("LogoImage_Height", -1)
    LogoImage_Width = PropBag.ReadProperty("LogoImage_Width", -1)
    LogoImage_Top = PropBag.ReadProperty("LogoImage_Top", -1)
    LogoImage_Left = PropBag.ReadProperty("LogoImage_Left", -1)
    TimeZoneOffset = PropBag.ReadProperty("TimeZoneOffset", 0)
    SoundPlayTick = PropBag.ReadProperty("SoundPlayTick", False)
    SoundTickStyle = PropBag.ReadProperty("SoundTickStyle", "Tick-1")
    ClockSize = PropBag.ReadProperty("ClockSize", 130)
    HighlightRotation = PropBag.ReadProperty("HighlightRotation", 0)
    Color_Highlight_Hue = PropBag.ReadProperty("Color_Highlight_Hue", 0)
    Color_Highlight_Saturation = PropBag.ReadProperty("Color_Highlight_Saturation", 0.5)
    Color_Highlight_Luminosity = PropBag.ReadProperty("Color_Highlight_Luminosity", 1)
    HandOffset_Hour_Horz = PropBag.ReadProperty("HandOffset_Hour_Horz", 0)
    HandOffset_Hour_Vert = PropBag.ReadProperty("HandOffset_Hour_Vert", 0)
    HandOffset_Minute_Horz = PropBag.ReadProperty("HandOffset_Minute_Horz", 0)
    HandOffset_Minute_Vert = PropBag.ReadProperty("HandOffset_Minute_Vert", 0)
    HandOffset_Second_Horz = PropBag.ReadProperty("HandOffset_Second_Horz", 0)
    HandOffset_Second_Vert = PropBag.ReadProperty("HandOffset_Second_Vert", 0)

    SetImages

End Sub

Private Sub UserControl_Resize()

  Dim dDC As Long

    UserControl.Size ClockSize * 15, ClockSize * 15

    dDC = GetDC(0&)

    If memDC.DC = 0& Then
        memDC.DC = CreateCompatibleDC(dDC)

     Else
        DeleteObject SelectObject(memDC.DC, memDC.hOldBmp)

    End If

    memDC.hBmp = CreateCompatibleBitmap(dDC, ClockSize, ClockSize)
    ReleaseDC 0&, dDC
    memDC.hOldBmp = SelectObject(memDC.DC, memDC.hBmp)

End Sub

Private Sub UserControl_Show()

    If UserControl.Ambient.UserMode = True Then
        tmrSystemClockTimer.Enabled = True
    End If

End Sub

Private Sub UserControl_Terminate()

    If Not memDC.DC = 0 Then
        DeleteObject SelectObject(memDC.DC, memDC.hOldBmp)
        DeleteDC memDC.DC
    End If

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Style", mStyle, 1)
    Call PropBag.WriteProperty("ShowSecondHand", mShowSecondHand, True)
    Call PropBag.WriteProperty("ShowLogoImage", mShowLogoImage, False)
    Call PropBag.WriteProperty("ShowHighlights", mShowHighLights, True)
    Call PropBag.WriteProperty("Opacity_Background", mOpacity_Background, 100)
    Call PropBag.WriteProperty("Opacity_Hand_Hour", mOpacity_Hand_Hour, 100)
    Call PropBag.WriteProperty("Opacity_Hand_Minute", mOpacity_Hand_Minute, 100)
    Call PropBag.WriteProperty("Opacity_Hand_Second", mOpacity_Hand_Second, 100)
    Call PropBag.WriteProperty("Opacity_Highlights", mOpacity_Highlights, 100)
    Call PropBag.WriteProperty("Opacity_LogoImage", mOpacity_LogoImage, 100)
    Call PropBag.WriteProperty("LogoImage_Height", mLogoImage.Height, -1)
    Call PropBag.WriteProperty("LogoImage_Width", mLogoImage.Width, -1)
    Call PropBag.WriteProperty("LogoImage_Top", mLogoImage.Top, -1)
    Call PropBag.WriteProperty("LogoImage_Left", mLogoImage.Left, -1)
    Call PropBag.WriteProperty("TimeZoneOffset", mTimeZoneOffset, 0)
    Call PropBag.WriteProperty("SoundPlayTick", mSoundPlayTick, 0)
    Call PropBag.WriteProperty("SoundTickStyle", mSoundTickStyle, 0)
    Call PropBag.WriteProperty("ClockSize", mClockSize, 130)
    Call PropBag.WriteProperty("HighlightRotation", mHighlightRotation, 0)
    Call PropBag.WriteProperty("Color_Highlight_Hue", mColor_Highlight.Hue, 0)
    Call PropBag.WriteProperty("Color_Highlight_Saturation", mColor_Highlight.Saturation, 0.5)
    Call PropBag.WriteProperty("Color_Highlight_Luminosity", mColor_Highlight.Luminosity, 1)
    Call PropBag.WriteProperty("HandOffset_Hour_Horz", mHands_Offset(eHour, eHorz), 0)
    Call PropBag.WriteProperty("HandOffset_Hour_Vert", mHands_Offset(eHour, eVert), 0)
    Call PropBag.WriteProperty("HandOffset_Minute_Horz", mHands_Offset(eMinute, eHorz), 0)
    Call PropBag.WriteProperty("HandOffset_Minute_Vert", mHands_Offset(eMinute, eVert), 0)
    Call PropBag.WriteProperty("HandOffset_Second_Horz", mHands_Offset(eSecond, eHorz), 0)
    Call PropBag.WriteProperty("HandOffset_Second_Vert", mHands_Offset(eSecond, eVert), 0)

End Sub

