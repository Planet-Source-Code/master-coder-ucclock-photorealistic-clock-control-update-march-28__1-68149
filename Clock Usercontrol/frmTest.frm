VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form ClockTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clock UserControl Test (Best Compiled)"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   11280
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTest.frx":030A
   ScaleHeight     =   562
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   752
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin ComCtl2.UpDown udTimeZoneOffset 
      Height          =   270
      Left            =   4290
      TabIndex        =   57
      Top             =   7845
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   476
      _Version        =   327681
      Alignment       =   0
      BuddyControl    =   "txtTimeZoneOffset"
      BuddyDispid     =   196609
      OrigLeft        =   226
      OrigTop         =   504
      OrigRight       =   243
      OrigBottom      =   522
      Max             =   11
      Min             =   -11
      SyncBuddy       =   -1  'True
      Wrap            =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin ComCtl2.UpDown udHandOffset 
      Height          =   270
      Index           =   2
      Left            =   3270
      TabIndex        =   48
      Top             =   7620
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   476
      _Version        =   327681
      BuddyControl    =   "txtOffset(2)"
      BuddyDispid     =   196610
      BuddyIndex      =   2
      OrigLeft        =   226
      OrigTop         =   504
      OrigRight       =   243
      OrigBottom      =   522
      Max             =   256
      Min             =   -256
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtTimeZoneOffset 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   4275
      TabIndex        =   58
      Text            =   "0"
      Top             =   7830
      Width           =   690
   End
   Begin ComCtl2.UpDown udHandOffset 
      Height          =   270
      Index           =   1
      Left            =   2280
      TabIndex        =   46
      Top             =   7620
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   476
      _Version        =   327681
      BuddyControl    =   "txtOffset(1)"
      BuddyDispid     =   196610
      BuddyIndex      =   1
      OrigLeft        =   157
      OrigTop         =   504
      OrigRight       =   174
      OrigBottom      =   522
      Max             =   256
      Min             =   -256
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtOffset 
      Height          =   300
      Index           =   1
      Left            =   1860
      TabIndex        =   47
      Text            =   "0"
      Top             =   7605
      Width           =   690
   End
   Begin PrjClockTest.ucFrame UserControl11 
      Height          =   750
      Left            =   405
      TabIndex        =   52
      Top             =   7335
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1323
      Begin VB.OptionButton optHand 
         Caption         =   "Hour"
         Height          =   240
         Index           =   1
         Left            =   0
         TabIndex        =   55
         Top             =   0
         Value           =   -1  'True
         Width           =   945
      End
      Begin VB.OptionButton optHand 
         Caption         =   "Minute"
         Height          =   240
         Index           =   2
         Left            =   0
         TabIndex        =   54
         Top             =   270
         Width           =   945
      End
      Begin VB.OptionButton optHand 
         Caption         =   "Second"
         Height          =   240
         Index           =   3
         Left            =   0
         TabIndex        =   53
         Top             =   525
         Width           =   900
      End
   End
   Begin VB.TextBox txtOffset 
      Height          =   300
      Index           =   2
      Left            =   2850
      TabIndex        =   49
      Text            =   "0"
      Top             =   7605
      Width           =   690
   End
   Begin VB.CheckBox chkPlayTick 
      Caption         =   "Audible Second Hand "
      Height          =   690
      Left            =   7200
      TabIndex        =   45
      Top             =   7545
      Width           =   945
   End
   Begin ComctlLib.Slider sldHighlightRotation 
      Height          =   300
      Left            =   270
      TabIndex        =   37
      Top             =   4920
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   529
      _Version        =   327682
      Max             =   359
      TickFrequency   =   26
   End
   Begin VB.CheckBox chkOptions 
      Caption         =   "Display Second Hand"
      Height          =   690
      Index           =   2
      Left            =   5985
      TabIndex        =   31
      Top             =   7545
      Value           =   1  'Checked
      Width           =   945
   End
   Begin VB.CheckBox chkOptions 
      Caption         =   "Display Logo"
      Height          =   255
      Index           =   1
      Left            =   405
      TabIndex        =   30
      Top             =   5430
      Width           =   1230
   End
   Begin VB.CheckBox chkOptions 
      Caption         =   "Display Highlights"
      Height          =   255
      Index           =   0
      Left            =   405
      TabIndex        =   29
      Top             =   4305
      Value           =   1  'Checked
      Width           =   1605
   End
   Begin ComctlLib.Slider sldLogo 
      Height          =   300
      Index           =   0
      Left            =   270
      TabIndex        =   21
      Top             =   5970
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   529
      _Version        =   327682
      Min             =   -1
      Max             =   256
      TickFrequency   =   26
   End
   Begin ComctlLib.Slider sldOpacity 
      Height          =   300
      Index           =   0
      Left            =   270
      TabIndex        =   9
      Top             =   2565
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   529
      _Version        =   327682
      Min             =   -100
      Max             =   0
      SelStart        =   -100
      TickFrequency   =   10
      Value           =   -100
   End
   Begin VB.OptionButton optStyle 
      Caption         =   "Style #8"
      Height          =   210
      Index           =   7
      Left            =   2760
      TabIndex        =   8
      Top             =   930
      Width           =   930
   End
   Begin VB.OptionButton optStyle 
      Caption         =   "Style #7"
      Height          =   210
      Index           =   6
      Left            =   2760
      TabIndex        =   7
      Top             =   660
      Width           =   930
   End
   Begin VB.OptionButton optStyle 
      Caption         =   "Style #6"
      Height          =   210
      Index           =   5
      Left            =   1575
      TabIndex        =   6
      Top             =   1215
      Width           =   930
   End
   Begin VB.OptionButton optStyle 
      Caption         =   "Style #5"
      Height          =   210
      Index           =   4
      Left            =   1575
      TabIndex        =   5
      Top             =   930
      Width           =   930
   End
   Begin VB.OptionButton optStyle 
      Caption         =   "Style #4"
      Height          =   210
      Index           =   3
      Left            =   1575
      TabIndex        =   4
      Top             =   660
      Width           =   930
   End
   Begin VB.OptionButton optStyle 
      Caption         =   "Style #3"
      Height          =   210
      Index           =   2
      Left            =   390
      TabIndex        =   3
      Top             =   1215
      Width           =   930
   End
   Begin VB.OptionButton optStyle 
      Caption         =   "Style #2"
      Height          =   210
      Index           =   1
      Left            =   390
      TabIndex        =   2
      Top             =   930
      Width           =   930
   End
   Begin VB.OptionButton optStyle 
      Caption         =   "Style #1"
      Height          =   210
      Index           =   0
      Left            =   390
      TabIndex        =   1
      Top             =   660
      Width           =   930
   End
   Begin ComctlLib.Slider sldOpacity 
      Height          =   300
      Index           =   1
      Left            =   270
      TabIndex        =   11
      Top             =   3195
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   529
      _Version        =   327682
      Min             =   -99
      Max             =   0
      SelStart        =   -99
      TickFrequency   =   10
      Value           =   -99
   End
   Begin ComctlLib.Slider sldOpacity 
      Height          =   300
      Index           =   2
      Left            =   270
      TabIndex        =   13
      Top             =   3780
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   529
      _Version        =   327682
      Min             =   -100
      Max             =   0
      SelStart        =   -100
      TickFrequency   =   10
      Value           =   -100
   End
   Begin ComctlLib.Slider sldOpacity 
      Height          =   300
      Index           =   3
      Left            =   2055
      TabIndex        =   15
      Top             =   2565
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   529
      _Version        =   327682
      Min             =   -100
      Max             =   0
      SelStart        =   -100
      TickFrequency   =   10
      Value           =   -100
   End
   Begin ComctlLib.Slider sldOpacity 
      Height          =   300
      Index           =   4
      Left            =   2055
      TabIndex        =   17
      Top             =   3195
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   529
      _Version        =   327682
      Min             =   -100
      Max             =   0
      SelStart        =   -100
      TickFrequency   =   10
      Value           =   -100
   End
   Begin ComctlLib.Slider sldOpacity 
      Height          =   300
      Index           =   5
      Left            =   2055
      TabIndex        =   19
      Top             =   3780
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   529
      _Version        =   327682
      Min             =   -100
      Max             =   0
      SelStart        =   -100
      TickFrequency   =   10
      Value           =   -100
   End
   Begin ComctlLib.Slider sldLogo 
      Height          =   300
      Index           =   2
      Left            =   2055
      TabIndex        =   23
      Top             =   5985
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   529
      _Version        =   327682
      Min             =   -1
      Max             =   256
      TickFrequency   =   26
   End
   Begin ComctlLib.Slider sldLogo 
      Height          =   300
      Index           =   1
      Left            =   270
      TabIndex        =   25
      Top             =   6525
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   529
      _Version        =   327682
      Min             =   -1
      Max             =   256
      TickFrequency   =   26
   End
   Begin ComctlLib.Slider sldLogo 
      Height          =   300
      Index           =   3
      Left            =   2055
      TabIndex        =   27
      Top             =   6540
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   529
      _Version        =   327682
      Min             =   -1
      Max             =   256
      TickFrequency   =   26
   End
   Begin ComctlLib.Slider sldClockSize 
      Height          =   450
      Left            =   645
      TabIndex        =   34
      Top             =   1470
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   794
      _Version        =   327682
      Min             =   20
      Max             =   512
      SelStart        =   200
      TickFrequency   =   26
      Value           =   200
   End
   Begin ComctlLib.Slider sldHighlightColor 
      Height          =   300
      Index           =   1
      Left            =   2460
      TabIndex        =   39
      Top             =   4770
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   529
      _Version        =   327682
      Max             =   100
      SelStart        =   50
      TickStyle       =   3
      TickFrequency   =   26
      Value           =   50
   End
   Begin ComctlLib.Slider sldHighlightColor 
      Height          =   300
      Index           =   2
      Left            =   2460
      TabIndex        =   41
      Top             =   5040
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   529
      _Version        =   327682
      Max             =   100
      SelStart        =   100
      TickStyle       =   3
      TickFrequency   =   26
      Value           =   100
   End
   Begin ComctlLib.Slider sldHighlightColor 
      Height          =   300
      Index           =   0
      Left            =   2460
      TabIndex        =   42
      Top             =   4500
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   529
      _Version        =   327682
      Max             =   100
      TickStyle       =   3
      TickFrequency   =   26
   End
   Begin PrjClockTest.ucClock ucClock 
      Height          =   3000
      Left            =   6255
      Top             =   2295
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   5292
      SoundTickStyle  =   "Tick-1"
      ClockSize       =   200
      Color_Highlight_Saturation=   1
      Color_Highlight_Luminosity=   0
      HandOffset_Hour_Horz=   1
      HandOffset_Hour_Vert=   2
      HandOffset_Minute_Horz=   3
      HandOffset_Minute_Vert=   4
      HandOffset_Second_Horz=   5
      HandOffset_Second_Vert=   6
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0C0&
      Height          =   855
      Left            =   345
      Shape           =   4  'Rounded Rectangle
      Top             =   7275
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Hours"
      Height          =   195
      Index           =   10
      Left            =   5010
      TabIndex        =   59
      Top             =   7875
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Time Zone Offset"
      Height          =   195
      Index           =   9
      Left            =   4215
      TabIndex        =   56
      Top             =   7500
      Width           =   1230
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00B2AFB0&
      BorderColor     =   &H00C0C000&
      FillColor       =   &H8000000F&
      FillStyle       =   0  'Solid
      Height          =   660
      Index           =   0
      Left            =   4020
      Shape           =   4  'Rounded Rectangle
      Top             =   7590
      Width           =   1665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Vert."
      Height          =   195
      Index           =   8
      Left            =   2910
      TabIndex        =   51
      Top             =   7395
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Horz."
      Height          =   195
      Index           =   7
      Left            =   1890
      TabIndex        =   50
      Top             =   7395
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Lum"
      Height          =   195
      Index           =   6
      Left            =   2145
      TabIndex        =   44
      ToolTipText     =   "Luminosity"
      Top             =   5055
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Hue"
      Height          =   195
      Index           =   4
      Left            =   2145
      TabIndex        =   43
      Top             =   4515
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sat"
      Height          =   195
      Index           =   5
      Left            =   2145
      TabIndex        =   40
      ToolTipText     =   "Saturation"
      Top             =   4785
      Width           =   240
   End
   Begin VB.Label lblHighlightRotation 
      AutoSize        =   -1  'True
      Caption         =   "Rotation [0]"
      Height          =   195
      Left            =   360
      TabIndex        =   38
      Tag             =   "Rotation"
      Top             =   4710
      Width           =   825
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BorderColor     =   &H00C0C000&
      FillColor       =   &H8000000F&
      FillStyle       =   0  'Solid
      Height          =   960
      Index           =   5
      Left            =   210
      Shape           =   4  'Rounded Rectangle
      Top             =   4425
      Width           =   3615
   End
   Begin VB.Label lblClockSize 
      AutoSize        =   -1  'True
      Caption         =   "[200x200]"
      Height          =   195
      Left            =   3000
      TabIndex        =   36
      Top             =   1530
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Size"
      Height          =   195
      Index           =   3
      Left            =   330
      TabIndex        =   35
      Top             =   1545
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   " Hand Position "
      Height          =   195
      Index           =   2
      Left            =   405
      TabIndex        =   33
      Top             =   6990
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   " Opacity Level"
      Height          =   195
      Index           =   1
      Left            =   405
      TabIndex        =   32
      Top             =   2055
      Width           =   1020
   End
   Begin VB.Label lblLogo 
      AutoSize        =   -1  'True
      Caption         =   "Logo Left [-1]"
      Height          =   195
      Index           =   3
      Left            =   2145
      TabIndex        =   28
      Tag             =   "Logo Left"
      Top             =   6345
      Width           =   945
   End
   Begin VB.Label lblLogo 
      AutoSize        =   -1  'True
      Caption         =   "Logo Height [-1]"
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   26
      Tag             =   "Logo Height"
      Top             =   6330
      Width           =   1140
   End
   Begin VB.Label lblLogo 
      AutoSize        =   -1  'True
      Caption         =   "Logo Top [-1]"
      Height          =   195
      Index           =   2
      Left            =   2145
      TabIndex        =   24
      Tag             =   "Logo Top"
      Top             =   5790
      Width           =   960
   End
   Begin VB.Label lblLogo 
      AutoSize        =   -1  'True
      Caption         =   "Logo Width [-1]"
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   22
      Tag             =   "Logo Width"
      Top             =   5775
      Width           =   1095
   End
   Begin VB.Label lblOpacity 
      AutoSize        =   -1  'True
      Caption         =   "Second Hand [100]"
      Height          =   195
      Index           =   5
      Left            =   2145
      TabIndex        =   20
      Tag             =   "Second Hand"
      Top             =   3585
      Width           =   1395
   End
   Begin VB.Label lblOpacity 
      AutoSize        =   -1  'True
      Caption         =   "Minute Hand [100]"
      Height          =   195
      Index           =   4
      Left            =   2145
      TabIndex        =   18
      Tag             =   "Minute Hand"
      Top             =   3000
      Width           =   1320
   End
   Begin VB.Label lblOpacity 
      AutoSize        =   -1  'True
      Caption         =   "Hour Hand [100]"
      Height          =   195
      Index           =   3
      Left            =   2145
      TabIndex        =   16
      Tag             =   "Hour Hand"
      Top             =   2355
      Width           =   1185
   End
   Begin VB.Label lblOpacity 
      AutoSize        =   -1  'True
      Caption         =   "Logo Image [100]"
      Height          =   195
      Index           =   2
      Left            =   360
      TabIndex        =   14
      Tag             =   "Logo Image"
      Top             =   3585
      Width           =   1245
   End
   Begin VB.Label lblOpacity 
      AutoSize        =   -1  'True
      Caption         =   "Highlights [99]"
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   12
      Tag             =   "Highlights"
      Top             =   3000
      Width           =   1005
   End
   Begin VB.Label lblOpacity 
      AutoSize        =   -1  'True
      Caption         =   "Background [100]"
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   10
      Tag             =   "Background"
      Top             =   2355
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   " Clock Style and Size "
      Height          =   195
      Index           =   0
      Left            =   420
      TabIndex        =   0
      Top             =   270
      Width           =   1545
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BorderColor     =   &H00C0C000&
      FillColor       =   &H8000000F&
      FillStyle       =   0  'Solid
      Height          =   1575
      Index           =   1
      Left            =   210
      Shape           =   4  'Rounded Rectangle
      Top             =   375
      Width           =   3615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BorderColor     =   &H00C0C000&
      FillColor       =   &H8000000F&
      FillStyle       =   0  'Solid
      Height          =   2130
      Index           =   2
      Left            =   210
      Shape           =   4  'Rounded Rectangle
      Top             =   2130
      Width           =   3615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BorderColor     =   &H00C0C000&
      FillColor       =   &H8000000F&
      FillStyle       =   0  'Solid
      Height          =   1365
      Index           =   3
      Left            =   210
      Shape           =   4  'Rounded Rectangle
      Top             =   5565
      Width           =   3615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00B2AFB0&
      BorderColor     =   &H00C0C000&
      FillColor       =   &H8000000F&
      FillStyle       =   0  'Solid
      Height          =   1140
      Index           =   4
      Left            =   210
      Shape           =   4  'Rounded Rectangle
      Top             =   7080
      Width           =   3615
   End
End
Attribute VB_Name = "ClockTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long

Private Const ShadowboxCenter_X = 507
Private Const ShadowboxCenter_Y = 250


Private Sub chkOptions_Click(Index As Integer)

  Dim bvalue As Boolean

    bvalue = IIf(chkOptions(Index) = 0, False, True)

    Select Case Index
     Case 0  'Highlights
        ucClock.ShowHighlights = bvalue

     Case 1  'Logo
        ucClock.ShowLogoImage = bvalue

     Case 2  'Second Hand
        ucClock.ShowSecondHand = bvalue

    End Select

End Sub

Private Sub chkPlayTick_Click()

    ucClock.SoundPlayTick = IIf(chkPlayTick = 0, False, True)

End Sub

Private Sub chkShowHighlights_Click()

  Dim clock As Control

    For Each clock In Controls
        Debug.Print clock.Name

        If Left(LCase(clock.Name), 7) = "ucclock" Then
            clock.ShowHighlights = Not clock.ShowHighlights
        End If

    Next clock

End Sub

Private Sub chkShowSecondHands_Click()

  Dim clock As Control

    For Each clock In Controls
        Debug.Print clock.Name

        If Left(LCase(clock.Name), 7) = "ucclock" Then
            clock.ShowSecondHand = Not clock.ShowSecondHand
        End If

    Next clock

End Sub

Private Sub Form_Initialize()

    InitCommonControls
    ucClock.Opacity_Highlights = 99

    ' AlarmAdd
    '
    '  Syntax:  .AlarmAdd Hour, Minute, Second, "Key", "Comment"
    '
    '            Hour    as Integer (0 to 23)   24-hour mode
    '            Minute  as Integer (0 to 59)
    '            Second  as Integer (0 to 59)
    '            Key     as String  (The Key can be used as a reference for deleting an Alarm)
    '            Comment as string  (Used as a reference, event returns this Comment)
    '
    '  Returns:  Index of entry as an Integer
    
    
    ' AlarmDelete
    '
    '  Syntax:  .AlarmDelete "Index or Key"
    '
    '  Returns:  Result of deletion as True/False (Boolean)
    
    
    ' AlarmRemoveAll
    '
    '  Syntax:  .AlarmRemoveAll
    '
    '  Returns:  Nothing
    
    With ucClock
        
        .AlarmAdd 22, 13, 0, "First time", "Sample Alarm 1"
        
    End With

End Sub

Private Sub optHand_Click(Index As Integer)

    Select Case Index
     Case 1  'Hour
        udHandOffset(eHorz).Value = ucClock.HandOffset_Hour_Horz
        udHandOffset(evert).Value = ucClock.HandOffset_Hour_Vert

     Case 2  'Minute
        udHandOffset(eHorz).Value = ucClock.HandOffset_Minute_Horz
        udHandOffset(evert).Value = ucClock.HandOffset_Minute_Vert

     Case 3  'Second
        udHandOffset(eHorz).Value = ucClock.HandOffset_Second_Horz
        udHandOffset(evert).Value = ucClock.HandOffset_Second_Vert

    End Select

End Sub

Private Sub optStyle_Click(Index As Integer)

    ucClock.Style = Index + 1

End Sub

Private Sub sldClockSize_Scroll()

    ucClock.Move ShadowboxCenter_X - (sldClockSize.Value / 2), ShadowboxCenter_Y - (sldClockSize.Value / 2)

    ucClock.ClockSize = sldClockSize.Value

    lblClockSize.Caption = "[" & sldClockSize.Value & "x" & sldClockSize.Value & "]"

End Sub

Private Sub sldHighlightColor_Scroll(Index As Integer)

    Select Case Index
     Case 0  'Hue
        ucClock.Color_Highlight_Hue = sldHighlightColor(Index).Value / 100

     Case 1  'Saturation
        ucClock.Color_Highlight_Saturation = sldHighlightColor(Index).Value / 100

     Case 2  'Luminosity
        ucClock.Color_Highlight_Luminosity = sldHighlightColor(Index).Value / 100

    End Select

End Sub

Private Sub sldHighlightRotation_Scroll()

    ucClock.HighlightRotation = sldHighlightRotation.Value
    lblHighlightRotation.Caption = lblHighlightRotation.Tag & " [" & sldHighlightRotation.Value & "]"

End Sub

Private Sub sldLogo_Scroll(Index As Integer)

    Select Case Index
     Case 0  'Width
        ucClock.LogoImage_Width = sldLogo(Index).Value

     Case 1  'Height
        ucClock.LogoImage_Height = sldLogo(Index).Value

     Case 2  'Top
        ucClock.LogoImage_Top = sldLogo(Index).Value

     Case 3  'Left
        ucClock.LogoImage_Left = sldLogo(Index).Value

    End Select

    lblLogo(Index).Caption = lblLogo(Index).Tag & " [" & sldLogo(Index).Value & "]"

End Sub

Private Sub sldOpacity_Scroll(Index As Integer)

  Dim iAdjustedValue As Integer

    iAdjustedValue = sldOpacity(Index).Value * -1

    sldOpacity(Index).ToolTipText = iAdjustedValue

    Select Case Index
     Case 0  'background
        ucClock.Opacity_Background = iAdjustedValue

     Case 1  'Highlights
        ucClock.Opacity_Highlights = iAdjustedValue

     Case 2  'Logo
        ucClock.Opacity_LogoImage = iAdjustedValue

     Case 3  'Hour Hand
        ucClock.Opacity_Hand_Hour = iAdjustedValue

     Case 4  'Minute Hand
        ucClock.Opacity_Hand_Minute = iAdjustedValue

     Case 5  'Second Hand
        ucClock.Opacity_Hand_Second = iAdjustedValue

    End Select

    lblOpacity(Index).Caption = lblOpacity(Index).Tag & " [" & iAdjustedValue & "]"

End Sub

Private Sub txtOffset_Change(Index As Integer)

    If optHand(eHour).Value Then

        ucClock.HandOffset_Hour_Horz = txtOffset(eHorz).Text
        ucClock.HandOffset_Hour_Vert = txtOffset(evert).Text

     ElseIf optHand(eMinute).Value Then

        ucClock.HandOffset_Minute_Horz = txtOffset(eHorz).Text
        ucClock.HandOffset_Minute_Vert = txtOffset(evert).Text

     ElseIf optHand(eSecond).Value Then

        ucClock.HandOffset_Second_Horz = txtOffset(eHorz).Text
        ucClock.HandOffset_Second_Vert = txtOffset(evert).Text

    End If

End Sub

Private Sub txtTimeZoneOffset_Change()

    ucClock.TimeZoneOffset = CInt(txtTimeZoneOffset.Text)

End Sub

Private Sub ucClock_Alarm(ByVal sHour As String, ByVal sMinute As String, ByVal sSecond As String, ByVal sKey As String, ByVal sComment As String)

    MsgBox "Clock Alarm: " & sComment & vbCrLf & vbCrLf & _
            "       Hour: " & sHour & vbCrLf & _
            "     Minute: " & sMinute & vbCrLf & _
            "     Second: " & sSecond & vbCrLf & _
            "        Key: " & sKey, _
            vbInformation, "Clock usercontrol"

End Sub

Private Sub ucClock_Click()

    MsgBox "UserControl was Clicked", vbInformation, "Clock usercontrol"

End Sub

Private Sub ucClock_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MsgBox "UserControl MouseDown event (Button = " & Button & ")", vbInformation, "Clock usercontrol"
    
End Sub

Private Sub ucClock_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MsgBox "UserControl MouseUp event (Button = " & Button & ")", vbInformation, "Clock usercontrol"
    
End Sub

