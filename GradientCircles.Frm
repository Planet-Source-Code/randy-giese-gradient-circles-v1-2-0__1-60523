VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form GradientCircles 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10245
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawWidth       =   2
   Icon            =   "GradientCircles.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   118
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   683
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picSpeed 
      Height          =   1035
      Left            =   150
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   656
      TabIndex        =   0
      Top             =   150
      Visible         =   0   'False
      Width           =   9900
      Begin MSComctlLib.Slider sldSpeed 
         Height          =   495
         Left            =   0
         TabIndex        =   1
         Top             =   495
         Width           =   9900
         _ExtentX        =   17463
         _ExtentY        =   873
         _Version        =   393216
         Min             =   -50
         Max             =   50
         TickFrequency   =   5
      End
      Begin VB.Label lblSpeed 
         Alignment       =   2  'Center
         Caption         =   "Faster"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   6600
         TabIndex        =   4
         Top             =   75
         Width           =   3300
      End
      Begin VB.Label lblSpeed 
         Alignment       =   2  'Center
         Caption         =   "Slower"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   0
         TabIndex        =   3
         Top             =   75
         Width           =   3300
      End
      Begin VB.Label lblSpeed 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Display Speed!"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   3300
         TabIndex        =   2
         Top             =   75
         Width           =   3300
      End
   End
   Begin VB.Timer tmrSpeed 
      Interval        =   1000
      Left            =   150
      Top             =   1320
   End
End
Attribute VB_Name = "GradientCircles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    GradientCircles                                              ****
'   ****                                                                 ****
'   ****    Written by:  Randy Giese                                     ****
'   ****    Written at:  www.RandyGrams.com                              ****
'   ****                                                                 ****
'   *************************************************************************
'   ****                                                                 ****
'   ****    Thank you Robert Rayment !!!!!                               ****
'   ****                                                                 ****
'   *************************************************************************
'   ****                                                                 ****
'   ****    I received a great deal of help from Robert Rayment in       ****
'   ****    debugging this program.  Robert; I am deeply indebted to     ****
'   ****    you for all of your assistance.                              ****
'   ****                                                                 ****
'   *************************************************************************
'   ****                                                                 ****
'   ****    I wrote GradientCircles a few years ago as part of a         ****
'   ****    larger project.  It has always been one of my favorite       ****
'   ****    Graphics programs.  But then I'm a sucker for colorful       ****
'   ****    random graphics.  I've seen some other interesting           ****
'   ****    graphic displays on PSC that are based on a specific         ****
'   ****    formula or pattern, so that once you see it.... you've       ****
'   ****    seen it.  In other words, it does one thing nicely, but      ****
'   ****    that's the only thing it can do.                             ****
'   ****    GradientCircles is my example of just the opposite.  It      ****
'   ****    is colorful and always changing.  The colors of the          ****
'   ****    circles are always different, the circle size always         ****
'   ****    changes and sometimes the circle builds from the center      ****
'   ****    out and other times it shrinks from the outer edge to the    ****
'   ****    center.                                                      ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'

Option Explicit

'** Variables **
Private blnPause                As Boolean
Private blnSpeed                As Boolean

Private curEnd                  As Currency
Private curStart                As Currency

Private lngCntr                 As Long
Private lngLoop2                As Long
Private lngPorM(0 To 3)         As Long
Private lngRGB(0 To 3)          As Long
Private lngStopTime             As Long

'   Define the constants for GetQueueStatus
Private Const QS_HOTKEY         As Long = &H80
Private Const QS_KEY            As Long = &H1
Private Const QS_MOUSEBUTTON    As Long = &H4
Private Const QS_MOUSEMOVE      As Long = &H2
Private Const QS_PAINT          As Long = &H20
Private Const QS_POSTMESSAGE    As Long = &H8
Private Const QS_SENDMESSAGE    As Long = &H40
Private Const QS_TIMER          As Long = &H10
Private Const QS_ALLINPUT       As Double = (QS_SENDMESSAGE Or QS_PAINT Or QS_TIMER Or QS_POSTMESSAGE Or QS_MOUSEBUTTON Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)

'** API Declarations **
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetQueueStatus Lib "user32" (ByVal fuFlags As Long) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Private Sub Form_DblClick()

    End

End Sub

Private Sub Form_Load()

    QueryPerformanceFrequency curFrequency  '  Get the timer frequency
    curFrequency = (curFrequency / 1000)    '  Break it down to a millisecond
    curFreq = curFrequency
    sldSpeed.Value = 0

    blnPause = False
    Randomize Timer
    Me.Show

    subGradientCircles

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
'       Right Click to open the Speed Menu.
        blnSpeed = True
        GradientSpeed.Show
        lngStartTime = Timer
    Else
'       Left Click to Pause the display.
        blnPause = Not blnPause

        Do
            DoEvents
        Loop While blnPause

        subGradientCircles
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next

    For lngCntr = (Forms.Count - 1) To 0 Step -1
        Unload Forms(lngCntr)               '  Unloads ALL forms
    Next lngCntr

    Close
    End

    On Error GoTo 0

End Sub

Private Sub subChangeColors()

'   This updates the Gradient's colors
    lngRGB(1) = lngRGB(1) + lngPorM(1)      '  Change first color value

'   Check all three color groups to make sure they are within legal bounds.
    For lngCntr = 1 To 3
        If lngRGB(lngCntr) < 2 Or lngRGB(lngCntr) > 253 Then   '  If color reaches min or max,
            subIncrementColors
        End If
    Next lngCntr

    If blnSpeed = True Then
'       Hide the Speed control if it's been visible for 5 seconds.
        lngStopTime = Timer

        If lngStopTime > lngStartTime + 5 Then
            Unload GradientSpeed
            blnSpeed = False
            Cls
        End If
    End If

End Sub

Private Sub subGradientCircles()

'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    G r a d i e n t   C i r c l e s                              ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'

Dim lngCenterCol As Long
Dim lngCenterRow As Long
Dim lngCircleDir As Long
Dim lngFrom      As Long
Dim lngRadius    As Long
Dim lngTo        As Long

    Do
'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    Determine the Size and Location of the Circle.               ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'
        lngCenterCol = ((ScaleWidth * 0.75) * Rnd) + (ScaleWidth * 0.125)     '  Starting Column
        lngCenterRow = ((ScaleHeight * 0.75) * Rnd) + (ScaleHeight * 0.125)   '  Starting Row
        lngCircleDir = Sgn(Rnd - 0.5)       '  Initial Direction of Color Change

'       This allows the radius to be anywhere up to 1/2 screen height.
'        lngRadius = ScaleHeight * 0.5 * Rnd '  Radius of Circle

'       This makes the radius anywhere from 1/4 to 1/2 screen height.
        lngRadius = (ScaleHeight * 0.25 * Rnd) + (ScaleHeight * 0.25) '  Radius of Circle

'       Determine whether the Circle Shrinks or Grows
        If lngCircleDir > 0 Then
'           The Circle Grows
            lngFrom = 1
            lngTo = lngRadius
        Else  '  Not lngCircleDir
'           The Circle Shrinks
            lngFrom = lngRadius
            lngTo = 1
        End If
'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    Determine the Circle's colors.                               ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'
        subGradientColor     '  Setup the initial Color Values and direction of change
'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    Draw the Circles.                                            ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'
        For lngLoop2 = lngFrom To lngTo Step lngCircleDir
            QueryPerformanceCounter curStart               '  Get the start time

            Me.ForeColor = RGB(lngRGB(1), lngRGB(2), lngRGB(3))
            Ellipse Me.hdc, lngCenterCol + lngLoop2, lngCenterRow + lngLoop2, lngCenterCol - lngLoop2, lngCenterRow - lngLoop2

'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    Comment out this statement:                                  ****
'   ****                                                                 ****
            subSlowDown
'   ****                                                                 ****
'   ****    to make GradientCircles run at max speed.                    ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'
            subChangeColors
        Next lngLoop2
    Loop

End Sub

Private Sub subGradientColor()

'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    Get the initial color and direction of the color change.     ****
'   ****                                                                 ****
'   ****    lngPorM(x) = Plus 1 or Minus 1.                              ****
'   ****    If lngPorM(x) = 1, then lngRGB(x) = 0                        ****
'   ****    If lngPorM(x) = -1, then lngRGB(x) = 255                     ****
'   ****                                                                 ****
'   ****    In other words, if lngPorM(1) = 1, then lngRGB(1) = 0.       ****
'   ****    Next time it goes to subIncrementColors, lngRGB(1) = 1,      ****
'   ****    then lngRGB(1) = 2, etc., up to 255.                         ****
'   ****                                                                 ****
'   ****    When a color reaches 255, lngPorM will change to a '-1'      ****
'   ****    and the lngRGB numbers will start getting smaller...         ****
'   ****    until it reaches '0' and starts the whole process over       ****
'   ****    again.                                                       ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'

Dim lngCounter      As Long

    For lngCntr = 1 To 3     '  Setup the initial Color Values and direction of change
'        lngPorM(lngCntr) = (Sgn(Rnd - 0.5)) * (Int(2 * Rnd) + 1) '  A Plus 1 or Minus 1 - lngPorM neans Plus or Minus
        lngPorM(lngCntr) = (Sgn(Rnd - 0.5)) '  A Plus 1 or Minus 1 - lngPorM neans Plus or Minus

        If lngCounter < 5 Then
'           This gives more randomness to the colors of the circles.
            lngRGB(lngCntr) = Int(254 * Rnd) + 1
        Else
'           This might provide better looking colors for the circles.
            If lngPorM(lngCntr) < 0 Then
                lngRGB(lngCntr) = 255
            Else  '  Not lngPorM(lngCntr)
                lngRGB(lngCntr) = 0
            End If
        End If
    Next lngCntr

    If lngCounter < 5 Then
        lngCounter = lngCounter + 1
    Else
        lngCounter = 0
    End If

End Sub

Private Sub subIncrementColors()

    lngPorM(lngCntr) = -lngPorM(lngCntr)    '  Reverse the color direction
    If lngCntr < 3 Then      '  If it's not the last color group (Green),
        lngRGB(lngCntr + 1) = lngRGB(lngCntr + 1) + lngPorM(lngCntr + 1)  '  then add 1 to next color group
    End If

End Sub

Private Sub subSlowDown()

'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    subSlowDown causes a very brief pause between the drawing    ****
'   ****    of each ring of the circle.  GradientCircles runs way too    ****
'   ****    fast on my PC without some sort of Slow Down routine.  If    ****
'   ****    you think it should run faster, try commenting out the       ****
'   ****    subSlowDown statement right after the 'Ellipse' statement    ****
'   ****    in 'subGradientCircles'.  If that looks better to you        ****
'   ****    then you can always delete this SubRoutine later.            ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'

    Do While curEnd < curStart + curFreq
        QueryPerformanceCounter curEnd      '  Get the end time

        If GetQueueStatus(QS_ALLINPUT) Then '  Yield for other processing.
            Exit Do
        End If
    Loop

    If GetQueueStatus(QS_ALLINPUT) Then '  Yield for other processing.
        DoEvents
    End If

End Sub
