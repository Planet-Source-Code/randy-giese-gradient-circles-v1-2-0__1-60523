VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form GradientSpeed 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10245
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   89
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   683
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSpeed 
      Height          =   1035
      Left            =   150
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   656
      TabIndex        =   0
      Top             =   150
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
         TabIndex        =   2
         Top             =   75
         Width           =   3300
      End
   End
End
Attribute VB_Name = "GradientSpeed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()

    With Me
        .Top = 100
        .Left = (Screen.Width - .Width) \ 2
    End With  '  Me

    sldSpeed.Value = lngSpeed

End Sub

'
'   *************************************************************************
'   *************************************************************************
'   ****                                                                 ****
'   ****    These are the 'Slider' routines for the Speed control.       ****
'   ****                                                                 ****
'   *************************************************************************
'   *************************************************************************
'

Private Sub sldSpeed_Change()

'
'   *************************************************************************
'   ****                                                                 ****
'   ****    The lngSpeed variable is used to maintain the position of    ****
'   ****    the pointer even though the form is unloaded.                ****
'   ****                                                                 ****
'   *************************************************************************
'

    lngSpeed = sldSpeed.Value

    Select Case lngSpeed
    Case Is < 0
'       Slow down the circles.
        curFreq = curFrequency * Abs(lngSpeed)
    Case 0
'       Default speed of the circles.
        curFreq = curFrequency
    Case Is > 0
'       Speed up the circles.
        curFreq = curFrequency / lngSpeed
    End Select

End Sub

Private Sub sldSpeed_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

'   This starts the 5 second timer.  The form is unloaded after 5 seconds.
    lngStartTime = Timer

End Sub
