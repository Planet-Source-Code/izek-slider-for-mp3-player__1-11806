VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2205
   DrawMode        =   1  'Blackness
   LinkTopic       =   "Form1"
   ScaleHeight     =   900
   ScaleWidth      =   2205
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   480
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.mp3"
      InitDir         =   "c:\"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "load"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1560
      Top             =   960
   End
   Begin VB.Line Line1 
      X1              =   45
      X2              =   2145
      Y1              =   120
      Y2              =   120
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   -1  'True
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   0
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pos, pos2, dragg As Boolean
Private Sub Command1_Click()
'show only mp3's in the open window
CommonDialog1.Filter = "MP3 Files|*.mp3"
'show the open window
CommonDialog1.ShowOpen
'load an mp3 into media player
temp$ = CommonDialog1.FileName
If temp$ <> "" Then MediaPlayer1.FileName = temp$
End Sub
Private Sub Form_Load()
'set the sliders position and draw the slider
pos = 45
Line (pos, 40)-(pos + 200, 200), , BF
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'if the slider is clicked on the disable the timer and enable dragging
If X >= pos And X <= pos + 200 And Y >= 40 And Y <= 200 Then dragg = True: Timer1.Enabled = False
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If dragg = True Then
    'clear the form
    Cls
    'the next 2 lines prevent the slider from going
    'outside of its range
    If X <= 145 Then X = 145
    If X >= 2045 Then X = 2045
    'draw the slider
    Line (X - 100, 40)-(X + 100, 200), , BF
    'set the position
    pos = X - 100
End If
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'disable dragging
dragg = False
'get the new position at which media player should play
On Error Resume Next
pos2 = Round(pos - 45) / Int(2000 / (MediaPlayer1.Duration))
'check if its below 0
If pos2 < 0 Then pos2 = 0
'check if its greater than duration
If pos2 > MediaPlayer1.Duration Then pos2 = MediaPlayer1.Duration
'set the position
MediaPlayer1.CurrentPosition = pos2
'enbable the timer
Timer1.Enabled = True
End Sub
Private Sub Timer1_Timer()
'clear the screen
Cls
'draw the slider
Line (pos, 40)-(pos + 200, 200), , BF
On Error Resume Next
'set new point for the slider
pos = 45 + (Int(MediaPlayer1.CurrentPosition) * Int(2000 / (MediaPlayer1.Duration)))
'prevent it from going outside of its range
If pos + 200 >= 2145 Then pos = 1945
End Sub
