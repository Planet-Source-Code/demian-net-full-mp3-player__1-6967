VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   Caption         =   "Jagis Mp3 Player "
   ClientHeight    =   2100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4290
   Icon            =   "Jagis Mp3 Player.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Jagis Mp3 Player.frx":0CCA
   ScaleHeight     =   2100
   ScaleWidth      =   4290
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture15 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   3430
      Picture         =   "Jagis Mp3 Player.frx":532E
      ScaleHeight     =   300
      ScaleWidth      =   480
      TabIndex        =   23
      Top             =   1440
      Width           =   480
   End
   Begin VB.PictureBox Picture14 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   2950
      Picture         =   "Jagis Mp3 Player.frx":685D
      ScaleHeight     =   300
      ScaleWidth      =   495
      TabIndex        =   22
      Top             =   1440
      Width           =   495
   End
   Begin VB.PictureBox Picture13 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   720
      Picture         =   "Jagis Mp3 Player.frx":7DBF
      ScaleHeight     =   135
      ScaleWidth      =   420
      TabIndex        =   19
      Top             =   100
      Width           =   420
   End
   Begin VB.PictureBox Picture12 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   250
      Picture         =   "Jagis Mp3 Player.frx":9177
      ScaleHeight     =   135
      ScaleWidth      =   420
      TabIndex        =   18
      Top             =   100
      Width           =   420
   End
   Begin VB.PictureBox Picture9 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2560
      Picture         =   "Jagis Mp3 Player.frx":A534
      ScaleHeight     =   285
      ScaleWidth      =   330
      TabIndex        =   17
      Top             =   1440
      Width           =   330
   End
   Begin MediaPlayerCtl.MediaPlayer am1 
      Height          =   615
      Left            =   720
      TabIndex        =   16
      Top             =   2880
      Width           =   735
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   -1  'True
      Balance         =   30
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
      PlayCount       =   1
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
      Volume          =   -480
      WindowlessVideo =   0   'False
   End
   Begin VB.PictureBox Picture11 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   270
      Left            =   2760
      Picture         =   "Jagis Mp3 Player.frx":B9BC
      ScaleHeight     =   210
      ScaleWidth      =   630
      TabIndex        =   15
      Top             =   3720
      Width           =   690
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   240
      Top             =   3480
   End
   Begin VB.PictureBox Picture10 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1440
      ScaleHeight     =   435
      ScaleWidth      =   930
      TabIndex        =   14
      Top             =   3600
      Width           =   990
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   2880
      Top             =   3000
   End
   Begin VB.PictureBox Picture8 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2200
      Picture         =   "Jagis Mp3 Player.frx":CF2F
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   13
      Top             =   1440
      Width           =   360
   End
   Begin VB.PictureBox Picture7 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Picture         =   "Jagis Mp3 Player.frx":E3BB
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   12
      Top             =   1440
      Width           =   345
   End
   Begin VB.PictureBox Picture6 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   360
      Picture         =   "Jagis Mp3 Player.frx":F86D
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   11
      Top             =   1440
      Width           =   360
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   720
      Picture         =   "Jagis Mp3 Player.frx":10D44
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   10
      Top             =   1440
      Width           =   345
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1440
      Picture         =   "Jagis Mp3 Player.frx":121C2
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   9
      Top             =   1440
      Width           =   345
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1080
      Picture         =   "Jagis Mp3 Player.frx":1367C
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   8
      Top             =   1440
      Width           =   345
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   0
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   372
      TabIndex        =   7
      Top             =   2160
      Width           =   5640
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   135
      LargeChange     =   100
      Left            =   2400
      Max             =   5000
      Min             =   -5000
      SmallChange     =   50
      TabIndex        =   2
      Top             =   900
      Value           =   1
      Width           =   1455
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   135
      LargeChange     =   10
      Left            =   360
      Max             =   2500
      SmallChange     =   10
      TabIndex        =   1
      Top             =   900
      Value           =   2500
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4320
      Top             =   960
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   360
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   233
      TabIndex        =   0
      Top             =   1130
      Width           =   3500
   End
   Begin MSComDlg.CommonDialog Cd1 
      Left            =   120
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3480
      MousePointer    =   10  'Up Arrow
      TabIndex        =   21
      Top             =   30
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3000
      MousePointer    =   10  'Up Arrow
      TabIndex        =   20
      Top             =   30
      Width           =   375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Center"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   690
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Balance:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   690
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   690
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Volume:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   680
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long


Private Const RGN_AND = 1
Private Const RGN_COPY = 5
Private Const RGN_DIFF = 4
Private Const RGN_OR = 2
Private Const RGN_XOR = 3
Dim pimp As Integer, pump As Integer, foop As String, loopy As Boolean
Dim bMoveFrom As Boolean, LastPoint As POINTAPI, Pause As Boolean

Private Function CreateFormRegion() As Long

    Dim ResultRegion As Long, HolderRegion As Long, ObjectRegion As Long, nRet As Long
    Dim PolyPoints() As POINTAPI
    ResultRegion = CreateRectRgn(0, 0, 0, 0)
    HolderRegion = CreateRectRgn(0, 0, 0, 0)

'!Shaped Form Region Definition
'!2,3,2,284,139,8,12,1
    ObjectRegion = CreateRoundRectRgn(2, 3, 284, 136, 18, 20)
    nRet = CombineRgn(ResultRegion, ObjectRegion, ObjectRegion, RGN_COPY)
    DeleteObject ObjectRegion
    CreateFormRegion = ResultRegion
End Function


Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
FormOnTop Form1
 Dim nRet As Long
    nRet = SetWindowRgn(Me.hwnd, CreateFormRegion, True)
    
'BitBlt Picture4.hdc, 0, 0, 50, 50, Picture2.hdc, 64, 0, SRCCOPY
Picture4.Refresh

'BitBlt Picture3.hdc, 0, 0, 50, 50, Picture2.hdc, 0, 0, SRCCOPY
Picture3.Refresh

'BitBlt Picture5.hdc, 0, 0, 50, 50, Picture2.hdc, 123.7, 0, SRCCOPY
Picture5.Refresh

'BitBlt Picture6.hdc, 0, 0, 50, 50, Picture2.hdc, 185.7, 0, SRCCOPY
Picture6.Refresh

'BitBlt Picture7.hdc, 0, 0, 50, 50, Picture2.hdc, 248.5, 0, SRCCOPY
Picture7.Refresh

'BitBlt Picture8.hdc, 0, 0, 50, 50, Picture2.hdc, 309, 0, SRCCOPY
Picture8.Refresh

'BitBlt Picture9.hdc, 0, 0, 50, 50, Picture10.hdc, 0, 0, SRCCOPY
Picture9.Refresh

 iTPPX& = Screen.TwipsPerPixelX
    iTPPY& = Screen.TwipsPerPixelY
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim POINT As POINTAPI
    GetCursorPos POINT
    LastPoint.X = POINT.X
    LastPoint.Y = POINT.Y
    bMoveFrom = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim iDX As Long, iDY As Long
    Dim POINT As POINTAPI
    If Not bMoveFrom Then Exit Sub
    GetCursorPos POINT
    iDX& = (POINT.X - LastPoint.X) * iTPPX&
    iDY& = (POINT.Y - LastPoint.Y) * iTPPY&
    LastPoint.X = POINT.X
    LastPoint.Y = POINT.Y
    Me.Move Me.Left + iDX&, Me.Top + iDY&
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 bMoveFrom = False
End Sub

Private Sub HScroll1_Change()
Dim pim, sha
sha = HScroll1.Value - 2500
am1.Volume = sha
Dim foo As Integer, poo As Integer
On Error GoTo hell
poo = HScroll1.Min
foo = HScroll1.Value
Label2.Caption = foo \ 25
hell:
Exit Sub
End Sub

Private Sub HScroll1_Scroll()

Dim foo As Integer, poo As Integer
On Error GoTo hell
Dim pim, sha
sha = HScroll1.Value - 2500
am1.Volume = sha
poo = HScroll1.Min
foo = HScroll1.Value
Label2.Caption = foo \ 25
hell:
Exit Sub
End Sub

Private Sub HScroll2_Change()
On Error GoTo hell
If HScroll2.Value > -500 And HScroll2.Value < 500 Then
Label4.Caption = "Center"
End If
If HScroll2.Value < -500 Then
Label4.Caption = "Left"
End If
If HScroll2.Value > 500 Then
Label4.Caption = "Right"
End If
am1.Balance = HScroll2.Value
hell:
Exit Sub

End Sub

Private Sub HScroll2_Scroll()

If HScroll2.Value > -2500 And HScroll2.Value < 2500 Then
Label4.Caption = "Center"
End If
If HScroll2.Value < -2500 Then
Label4.Caption = "Left"
End If
If HScroll2.Value > 2500 Then
Label4.Caption = "Right"
End If
am1.Balance = HScroll2.Value
End Sub







Private Sub Label8_Click()

End Sub

Private Sub Label9_Click()

End Sub










Private Sub Label5_Click()
Form2.Show 1
End Sub

Private Sub Label6_Click()
Form3.Show 1
End Sub

Private Sub Picture12_Click()
Form1.WindowState = 1
End Sub

Private Sub Picture13_Click()
End
End Sub

Private Sub Picture14_Click()
PlayList.Show
End Sub

Private Sub Picture15_Click()
frmId3.Show
End Sub

Private Sub Picture3_Click()
On Error GoTo hell:
If Pause = True Then
Timer2.Enabled = True
Pause = False
End If
am1.Play
hell:
Exit Sub
End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'BitBlt Picture3.hdc, 0, 0, 50, 50, Picture2.hdc, 31, 0, SRCCOPY
Picture3.Refresh
End Sub

Private Sub Picture3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'BitBlt Picture3.hdc, 0, 0, 50, 50, Picture2.hdc, 0, 0, SRCCOPY
Picture3.Refresh
End Sub

Private Sub Picture4_Click()
On Error GoTo hell:
Pause = True
Timer2.Enabled = False
am1.Pause
hell:
Exit Sub
End Sub

Private Sub Picture4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'BitBlt Picture4.hdc, 0, 0, 50, 50, Picture2.hdc, 93.5, 0, SRCCOPY
Picture4.Refresh
End Sub

Private Sub Picture4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'BitBlt Picture4.hdc, 0, 0, 50, 50, Picture2.hdc, 63.5, 0, SRCCOPY
Picture4.Refresh
End Sub

Private Sub Picture5_Click()
On Error GoTo hell
am1.Stop
am1.CurrentPosition = 0
hell:
Exit Sub
End Sub

Private Sub Picture5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'BitBlt Picture5.hdc, 0, 0, 50, 50, Picture2.hdc, 154.5, 0, SRCCOPY
Picture5.Refresh
End Sub

Private Sub Picture5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'BitBlt Picture5.hdc, 0, 0, 50, 50, Picture2.hdc, 123.7, 0, SRCCOPY
Picture5.Refresh
End Sub

Private Sub Picture6_Click()
On Error GoTo hell
am1.CurrentPosition = am1.CurrentPosition - 5
hell:
Exit Sub
End Sub

Private Sub Picture6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'BitBlt Picture6.hdc, 0, 0, 50, 50, Picture2.hdc, 215, 0, SRCCOPY
Picture6.Refresh
End Sub

Private Sub Picture6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'BitBlt Picture6.hdc, 0, 0, 50, 50, Picture2.hdc, 185.7, 0, SRCCOPY
Picture6.Refresh
End Sub

Private Sub Picture7_Click()
On Error GoTo hell
am1.CurrentPosition = am1.CurrentPosition + 5
hell:
Exit Sub
End Sub

Private Sub Picture7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'BitBlt Picture7.hdc, 0, 0, 50, 50, Picture2.hdc, 277, 0, SRCCOPY
Picture7.Refresh

End Sub

Private Sub Picture7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'BitBlt Picture7.hdc, 0, 0, 50, 50, Picture2.hdc, 248.5, 0, SRCCOPY
Picture7.Refresh

End Sub

Private Sub Picture8_Click()
On Error GoTo poop
cd1.Filter = "Music(*.mp3;*.wav)|*.mp3;*.wav"
cd1.ShowOpen
If cd1.Filename = "" Then Exit Sub
am1.Filename = cd1.Filename
poop:
Exit Sub
End Sub

Private Sub Picture8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'BitBlt Picture8.hdc, 0, 0, 50, 50, Picture2.hdc, 339, 0, SRCCOPY
Picture8.Refresh
End Sub

Private Sub Picture8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'BitBlt Picture8.hdc, 0, 0, 50, 50, Picture2.hdc, 309, 0, SRCCOPY
Picture8.Refresh
End Sub

Private Sub Picture9_Click()
If loopy = False Then
'BitBlt Picture9.hdc, 0, 0, 50, 50, Picture10.hdc, 31, 0, SRCCOPY
Picture9.Refresh
loopy = True
Else
'BitBlt Picture9.hdc, 0, 0, 50, 50, Picture10.hdc, 0, 0, SRCCOPY
Picture9.Refresh
loopy = False
End If
End Sub

Private Sub Timer1_Timer()
On Error GoTo hell
pimp = am1.CurrentPosition
Dim ex
Picture1.ScaleWidth = 100
ex = Percent(pimp, am1.Duration, Picture1.Width / 100 * 5.3)
Picture1.Line (0, 0)-(Picture1.Width, Picture1.Height), vbBlack, BF
BitBlt Picture1.hdc, ex, 0, Picture11.Width, Form1.Picture11.Height - 1, Picture11.hdc, 0, 0, SRCCOPY


hell:
Exit Sub
End Sub




Private Sub Timer2_Timer()
On Error GoTo hell
If loopy = True Then
If am1.Duration >= am1.SelectionEnd Then
am1.Run
hell:
Exit Sub
End If
End If
End Sub

