VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form PlayList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Play List"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   Icon            =   "PlayList.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   Picture         =   "PlayList.frx":0CCA
   ScaleHeight     =   4305
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture5 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   3960
      Picture         =   "PlayList.frx":71C8
      ScaleHeight     =   330
      ScaleWidth      =   1095
      TabIndex        =   5
      Top             =   3260
      Width           =   1095
   End
   Begin VB.PictureBox Picture4 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   3960
      Picture         =   "PlayList.frx":88AE
      ScaleHeight     =   330
      ScaleWidth      =   1095
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   3960
      Picture         =   "PlayList.frx":9F67
      ScaleHeight     =   330
      ScaleWidth      =   1095
      TabIndex        =   3
      Top             =   2680
      Width           =   1095
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   3960
      Picture         =   "PlayList.frx":B608
      ScaleHeight     =   330
      ScaleWidth      =   1095
      TabIndex        =   2
      Top             =   2100
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   2280
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   3960
      Picture         =   "PlayList.frx":CD29
      ScaleHeight     =   330
      ScaleWidth      =   1095
      TabIndex        =   1
      Top             =   1520
      Width           =   1095
   End
   Begin VB.ListBox lstfavs 
      Height          =   2790
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   3615
   End
End
Attribute VB_Name = "PlayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'App.Path + "/" + "playlist.txt"

Private Sub Form_Load()
FormOnTop PlayList
End Sub

Private Sub Label5_Click()

End Sub

Private Sub lstfavs_DblClick()
Form1.am1.Filename = lstfavs.Text
End Sub

Private Sub Picture1_Click()
On Error Resume Next
cd1.Filter = "Music(*.mp3;*.wav)|*.mp3;*.wav"
cd1.ShowOpen
If cd1.Filename = "" Then Exit Sub
lstfavs.AddItem cd1.Filename
End Sub

Private Sub Picture2_Click()
On Error Resume Next
lstfavs.RemoveItem lstfavs.ListIndex
End Sub

Private Sub Picture3_Click()
On Error Resume Next
cd1.Filter = "Jagis Playlist(*.jag)|*.jag"
cd1.ShowOpen
If cd1.Filename = "" Then Exit Sub
Call WriteList(lstfavs, cd1.Filename)

End Sub

Private Sub Picture4_Click()
On Error Resume Next
cd1.Filter = "Jagis Playlist(*.jag)|*.jag"
cd1.ShowOpen
If cd1.Filename = "" Then Exit Sub
Call ReadList(lstfavs, cd1.Filename, True)
End Sub

Private Sub Picture5_Click()
lstfavs.Clear
End Sub

Private Sub Picture6_Click()

End Sub
