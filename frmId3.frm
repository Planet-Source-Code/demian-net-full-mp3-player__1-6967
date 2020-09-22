VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmId3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ID3 Tagger"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6675
   Icon            =   "frmId3.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmId3.frx":0CCA
   ScaleHeight     =   3690
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   2520
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6000
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "*.mp3"
      DialogTitle     =   "Mp3 Filez"
      Filter          =   "Mp3 filez (*.mp3)|*.mp3"
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1440
      TabIndex        =   13
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   4920
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   4
      Top             =   2880
      Width           =   6375
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   3720
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   2
      Top             =   1800
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3720
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   1
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   0
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Genre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5460
      TabIndex        =   11
      Top             =   1680
      Width           =   555
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2880
      TabIndex        =   10
      Top             =   2520
      Width           =   885
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4080
      TabIndex        =   9
      Top             =   1560
      Width           =   435
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Album"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1320
      TabIndex        =   8
      Top             =   1560
      Width           =   555
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Artist"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4920
      TabIndex        =   7
      Top             =   840
      Width           =   465
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1440
      TabIndex        =   6
      Top             =   840
      Width           =   405
   End
End
Attribute VB_Name = "frmId3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
CommonDialog1.ShowOpen
GetId3 CommonDialog1.Filename           ' Get the filename
Text1 = RTrim(id3Info.Title)            ' since the fields in the type are
Text2 = RTrim(id3Info.Artist)                  ' fixed lenght, we use Rtrim to cut the
Text3 = RTrim(id3Info.Album)                   ' trailing bytes
Text4 = RTrim(id3Info.sYear)
Text5 = RTrim(id3Info.Comments)
Combo1.ListIndex = id3Info.Genre        ' fill in all the correct info.
Command2.Enabled = True
End Sub

Private Sub Command2_Click()
        
id3Info.Title = Text1           ' just filling in the information into the type
id3Info.Artist = Text2
id3Info.Album = Text3
id3Info.sYear = Text4
id3Info.Comments = Text5
'id3Info.Genre = Combo1.ListIndex
On Error GoTo ErrHandle             ' If the file is writeprotected
SaveId3 CommonDialog1.Filename, id3Info     ' Calling the Saveid3 function
Exit Sub


ErrHandle:
If Err.Number = 75 Then
MsgBox "File is Write Protected"
Else
MsgBox Err.Description
End If
End Sub

Private Sub Form_Load()
FormOnTop frmId3

On Error Resume Next
GenreArray = Split(sGenreMatrix, "|")   ' we fill the array with the Genre's
For I = LBound(GenreArray) To UBound(GenreArray)
Combo1.AddItem GenreArray(I)        ' now fill the Combobox with the array, and voila, the code you
                                    ' you recieve form the Genre part of the Type, represents the combobox Listindex =)
Next
'Dim Position(0 To 147) As Long
'Dim Start As Long
'Start = 1
'Position(0) = 1
'On Error Resume Next        ' it creates an error once sGenreMatrix runs out of "|"
'For I = 1 To 147             ' number of Genre's in sGenreMatrix
'pt = InStr(start,sGenreMatrix, "|")position ](i) = pt + 1
'Start = pt + 1
'Next
'For I = 0 To 147
'X = (Mid$(sGenreMatrix, Position(I), Position(I + 1) - Position(I) - 1))
'Combo1.AddItem X

'Next

'GetId3 Form1.am1.Filename           ' Get the filename
'Text1 = RTrim(id3Info.Title)            ' since the fields in the type are
'Text2 = RTrim(id3Info.Artist)                  ' fixed lenght, we use Rtrim to cut the
'Text3 = RTrim(id3Info.Album)                   ' trailing bytes
'Text4 = RTrim(id3Info.sYear)
'Text5 = RTrim(id3Info.Comments)
'Text6 = RTrim(id3Info.Genre)
'Combo1.ListIndex = id3Info.Genre        ' fill in all the correct info.
'Command2.Enabled = True

End Sub

