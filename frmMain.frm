VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "ModPlayer"
   ClientHeight    =   3300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5220
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   220
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   348
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4320
      MaskColor       =   &H008080FF&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   60
      Width           =   375
   End
   Begin VB.Frame framInfo 
      Caption         =   "Song Info"
      Height          =   2535
      Left            =   1680
      TabIndex        =   5
      Top             =   600
      Width           =   3375
      Begin MSComctlLib.Slider sldSpeed 
         Height          =   615
         Left            =   720
         TabIndex        =   17
         Top             =   1800
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   1085
         _Version        =   393216
         Enabled         =   0   'False
         Max             =   100
         TickStyle       =   2
         TickFrequency   =   5
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   1800
         Top             =   960
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Volume:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   1920
         Width           =   570
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "SS:MSMSMS"
         Height          =   195
         Left            =   600
         TabIndex        =   16
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   600
         TabIndex        =   15
         Top             =   1320
         Width           =   45
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Time:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   390
      End
      Begin VB.Label lblChan 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   960
         TabIndex        =   13
         Top             =   1080
         Width           =   45
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Channels:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   705
      End
      Begin VB.Label lblInst 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1080
         TabIndex        =   11
         Top             =   840
         Width           =   45
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Instruments:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   600
         TabIndex        =   9
         Top             =   600
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Type:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   405
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1080
         TabIndex        =   7
         Top             =   360
         Width           =   45
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Song Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H008080FF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4800
      MaskColor       =   &H008080FF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   60
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop Song"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play Song"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open Song"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   2280
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   $"frmMain.frx":0576
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H80000002&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000002&
      FillColor       =   &H80000002&
      Height          =   3105
      Left            =   -60
      Top             =   360
      Width           =   135
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000002&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000002&
      FillColor       =   &H80000002&
      Height          =   3135
      Left            =   5160
      Top             =   360
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000002&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000002&
      FillColor       =   &H80000002&
      Height          =   105
      Left            =   0
      Top             =   3240
      Width           =   5295
   End
   Begin VB.Image Image1 
      Height          =   225
      Left            =   150
      Picture         =   "frmMain.frx":05FD
      Top             =   120
      Width           =   195
   End
   Begin VB.Label lblCaption 
      BackColor       =   &H80000002&
      Caption         =   "ModPlayer"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   450
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000002&
      BorderColor     =   &H80000002&
      FillColor       =   &H80000002&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   5220
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public declarations of variables holding songs, samples and streams
Dim songHandle As Long
Dim sampleHandle As Long
Dim sampleChannel As Long
Dim streamHandle As Long
Dim streamChannel As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Sub cmdClose_Click()
FSOUND_Close
Unload Me
End Sub

Private Sub cmdOpen_Click()
dlg.Filter = "Extended Module (*.xm)|*.xm|Impulse Tracker Module (*.it)|*.it|Scream Tracker Module  (*.s3m)|*.s3m|ProTracker Module (*.mod)|*.mod"
dlg.ShowOpen

songHandle = FMUSIC_LoadSong(dlg.filename)
lblName.Caption = dlg.FileTitle
lblType.Caption = FMUSIC_GetType(songHandle)
lblInst.Caption = FMUSIC_GetNumInstruments(songHandle)
lblChan.Caption = FMUSIC_GetNumChannels(songHandle)
Timer1.Enabled = True
End Sub

Public Sub MoveWindow(TheHwnd As Long)

    'Drag the form with the mouse
    ReleaseCapture
    SendMessage TheHwnd, &HA1, 2, 0&
End Sub

Private Sub cmdPlay_Click()
Dim result As Boolean
FMUSIC_PlaySong songHandle
sldSpeed.Enabled = True
End Sub

Private Sub cmdStop_Click()
FMUSIC_StopSong (songHandle)
Timer1.Enabled = False
lblTime.Caption = "00:000"
End Sub





Private Sub Command1_Click()
Me.WindowState = 1
End Sub

Private Sub Form_Load()
Dim result As Boolean
FSOUND_Init 44100, 32, 0

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveWindow (Me.hwnd)
End Sub

Private Sub framInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveWindow (Me.hwnd)
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveWindow (Me.hwnd)
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveWindow (Me.hwnd)
End Sub



Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveWindow (Me.hwnd)
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveWindow (Me.hwnd)
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveWindow (Me.hwnd)
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveWindow (Me.hwnd)
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveWindow (Me.hwnd)
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveWindow (Me.hwnd)
End Sub

Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveWindow (Me.hwnd)
End Sub

Private Sub sldSpeed_Change()
If sldSpeed.Enabled = True Then
FMUSIC_SetMasterVolume songHandle, sldSpeed.value
End If
End Sub

Private Sub sldSpeed_Click()
If sldSpeed.Enabled = True Then
FMUSIC_SetMasterVolume songHandle, sldSpeed.value
End If
End Sub

Private Sub sldSpeed_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If sldSpeed.Enabled = True Then
FMUSIC_SetMasterVolume songHandle, sldSpeed.value
End If
End Sub

Private Sub sldSpeed_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If sldSpeed.Enabled = True Then
FMUSIC_SetMasterVolume songHandle, sldSpeed.value
End If
End Sub

Private Sub sldSpeed_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If sldSpeed.Enabled = True Then
FMUSIC_SetMasterVolume songHandle, sldSpeed.value
End If
End Sub

Private Sub sldSpeed_Scroll()
If sldSpeed.Enabled = True Then
FMUSIC_SetMasterVolume songHandle, sldSpeed.value
End If
End Sub

Private Sub Timer1_Timer()
lblTime.Caption = Format(FMUSIC_GetTime(songHandle), "##:###")
End Sub


