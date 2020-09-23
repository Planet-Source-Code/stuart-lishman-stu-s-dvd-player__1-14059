VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14670
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0442
   ScaleHeight     =   2205
   ScaleWidth      =   14670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5520
      ScaleHeight     =   495
      ScaleWidth      =   2295
      TabIndex        =   1
      Top             =   1080
      Width           =   2295
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Index           =   1
         Left            =   1440
         TabIndex        =   6
         ToolTipText     =   "Current Time Position"
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Index           =   0
         Left            =   600
         TabIndex        =   5
         ToolTipText     =   "Current Time Position"
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Index           =   2
         Left            =   1680
         TabIndex        =   4
         ToolTipText     =   "Current Time Position"
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Index           =   1
         Left            =   840
         TabIndex        =   3
         ToolTipText     =   "Current Time Position"
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Index           =   0
         Left            =   0
         TabIndex        =   2
         ToolTipText     =   "Current Time Position"
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   360
      Top             =   240
   End
   Begin VB.PictureBox Progres 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   4800
      ScaleHeight     =   105
      ScaleWidth      =   5265
      TabIndex        =   0
      ToolTipText     =   "Play Position Progress Indicator"
      Top             =   1800
      Width           =   5295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5550
      TabIndex        =   10
      ToolTipText     =   "Total Play Time"
      Top             =   820
      Width           =   975
   End
   Begin VB.Image Image22 
      Height          =   180
      Left            =   2610
      Picture         =   "Form1.frx":69822
      Top             =   1110
      Width           =   330
   End
   Begin VB.Image Image21 
      Height          =   180
      Left            =   2160
      Picture         =   "Form1.frx":69B96
      Top             =   1080
      Width           =   330
   End
   Begin VB.Label Subs 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   9600
      TabIndex        =   9
      ToolTipText     =   "Subtitles On (Click to Select Language)"
      Top             =   840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image19 
      Height          =   150
      Left            =   10780
      Picture         =   "Form1.frx":69F0A
      ToolTipText     =   "Subtitles On / Off"
      Top             =   640
      Width           =   270
   End
   Begin VB.Image Image18 
      Height          =   165
      Left            =   13060
      Picture         =   "Form1.frx":6A17E
      ToolTipText     =   "Previous Chapter"
      Top             =   695
      Width           =   525
   End
   Begin VB.Image Image17 
      Height          =   165
      Left            =   13650
      Picture         =   "Form1.frx":6A666
      ToolTipText     =   "Next Chapter"
      Top             =   720
      Width           =   585
   End
   Begin VB.Image Image16 
      Height          =   750
      Left            =   4800
      Picture         =   "Form1.frx":6ABD2
      ToolTipText     =   "Cue Reverse"
      Top             =   960
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image Image15 
      Height          =   750
      Left            =   4800
      Picture         =   "Form1.frx":6B1CC
      ToolTipText     =   "Cue Forward"
      Top             =   960
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Chap 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   8040
      TabIndex        =   8
      ToolTipText     =   "Current Chapter"
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Speed 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "x1"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   9840
      TabIndex        =   7
      ToolTipText     =   "Play Speed"
      Top             =   1080
      Width           =   255
   End
   Begin VB.Image Image14 
      Height          =   750
      Left            =   4800
      Picture         =   "Form1.frx":6B7C6
      ToolTipText     =   "Stop"
      Top             =   960
      Width           =   750
   End
   Begin VB.Image Image13 
      Height          =   750
      Left            =   4800
      Picture         =   "Form1.frx":6BDC0
      ToolTipText     =   "Pause"
      Top             =   960
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image Image12 
      Height          =   750
      Left            =   4800
      Picture         =   "Form1.frx":6C3BA
      ToolTipText     =   "Play"
      Top             =   960
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image Image11 
      Height          =   255
      Left            =   9880
      Picture         =   "Form1.frx":6C9B4
      Stretch         =   -1  'True
      ToolTipText     =   "Mute On"
      Top             =   850
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image10 
      Height          =   255
      Left            =   9880
      Picture         =   "Form1.frx":6D1F6
      Stretch         =   -1  'True
      ToolTipText     =   "Mute Off"
      Top             =   850
      Width           =   255
   End
   Begin VB.Image Image9 
      Height          =   150
      Left            =   11110
      Picture         =   "Form1.frx":6DA38
      ToolTipText     =   "Full Screen Mode"
      Top             =   1140
      Width           =   270
   End
   Begin VB.Image Image8 
      Height          =   150
      Left            =   10500
      Picture         =   "Form1.frx":6DCAC
      ToolTipText     =   "Mute Audio"
      Top             =   1140
      Width           =   270
   End
   Begin VB.Image Image7 
      Height          =   195
      Left            =   12240
      Picture         =   "Form1.frx":6DF20
      ToolTipText     =   "Eject"
      Top             =   385
      Width           =   525
   End
   Begin VB.Image Image6 
      Height          =   165
      Left            =   13080
      Picture         =   "Form1.frx":6E4E0
      ToolTipText     =   "Cue Reverse"
      Top             =   400
      Width           =   525
   End
   Begin VB.Image Image5 
      Height          =   165
      Left            =   13655
      Picture         =   "Form1.frx":6E9C8
      ToolTipText     =   "Cue Forward"
      Top             =   400
      Width           =   585
   End
   Begin VB.Image Image4 
      Height          =   315
      Left            =   12190
      Picture         =   "Form1.frx":6EF34
      ToolTipText     =   "Stop"
      Top             =   1040
      Width           =   600
   End
   Begin VB.Image Image3 
      Height          =   315
      Left            =   13050
      Picture         =   "Form1.frx":6F950
      ToolTipText     =   "Pause"
      Top             =   1040
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   13680
      Picture         =   "Form1.frx":703C0
      ToolTipText     =   "Play"
      Top             =   1050
      Width           =   585
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   450
      Picture         =   "Form1.frx":70D64
      ToolTipText     =   "On / Off Switch"
      Top             =   990
      Width           =   645
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   4750
      Top             =   840
      Width           =   5370
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XS, XY, MuteOn As Boolean
Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = Screen.Height - Me.Height
Form2.Show
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
XS = X: XY = Y
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then GoTo ef
xd = XS - X: yd = YS - Y: XS = X: YS = Y
Me.Left = Me.Left - xd
Me.Top = Me.Top - yd
ef:
End Sub
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Top = 970
Image1.Left = 470
End Sub
Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Image1.Top = 990
Image1.Left = 450
ans = MsgBox("Are you sure you want to EXIT", vbYesNo, "Exit")
If ans = 6 Then End
End Sub

Private Sub Image17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image17.Top = 720
Image17.Left = 13660
End Sub
Private Sub Image17_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Image17.Top = 720
Image17.Left = 13650
Form2.DVD1.PlayNextChapter
End Sub

Private Sub Image18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Image18.Left = 13050
Image18.Top = 695
Form2.DVD1.PlayPrevChapter
End Sub
Private Sub Image18_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image18.Left = 13060
Image18.Top = 695
End Sub

Private Sub Image19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image19.Top = 660
Image19.Left = 10765
End Sub

Private Sub Image19_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Image19.Top = 640
Image19.Left = 10780
If Form2.DVD1.SubpictureOn = True Then Form2.DVD1.SubpictureOn = False: Subs.Visible = False: GoTo nnd
If Form2.DVD1.SubpictureOn = False Then Form2.DVD1.SubpictureOn = True: Subs.Visible = True: GoTo nnd
nnd:
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Image2.Top = 1040
Image2.Left = 13680
End Sub
Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Image2.Top = 1050
Image2.Left = 13700
Speed.Caption = "x1"
Form2.DVD1.Play
Image12.Visible = True
Image13.Visible = False
Image14.Visible = False
Image15.Visible = False
Image16.Visible = False
End Sub

Private Sub Image21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image21.Top = 1090
Image21.Left = 2160
End Sub
Private Sub Image21_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Image21.Top = 1110
Image21.Left = 2160
Form2.DVD1.SaveBookmark
End Sub

Private Sub Image22_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image22.Top = 1090
Image22.Left = 2610
End Sub
Private Sub Image22_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Image22.Top = 1110
Image22.Left = 2610
Form2.DVD1.RestoreBookmark
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Top = 1020
Image3.Left = 13030
End Sub
Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Image3.Top = 1040
Image3.Left = 13050
Form2.DVD1.Pause
Image12.Visible = False
Image13.Visible = True
Image14.Visible = False
Image15.Visible = False
Image16.Visible = False
End Sub
Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Image4.Top = 1020
Image4.Left = 12180
End Sub
Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Image4.Top = 1040
Image4.Left = 12190
Form2.DVD1.Stop
Image12.Visible = False
Image13.Visible = False
Image14.Visible = True
Image15.Visible = False
Image16.Visible = False
End Sub
Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image5.Top = 400
Image5.Left = 13665
End Sub
Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Image5.Top = 400
Image5.Left = 13655
If Speed.Caption = "x1" Then Speed.Caption = "x2": Form2.DVD1.PlayForwards (2): GoTo eeen
If Speed.Caption = "x2" Then Speed.Caption = "x4": Form2.DVD1.PlayForwards (4): GoTo eeen
If Speed.Caption = "x4" Then Speed.Caption = "x8": Form2.DVD1.PlayForwards (8): GoTo eeen
If Speed.Caption = "x8" Then Speed.Caption = "x1": Form2.DVD1.PlayForwards (1): GoTo eeen
eeen:
If Speed.Caption = "x1" Then Image12.Visible = True: Image15.Visible = False
Image13.Visible = False
Image14.Visible = False
If Speed.Caption <> "x1" Then Image12.Visible = False:: Image15.Visible = True
Image16.Visible = False
End Sub
Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image6.Top = 400
Image6.Left = 13060
End Sub
Private Sub Image6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Image6.Top = 400
Image6.Left = 13080
If Speed.Caption = "x1" Then Speed.Caption = "x8": Form2.DVD1.PlayBackwards (8): GoTo eeens
If Speed.Caption = "x8" Then Speed.Caption = "x4": Form2.DVD1.PlayBackwards (4): GoTo eeens
If Speed.Caption = "x4" Then Speed.Caption = "x2": Form2.DVD1.PlayBackwards (2): GoTo eeens
If Speed.Caption = "x2" Then Speed.Caption = "x1": Form2.DVD1.PlayBackwards (1): GoTo eeens
eeens:
Image12.Visible = False
Image13.Visible = False
Image14.Visible = False
Image15.Visible = False
Image16.Visible = True
End Sub
Private Sub Image7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image7.Top = 375
Image7.Left = 12230
End Sub
Private Sub Image7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Image7.Top = 385
Image7.Left = 12240
Form2.DVD1.Eject
End Sub
Private Sub Image8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image8.Top = 1130
Image8.Left = 10485
End Sub
Private Sub Image8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Image8.Top = 1140
Image8.Left = 10500
If MuteOn = False Then Form2.DVD1.Mute = True: MuteOn = True: GoTo enn
If MuteOn = True Then Form2.DVD1.Mute = False: MuteOn = False: GoTo enn
enn:
If MuteOn = True Then Image11.Visible = True Else Image11.Visible = False
End Sub
Private Sub Image9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image9.Top = 1130
Image9.Left = 11095
End Sub
Private Sub Image9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Image9.Top = 1140
Image9.Left = 11110
Form2.DVD1.FullScreenMode = True
End Sub

Private Sub Subs_Click()
Form3.Show
End Sub
Private Sub Timer1_Timer()
Dim num(10)
On Error Resume Next
tt = Form2.DVD1.CurrentTime
tot = Form2.DVD1.TotalTitleTime
tt = Str(tt)
tot = Str(tot)
If Len(tt) < 10 Then totaltm = 1: posnow = 1:  GoTo hj
If Len(tot) < 10 Then totaltm = 1: posnow = 1:  GoTo hj
num(1) = Mid$(tt, 1, 2)
num(2) = Mid$(tt, 4, 2)
num(3) = Mid$(tt, 7, 2)
num(4) = Mid$(tt, 10, 2)
num(5) = Mid$(tot, 1, 2)
num(6) = Mid$(tot, 4, 2)
num(7) = Mid$(tot, 7, 2)
num(8) = Mid$(tot, 10, 2)
Label1(0).Caption = num(1)
Label1(1).Caption = num(2)
Label1(2).Caption = num(3)
totaltm = ((num(5) * 60) * 60) + (num(6) * 60) + (num(7))
posnow = ((num(1) * 60) * 60) + (num(2) * 60) + (num(3))
hj:
Progres.ScaleWidth = totaltm
Progres.Line (0, 0)-(posnow, 150), RGB(200, 0, 0), BF
Progres.Line (posnow + 1, 0)-(Progres.ScaleWidth, 150), RGB(0, 0, 0), BF
If Form2.DVD1.CurrentTitle <> 1 Then Chap.Caption = "Title " & (Form2.DVD1.CurrentTitle) - 1
If Form2.DVD1.CurrentTitle = 1 Then Chap.Caption = "Chapter " & Form2.DVD1.CurrentChapter
SubCount = Form2.DVD1.SubpictureStreamsAvailable
For a = 1 To SubCount
SubTitleLang(a) = Form2.DVD1.GetSubpictureLanguage(a)
Next a
Label3.Caption = Left$(Form2.DVD1.TotalTitleTime, 8)
End Sub


