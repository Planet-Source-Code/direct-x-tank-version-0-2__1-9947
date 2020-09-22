VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Game 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Direct 3D"
   ClientHeight    =   5295
   ClientLeft      =   3135
   ClientTop       =   1650
   ClientWidth     =   7275
   ControlBox      =   0   'False
   FillColor       =   &H00404040&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   7275
   Begin VB.CommandButton Quit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   5520
      TabIndex        =   2
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton mnuLoad 
      Caption         =   "Load 40"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   4560
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog cdLoad 
      Left            =   1920
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox View 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000008&
      Height          =   4455
      Left            =   0
      ScaleHeight     =   293
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   485
      TabIndex        =   0
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "Game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim unit As Co
Dim VertArray() As D3DVECTOR
Dim SideFaces() As Long
Dim XMouse, YMouse As Integer

Private Sub Quit_Click()
End
End Sub

Private Sub cLoadModels()
On Error GoTo Knackered
Open cdLoad.FileName For Input As #1
Input #1, TotalPoints
Input #1, TotalEdges

ReDim VertArray(TotalPoints)
ReDim SideFaces(TotalEdges)

Xof = Rnd * 1000
Zof = Rnd * 1000

For Tp = 0 To TotalPoints


  Input #1, VertArray(Tp).X
  Input #1, VertArray(Tp).Y
  Input #1, VertArray(Tp).Z

  VertArray(Tp).X = VertArray(Tp).X + Xof
  VertArray(Tp).Y = VertArray(Tp).Y + Yof
  VertArray(Tp).Z = VertArray(Tp).Z + Zof
  
Next Tp

For Te = 0 To TotalEdges
 Input #1, SideFaces(Te)
Next Te

Close

AddShape TotalPoints + 1, VertArray, SideFaces
Knackered:
End Sub

Private Sub Form_Load()
    Show
    InitRM View
    InitScene
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CleanUp
End Sub

Private Sub View_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   XMouse = X
   YMouse = Y
End Sub

Private Sub mnuLoad_Click()
 cdLoad.ShowOpen
 For n = 1 To 40
 Game.Caption = 40 - n
  cLoadModels
 Next n
 PlayGame
End Sub

Sub PlayGame()
Pie = (22 / 7) * 18.2

Do
KeepTime = Int(Timer * 10)
DoEvents
Game.Caption = Counter


If XMouse < 50 Then
 RotatePlayer (-4 / Pie)
End If

If XMouse > View.ScaleWidth - 50 Then
 RotatePlayer (4 / Pie)
End If


MovePlayerForward (Speed)

If YMouse < 50 Then
 Speed = Speed + 1
 If Speed > 15 Then Speed = 15
ElseIf YMouse > View.ScaleHeight - 50 Then
 Speed = Speed - 1
 If Speed < -15 Then Speed = -15
Else
 Speed = Speed * 0.8
 If Speed < 1 Then Speed = 0
End If


RenderScene



Counter = Counter + 1
Do: Loop While Int(Timer * 10) = KeepTime
Loop
End Sub

