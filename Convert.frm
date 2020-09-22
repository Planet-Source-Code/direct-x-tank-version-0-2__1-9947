VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File converter"
   ClientHeight    =   30
   ClientLeft      =   300
   ClientTop       =   870
   ClientWidth     =   1500
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   30
   ScaleWidth      =   1500
   Begin MSComDlg.CommonDialog cdLoad 
      Left            =   360
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuRun 
         Caption         =   "Select file and convert"
      End
      Begin VB.Menu space1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Colour
  Red As Byte
  Green As Byte
  Blue As Byte
End Type

Private Type Co
 X As Integer
 Y As Integer
 Z As Integer
End Type

Private Type TempWorldDis
  Position(1 To 80) As Co
End Type

Private Type Player
  Pos As Co
  Ang As Co
  Speed As Single
End Type

Private Type FaceType
  Position(1 To 20) As Integer
  Colour As Colour
  Poly As Integer
End Type

Private Type modelsz
  Face(400) As FaceType
  FaceCount As Integer
  Points(800) As Co
  PointCount As Integer
End Type

Dim model(1) As modelsz

Private Sub mnuAbout_Click()


X = X + "About file converter" + vbNewLine + vbNewLine
X = X + "File converter is a program which takes a 3D model, and turns it into a more efficient, but more complicted version of that model." + vbNewLine + vbNewLine
X = X + "This new file can then be quickly be loaded into my DirectX programs." + vbNewLine + vbNewLine
X = X + "The program will create a file with the same name as the file you select, but with the extension *.RAW." + vbNewLine + vbNewLine
X = X + "To create models, either create them in any text editor, or use the simple dos program you should have got with this one. Yeah, yeah, I need a better way of making objects"

MsgBox X, 48, "About File Converter..."



End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuRun_Click()
cdLoad.ShowOpen

Dim TempWorld(400) As TempWorldDis

model(1).FaceCount = 0
model(1).PointCount = 0

'
'On Error GoTo CouldntConvert

Dim Found As Boolean
Npnt = 0
CMod = 1
FalName = cdLoad.FileName

Open FalName For Input As #1
Input #1, model(CMod).FaceCount

For n = 1 To model(CMod).FaceCount
  Input #1, model(CMod).Face(n).Poly
  Input #1, model(CMod).Face(n).Colour.Red
  Input #1, model(CMod).Face(n).Colour.Green
  Input #1, model(CMod).Face(n).Colour.Blue



  For m = 1 To model(CMod).Face(n).Poly
   
   
   Input #1, TempWorld(n).Position(m).X
   Input #1, TempWorld(n).Position(m).Y
   Input #1, TempWorld(n).Position(m).Z
  Next m
Next n

Close #1

Npnt = 0


For X = 1 To 300
   model(CMod).Points(X).X = 0
   model(CMod).Points(X).Y = 0
   model(CMod).Points(X).Z = 0
Next X

For n = 1 To model(CMod).FaceCount
 For m = 1 To model(CMod).Face(n).Poly
  Found = False
  
  For X = 1 To 300
   If TempWorld(n).Position(m).X = model(CMod).Points(X).X Then
    If TempWorld(n).Position(m).Y = model(CMod).Points(X).Y Then
     If TempWorld(n).Position(m).Z = model(CMod).Points(X).Z Then
      Found = True
      model(CMod).Face(n).Position(m) = X
      Exit For
     End If
    End If
   End If
   Next X
   
   If Found = False Then
    Npnt = Npnt + 1
    model(CMod).Face(n).Position(m) = Npnt
    model(CMod).Points(Npnt).X = TempWorld(n).Position(m).X
    model(CMod).Points(Npnt).Y = TempWorld(n).Position(m).Y
    model(CMod).Points(Npnt).Z = TempWorld(n).Position(m).Z
   End If
  Next m
Next n
model(CMod).PointCount = Npnt

MsgBox model(CMod).PointCount



'On Error GoTo CouldntConvert


TotalPoints = model(1).PointCount - 1


TotalEdges = 0
For n = 1 To model(1).FaceCount
 TotalEdges = TotalEdges + model(1).Face(n).Poly + 1
Next n

ReDim VertArray(TotalPoints) As Co
ReDim SideFaces(TotalEdges)

For n = 1 To model(1).PointCount
 VertArray(n - 1).X = (model(1).Points(n).X) / 10
 VertArray(n - 1).Y = (model(1).Points(n).Y) / 10
 VertArray(n - 1).Z = (model(1).Points(n).Z) / 10
Next n

Posish = 0
For n = 1 To model(1).FaceCount
 SideFaces(Posish) = model(1).Face(n).Poly
 Posish = Posish + 1
  
  For m = 1 To model(1).Face(n).Poly
   SideFaces(Posish) = model(1).Face(n).Position(m) - 1
   Posish = Posish + 1
  Next m
Next n
Close


NewFileName = Mid(cdLoad.FileName, 1, Len(cdLoad.FileName) - 4) + ".Raw"

Open NewFileName For Output As #9
Print #9, TotalPoints
Print #9, TotalEdges

For n = 0 To TotalPoints
  Print #9, VertArray(n).X
  Print #9, VertArray(n).Y
  Print #9, VertArray(n).Z
Next n

For n = 0 To TotalEdges
 Print #9, SideFaces(n)
Next n
Close


CouldntConvert:
End Sub
