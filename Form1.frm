VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Игра Shifter"
   ClientHeight    =   5955
   ClientLeft      =   2340
   ClientTop       =   3960
   ClientWidth     =   9150
   DrawMode        =   16  'Merge Pen
   BeginProperty Font 
      Name            =   "Arial Black"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   9150
   Begin VB.Frame GamePanel 
      Caption         =   "Игровое поле"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   50
      TabIndex        =   1
      Top             =   0
      Width           =   5295
      Begin VB.CommandButton ResetMap 
         Caption         =   "Заново"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   5280
         Width           =   5000
      End
      Begin VB.PictureBox MainPlace 
         AutoRedraw      =   -1  'True
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5000
         Left            =   120
         ScaleHeight     =   4935
         ScaleWidth      =   4935
         TabIndex        =   7
         Top             =   240
         Width           =   5000
         Begin VB.Shape Cell 
            BorderColor     =   &H80000007&
            FillColor       =   &H0000FF00&
            FillStyle       =   0  'Заливка
            Height          =   615
            Left            =   1320
            Top             =   1440
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Shape BgCell 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Заливка
            Height          =   735
            Left            =   840
            Top             =   1080
            Visible         =   0   'False
            Width           =   855
         End
      End
   End
   Begin VB.Frame ControlPanel 
      Caption         =   "Настройки"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   5370
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin VB.PictureBox MapPlace 
         AutoRedraw      =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2500
         Left            =   1080
         ScaleHeight     =   2445
         ScaleWidth      =   2445
         TabIndex        =   8
         Top             =   240
         Width           =   2500
      End
      Begin VB.CommandButton SetMap 
         Caption         =   "Начать"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         TabIndex        =   5
         Top             =   2880
         Width           =   2535
      End
      Begin VB.OptionButton MapR3 
         Caption         =   "Узор 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1455
      End
      Begin VB.OptionButton MapR2 
         Caption         =   "Узор 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton MapR1 
         Caption         =   "Узор 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Победа!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   36
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5640
      TabIndex        =   9
      Top             =   4080
      Visible         =   0   'False
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const GamePlaceSizeX = 4
Const GamePlaceSizeY = 4

Private Type TGamePlace
  Data(GamePlaceSizeX, GamePlaceSizeY) As Byte
End Type

Dim GamePlace As TGamePlace, CurrentGameMap As TGamePlace, Map1 As TGamePlace, Map2 As TGamePlace, Map3 As TGamePlace, CurrentGamePlace As TGamePlace
Dim CurrentPartCount As Byte
Dim MousePushed As Boolean, Moved As Boolean
Dim PosX As Integer, PosY As Integer, NewPosX As Integer, NewPosY As Integer

Private Sub InitArr(ByRef Arr As TGamePlace)
Dim i As Byte, j As Byte
  For i = 0 To GamePlaceSizeX Step 1
    For j = 0 To GamePlaceSizeY Step 1
      Arr.Data(i, j) = 0
    Next
  Next
End Sub

Private Sub FillMap(ByRef CGP As TGamePlace, Map As Byte)
Dim i As Byte, j As Byte
  CurrentPartCount = 0
  Call InitArr(CGP)
  Select Case Map
    Case 1
      For i = 0 To GamePlaceSizeX Step 1
        For j = 0 To GamePlaceSizeY Step 1
          If (i And 1) Then
            If (j And 1) Then
              CGP.Data(i, j) = 1
            End If
          Else
            If Not (j And 1) Then
              CGP.Data(i, j) = 1
            End If
          End If
        Next
      Next
    Case 2
      For i = 0 To GamePlaceSizeX Step 1
        CGP.Data(i, i) = 1
        CGP.Data(i, GamePlaceSizeX - i) = 1
      Next
    Case 3
      For i = 0 To GamePlaceSizeX Step 1
        For j = 0 To GamePlaceSizeY Step 1
          If Not (i And 1) Then
            If (j And 1) Then CGP.Data(i, j) = 1 Else CGP.Data(i, j) = 0
          End If
        Next
      Next
  End Select
  For i = 0 To GamePlaceSizeX Step 1
    For j = 0 To GamePlaceSizeY Step 1
      CurrentPartCount = CurrentPartCount + CGP.Data(i, j)
    Next
  Next
End Sub

Private Sub PaintCell(Place As PictureBox, CellX As Integer, CellY As Integer, CellColor As ColorConstants, MyScale As Byte)
  Place.Line (CellX, CellY)-(CellX + (500 * MyScale), CellY + (500 * MyScale)), CellColor, BF
  Place.Line (CellX, CellY)-(CellX + (500 * MyScale), CellY + (500 * MyScale)), , B
End Sub

Private Sub PaintPlace(Place As PictureBox, MyScale As Byte, ByRef GP As TGamePlace)
Dim i As Byte, j As Byte
  Place.ForeColor = vbBlack
  For i = 0 To GamePlaceSizeX Step 1
    For j = 0 To GamePlaceSizeY Step 1
      If (GP.Data(i, j) = 0) Then Call PaintCell(Place, i * 500 * MyScale, j * 500 * MyScale, vbRed, MyScale) Else Call PaintCell(Place, i * 500 * MyScale, j * 500 * MyScale, vbGreen, MyScale)
    Next
  Next
  Place.ZOrder (0)
End Sub

Private Function CompareArray(ByRef arr1 As TGamePlace, ByRef arr2 As TGamePlace) As Boolean
Dim i As Byte, j As Byte
Dim Equal As Boolean
  Equal = True
  i = 0
  While ((Equal) And (i <= GamePlaceSizeX))
    For j = 0 To GamePlaceSizeY - 1 Step 1
      If (arr1.Data(i, j) <> arr2.Data(i, j)) Then
        Equal = False
      End If
    Next
    i = i + 1
  Wend
  CompareArray = Equal
End Function

Private Sub FillGamePlace(ByRef GP As TGamePlace)
Dim i As Byte, j As Byte, TempPartCount As Byte
  TempPartCount = 0
  Call InitArr(GP)
  i = 0
  While (TempPartCount < CurrentPartCount)
    j = 0
    While ((TempPartCount < CurrentPartCount) And (j <= GamePlaceSizeY))
      GP.Data(i, j) = 1
      TempPartCount = TempPartCount + 1
      j = j + 1
    Wend
    i = i + 1
  Wend
End Sub

Private Sub SetMap_Click()
  MainPlace.Enabled = True
  Call FillGamePlace(GamePlace)
  Call PaintPlace(MainPlace, 2, GamePlace)
  ResetMap.Enabled = True
  Label1.Visible = False
End Sub

Private Sub MapR1_Click()
  Call FillMap(CurrentGamePlace, 1)
  Call PaintPlace(MapPlace, 1, CurrentGamePlace)
End Sub

Private Sub MapR2_Click()
  Call FillMap(CurrentGamePlace, 2)
  Call PaintPlace(MapPlace, 1, CurrentGamePlace)
End Sub

Private Sub MapR3_Click()
  Call FillMap(CurrentGamePlace, 3)
  Call PaintPlace(MapPlace, 1, CurrentGamePlace)
End Sub

Private Sub ResetMap_Click()
  Call SetMap_Click
End Sub

Private Sub Form_Load()
  Randomize
  Cell.Width = 1000
  Cell.Height = 1000
  BgCell.Width = Cell.Width
  BgCell.Height = Cell.Height
End Sub

Private Sub MainPlace_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  PosX = X \ 1000
  PosY = Y \ 1000
  Moved = False
  If GamePlace.Data(PosX, PosY) = 1 Then
    BgCell.Left = PosX * 1000
    BgCell.Top = PosY * 1000
    Cell.Left = X - (Cell.Width \ 2)
    Cell.Top = Y - (Cell.Height \ 2)
    MousePushed = True
    Cell.Visible = True
    BgCell.Visible = True
    Moved = True
  End If
End Sub

Private Sub MainPlace_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim TempPosX As Integer, TempPosY As Integer
  If MousePushed Then
    TempPosX = X \ 1000
    TempPosY = Y \ 1000
    Cell.Left = X - (Cell.Width \ 2)
    Cell.Top = Y - (Cell.Height \ 2)
  End If
End Sub

Private Sub MainPlace_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  NewPosX = X \ 1000
  NewPosY = Y \ 1000
  MousePushed = False
  Cell.Visible = False
  BgCell.Visible = False
  If ((GamePlace.Data(NewPosX, NewPosY) = 0) And Moved) Then
    If ((Abs(NewPosX - PosX) < 2) And (Abs(NewPosY - PosY) < 2)) Then
      If Not ((PosX <> NewPosX) And (PosY <> NewPosY)) Then
        GamePlace.Data(NewPosX, NewPosY) = 1
        GamePlace.Data(PosX, PosY) = 0
        Call PaintPlace(MainPlace, 2, GamePlace)
        If CompareArray(GamePlace, CurrentGamePlace) Then Label1.Visible = True
      End If
    End If
  End If
End Sub
