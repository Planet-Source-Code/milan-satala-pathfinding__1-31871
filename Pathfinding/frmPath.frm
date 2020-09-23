VERSION 5.00
Begin VB.Form frmPath 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pathfinding by Milan Satala (thanks to Brayn Stout for his Path Search Demo)"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar scrSpeed 
      Height          =   255
      Left            =   8640
      Max             =   5
      Min             =   1
      TabIndex        =   12
      Top             =   4800
      Value           =   4
      Width           =   3135
   End
   Begin VB.OptionButton optNotOptimized 
      Caption         =   "Not Optimized (Slow but path is the best possible)"
      Height          =   375
      Left            =   8640
      TabIndex        =   11
      Top             =   4080
      Width           =   3135
   End
   Begin VB.OptionButton optOptimized 
      Caption         =   "Optimized (Faster but path is ... not alway the best)"
      Height          =   375
      Left            =   8640
      TabIndex        =   10
      Top             =   3600
      Value           =   -1  'True
      Width           =   3135
   End
   Begin VB.CommandButton cmdGoal 
      Caption         =   "Goal"
      Height          =   615
      Left            =   10680
      TabIndex        =   7
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H8000000A&
      Caption         =   "Start"
      Height          =   615
      Left            =   9720
      TabIndex        =   6
      Top             =   2640
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton cmdTerrain 
      Caption         =   "Terrain"
      Height          =   615
      Left            =   8880
      TabIndex        =   8
      Top             =   2640
      Width           =   855
   End
   Begin VB.Timer tmrSearch 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9960
      Top             =   6240
   End
   Begin VB.ListBox lstTimes 
      Height          =   2400
      Left            =   8640
      TabIndex        =   5
      Top             =   5400
      Width           =   3135
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   615
      Left            =   8640
      TabIndex        =   4
      Top             =   1560
      Width           =   3135
   End
   Begin VB.CommandButton cmdKoniec 
      Caption         =   "End"
      Height          =   615
      Left            =   8640
      TabIndex        =   3
      Top             =   7920
      Width           =   3135
   End
   Begin VB.CommandButton cmdInvert 
      Caption         =   "Invert"
      Height          =   615
      Left            =   8640
      TabIndex        =   2
      Top             =   840
      Width           =   3135
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search !"
      Height          =   615
      Left            =   8640
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
   Begin VB.PictureBox picMap 
      Height          =   8535
      Left            =   0
      ScaleHeight     =   565
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   565
      TabIndex        =   0
      Top             =   0
      Width           =   8535
   End
   Begin VB.Label Label3 
      Caption         =   "Search speed"
      Height          =   255
      Left            =   9720
      TabIndex        =   15
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Slower"
      Height          =   255
      Left            =   8640
      TabIndex        =   13
      Top             =   5040
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Faster"
      Height          =   255
      Left            =   11280
      TabIndex        =   14
      Top             =   5040
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000009&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   9300
      Top             =   3300
      Width           =   135
   End
   Begin VB.Shape Shape3 
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   9180
      Top             =   3300
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      FillColor       =   &H000040C0&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   10200
      Shape           =   3  'Circle
      Top             =   3240
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      FillColor       =   &H00C000C0&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   11040
      Shape           =   3  'Circle
      Top             =   3240
      Width           =   135
   End
   Begin VB.Label lblClick 
      Alignment       =   2  'Center
      Caption         =   "Terrain"
      Height          =   255
      Left            =   8880
      TabIndex        =   9
      Top             =   2280
      Width           =   2655
   End
End
Attribute VB_Name = "frmPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This module isn't commented. Becouse it's used only for editing.
'Go to modPathfinding for pathfinding ...
Option Explicit

Dim BoxSize As Single
Dim WhenClick As Byte

Private Sub cmdClear_Click()
  Dim a As Integer, b As Integer
  
  For a = 0 To MapSize - 1
  For b = 0 To MapSize - 1
   Map(a, b) = 0
  Next b
  Next a
  DrawMap
End Sub

Private Sub cmdInvert_Click()
  Dim a As Integer, b As Integer
  
  For a = 0 To MapSize - 1
  For b = 0 To MapSize - 1
   If Map(a, b) = 0 And (a <> StartX Or b <> StartY) And (a <> GoalX Or b <> GoalY) Then
    Map(a, b) = 1
   Else
    Map(a, b) = 0
   End If
  Next b
  Next a
  DrawMap
End Sub

Private Sub cmdKoniec_Click()
  End
End Sub

Public Sub DrawMap()
  Dim a As Integer, b As Integer
  
  For a = 0 To MapSize - 1
  For b = 0 To MapSize - 1
   If Map(a, b) = 0 Then
    Me.picMap.Line (a * BoxSize, b * BoxSize)-Step(BoxSize, BoxSize), vbWhite, BF
   Else
    Me.picMap.Line (a * BoxSize, b * BoxSize)-Step(BoxSize, BoxSize), vbBlack, BF
   End If
  Next b
  Next a
  
  DrawBox GoalX, GoalY, &HC000C0
  DrawBox StartX, StartY, &H40C0&
End Sub

Private Sub cmdSearch_Click()
  If tmrSearch.Enabled = True Then    'STOP
   tmrSearch.Enabled = False
   Me.cmdSearch.Caption = "New"
  Else
   If picMap.Enabled = False Then     'NEW
    DrawMap
    picMap.Enabled = True
    Me.cmdSearch.Caption = "Search !"
   Else                               'START !
    Me.cmdSearch.Caption = "Stop"
    StartNewSearch
    picMap.Enabled = False
    tmrSearch.Enabled = True
    If frmPath.scrSpeed = 5 Then
     Do
      If tmrSearch.Enabled = False Then Exit Do
      SearchPathCyclus
     Loop
    Else
     SearchPathCyclus
    End If
   End If
  End If
End Sub

Private Sub cmdStart_Click()
  Me.lblClick = "Start"
  WhenClick = 1 'Start
End Sub

Private Sub cmdGoal_Click()
  Me.lblClick = "Goal"
  WhenClick = 2 'Goal
End Sub

Private Sub cmdTerrain_Click()
  Me.lblClick = "Terrain"
  WhenClick = 0 'Terrain
End Sub

Private Sub Form_Load()
  ScaleMode = 3
  BoxSize = Me.picMap.Width / MapSize
End Sub

Private Sub Form_Paint()
  DoEvents
  DrawMap
End Sub

Private Sub picMap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim a As Integer, b As Integer
  
  a = Int(X / BoxSize)
  b = Int(Y / BoxSize)
  
  Select Case WhenClick
   Case 0 'TERRAIN
    If (a <> StartX Or b <> StartY) And (a <> GoalX Or b <> GoalY) Then
     If Button = 1 Then
      Me.picMap.Line (a * BoxSize, b * BoxSize)-Step(BoxSize, BoxSize), vbBlack, BF
      Map(a, b) = 1
     ElseIf Button = 2 Then
      Me.picMap.Line (a * BoxSize, b * BoxSize)-Step(BoxSize, BoxSize), vbWhite, BF
      Map(a, b) = 0
     End If
    End If
    
   Case 1 'START
    DrawBox StartX, StartY, vbWhite
    StartX = a
    StartY = b
    Map(a, b) = 0
    DrawBox a, b, &H40C0&
   Case 2 'GOAL
    DrawBox GoalX, GoalY, vbWhite
    GoalX = a
    GoalY = b
    Map(a, b) = 0
    DrawBox a, b, &HC000C0
    
  End Select
End Sub

Private Sub picMap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim a As Integer, b As Integer
  
  a = Int(X / BoxSize)
  b = Int(Y / BoxSize)
  
  If WhenClick = 0 And (a <> StartX Or b <> StartY) And (a <> GoalX Or b <> GoalY) _
  And a >= 0 And b >= 0 And a < MapSize And b < MapSize Then
   If Button = 1 Then
    Me.picMap.Line (a * BoxSize, b * BoxSize)-Step(BoxSize, BoxSize), vbBlack, BF
    Map(a, b) = 1
   ElseIf Button = 2 Then
    Me.picMap.Line (a * BoxSize, b * BoxSize)-Step(BoxSize, BoxSize), vbWhite, BF
    Map(a, b) = 0
   End If
  End If
End Sub

Sub DrawBox(X As Integer, Y As Integer, Color As Long)
  Me.picMap.Line (X * BoxSize, Y * BoxSize)-Step(BoxSize, BoxSize), Color, BF
End Sub

Private Sub scrSpeed_Change()
  Select Case Me.scrSpeed.Value
   Case 4: Me.tmrSearch.Interval = 1
   Case 3: Me.tmrSearch.Interval = 50
   Case 2: Me.tmrSearch.Interval = 100
   Case 1: Me.tmrSearch.Interval = 300
  End Select
End Sub

Private Sub tmrSearch_Timer()
  SearchPathCyclus
End Sub
