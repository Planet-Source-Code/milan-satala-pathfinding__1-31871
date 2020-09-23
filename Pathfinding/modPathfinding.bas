Attribute VB_Name = "modPathfinding"
Option Explicit

Dim CyclusCount As Integer
Dim NodeCount As Integer

Public Type NodeType
  X As Integer
  Y As Integer
  Points As Single
  Parent As Integer
End Type

Dim Nodes(MapSize * MapSize) As NodeType

Dim Field() As Byte
Dim Order As Collection
Public StartX As Integer, StartY As Integer, GoalX As Integer, GoalY As Integer

Public Sub StartNewSearch()
  Set Order = New Collection
  Order.Add 1 'order is order of search
  
  NodeCount = 1 'Create first node in start position
  Nodes(1).X = StartX
  Nodes(1).Y = StartY
  
  ReDim Field(MapSize, MapSize) 'resize field
  
  SearchPathCyclus 'do first cyclus
End Sub

Public Sub SearchPathCyclus()
  Dim a As Integer, b As Integer
  Dim Parent As Integer, Points As Single, Position As Integer
  Dim Item As Variant
  
  If Order.Count = 0 Then 'Goal can't be reached from start position :-(
   frmPath.cmdSearch.Caption = "New"
   frmPath.tmrSearch.Enabled = False
   MsgBox "Goal can't be reached from start position :-("
   Exit Sub
  End If
  
  Parent = Order.Item(1) 'get number of first node
  Order.Remove 1 'remove first node from order so second node is now first ...
  
  If frmPath.scrSpeed < 5 Then frmPath.DrawBox Nodes(Parent).X, Nodes(Parent).Y, vbBlue 'nothing important (graphic)
  
  For a = Nodes(Parent).X - 1 To Nodes(Parent).X + 1 'test all fields around it
  For b = Nodes(Parent).Y - 1 To Nodes(Parent).Y + 1
   If a >= 0 And a < MapSize And b >= 0 And b < MapSize And (a <> Nodes(Parent).X Or b <> Nodes(Parent).Y) Then  'make sure that it is still in field
    
    If a = GoalX And b = GoalY Then 'if true then we have reached the goal - VICTORY !!!
     frmPath.cmdSearch.Caption = "New" 'nothing important
     frmPath.tmrSearch.Enabled = False
     DrawPath Parent
     Exit Sub
    End If
    
    'if Map(?,?) = 0 then there is no barrier on that field
    'if Field(?,?) = 0 then there isn't no node on this field yet
    If Field(a, b) = 0 And Map(a, b) = 0 Then 'true - there is no barrier there and there isn't no node on this field yet
     
     Field(a, b) = 1 'well, now there is some node on it so ...
      
     'The diference between optimized and not optimized is that the optimized uses
     'the nodes that are closest to goal
     If frmPath.optOptimized.Value = True Then 'OPTIMIZED
      Points = 0.2 * CyclusCount + Sqr((a - GoalX) ^ 2 + (b - GoalY) ^ 2)
     Else                                      'NOT OPTIMIZED
      Points = CyclusCount
     End If
     
     NodeCount = NodeCount + 1 'create new node and add points to it
       
     Nodes(NodeCount).Parent = Parent 'remember its parent so we can find out the path later
     Nodes(NodeCount).X = a
     Nodes(NodeCount).Y = b
     Nodes(NodeCount).Points = Points
     
     If frmPath.scrSpeed < 5 Then frmPath.DrawBox a, b, vbRed 'nothing important (graphic)
     
     Position = 0
     For Each Item In Order
      Position = Position + 1
      
      If Points < Nodes(Item).Points Then 'put it into order (less points the better)
       Order.Add NodeCount, , Position
       GoTo Next_
      End If
     Next
     Order.Add NodeCount 'it dont have more (less) points than any other node so put it into bottom of order
     
    End If
   End If
Next_:
  Next b
  Next a
   
  CyclusCount = CyclusCount + 1
End Sub

Public Sub DrawPath(Index As Integer)
  Dim PathDistandce As Integer
  Dim Last As Integer
  
  Last = Index 'draw all nodes - first draws node that came to goal then its parent
  'then parent of goal node parent ... until it comes to the first one (start)
  Do Until Last = 1
   PathDistandce = PathDistandce + 1
   frmPath.DrawBox Nodes(Last).X, Nodes(Last).Y, vbGreen
   
   Last = Nodes(Last).Parent
  Loop
  
  
  TestCount = TestCount + 1 'nothing important
  
  If frmPath.optOptimized.Value = True Then
   frmPath.lstTimes.AddItem "Test #" & TestCount & ": Dist: " & PathDistandce & " Nodes: " & NodeCount & " (Optimized)"
  Else
   frmPath.lstTimes.AddItem "Test #" & TestCount & ": Dist: " & PathDistandce & " Nodes: " & NodeCount & " (Not Optim.)"
  End If
End Sub
