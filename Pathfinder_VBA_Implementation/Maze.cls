VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Maze"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private pBottomRow As Integer
Private pRightCol As Integer
Private pWalls As Collection
Private pStart As String
Private pEnding As String
Private pSolution As Variant
Private pNum_Explored As Integer
Public Property Get Ending() As Variant
    Ending = pEnding
End Property
Public Property Let Ending(ByVal vNewValue As Variant)
    pEnding = vNewValue
End Property
Public Property Get Start() As Variant
    Start = pStart
End Property

Public Property Let Start(ByVal vNewValue As Variant)
    pStart = vNewValue
End Property
Public Property Get Bottom() As Variant
    Bottom = pBottomRow
End Property

Public Property Let Bottom(ByVal vNewValue As Variant)
    pBottomRow = vNewValue
End Property
Public Property Get Right() As Variant
    Right = pRightCol
End Property

Public Property Let Right(ByVal vNewValue As Variant)
    pRightCol = vNewValue
End Property

Public Property Get num_Explored() As Integer
     num_Explored = pNum_Explored
End Property

Public Property Let num_Explored(ByVal vNewValue As Integer)
     pNum_Explored = vNewValue
End Property
Public Property Get solution() As Variant
     Set solution = pSolution
End Property

Public Property Let solution(ByVal vNewValue As Variant)
     Set pSolution = vNewValue
End Property
Public Property Get walls() As Variant
     Set walls = pWalls
End Property

Public Property Let walls(ByVal vNewValue As Variant)
    Set pWalls = vNewValue
End Property

Sub init()
    Dim contents As Range
    Set contents = ThisWorkbook.Worksheets(1).Range("A1").CurrentRegion
    
:
:
    Me.Bottom = ws.cells(ws.Rows.Count, 1).End(xlUp).row
    Me.Right = ws.cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    Set pWalls = New Collection
    
    Dim i As Integer, j As Integer
    
    For i = 1 To pBottomRow
        Dim row As Variant
        ReDim row(1 To pRightCol)
        For j = 1 To pRightCol
            If ws.cells(i, j) = "A" Then
                pStart = ws.cells(i, j).Address
                row(j) = False
            ElseIf ws.cells(i, j) = "B" Then
                pEnding = ws.cells(i, j).Address
                row(j) = False
            ElseIf ws.cells(i, j) = "" Then
                row(j) = False
            Else
                row(j) = True
            End If
        Next j
        Me.walls.add row
    Next i
    
    Call printMaze
End Sub

Sub printMaze()
    Dim i As Integer
    Debug.Print "Start: " & Me.Start
    Debug.Print "End: " & Me.Ending
    Debug.Print "                   "
    Dim element As Variant
    For Each element In Me.walls
        Dim newLine As String: newLine = ""
        For i = 1 To Me.Right
            If element(i) = True Then
                newLine = newLine & "X"
            Else
                newLine = newLine & " "
            End If
        Next i
       
        Debug.Print newLine
    Next element
End Sub
'Checking if there is any possibility to move
Function neighbors(state) As Collection
    Dim row As Integer, col As Integer
    row = ws.Range(state).row
    col = ws.Range(state).Column
    
    Dim candidates As Variant
    candidates = Array( _
        Array("up", Array(row - 1, col)), _
        Array("down", Array(row + 1, col)), _
        Array("left", Array(row, col - 1)), _
        Array("right", Array(row, col + 1)) _
    )
    
    ' Randomize the order of candidates
    Dim i As Integer, j As Integer, temp As Variant
    For i = LBound(candidates) To UBound(candidates)
        j = Int((UBound(candidates) - LBound(candidates) + 1) * Rnd + LBound(candidates))
        temp = candidates(i)
        candidates(i) = candidates(j)
        candidates(j) = temp
    Next i
    
    Dim result As New Collection
    Dim action As String, r As Integer, c As Integer
    For i = LBound(candidates) To UBound(candidates)
        action = candidates(i)(0)
        r = candidates(i)(1)(0)
        c = candidates(i)(1)(1)
        
        If ifCellIsNeighbouringAndIsNotWall(r, c, row, col) Then
            result.add Array(action, Array(r, c))
        End If
    Next i
    
    Set neighbors = result
End Function


Function ifCellIsNeighbouringAndIsNotWall(r As Integer, c As Integer, row As Integer, col As Integer) As Boolean
    ' Check if the target cell is within bounds
    If r >= 1 And r <= Me.Bottom And c >= 1 And c <= Me.Right Then
        ' Check if the target cell is not a wall
        If Not Me.walls(r)(c) = True Then
'            ' Check if the target cell is a direct neighbor (top, down, left, right)
'            If (Abs(row - r) = 1 And col = c) Or (Abs(col - c) = 1 And row = r) Then
                ifCellIsNeighbouringAndIsNotWall = True
                Exit Function
'            End If
        End If
    End If
    ifCellIsNeighbouringAndIsNotWall = False
End Function

Sub solve()
    ' Keep track of number of states explored
    Me.num_Explored = 0

    ' Initialize frontier to just the starting position
    Dim Start As New node
    Start.init_first_Node Me.Start
    states.add Me.Start
    Dim frontier As New StackFrontier
    frontier.add Start

    ' Initialize an empty explored set
    Dim explored As New Collection

    ' Keep looping until solution found
    Do
        ' If nothing left in frontier, then no path
        If frontier.isEmpty Then
            MsgBox "No solution"
            Exit Sub
        End If

        ' Choose a node from the frontier
        Dim node As node
        Set node = frontier.remove
        Me.num_Explored = Me.num_Explored + 1

        ' If node is the goal, then we have a solution
        If node.state = Me.Ending Then
            MsgBox "Found the way"
            Call PrintResult
            Exit Sub
        End If

        ' Mark node as explored
        explored.add node.state

        ' Add neighbors to frontier
        Dim neighbors As Collection
        Set neighbors = Me.neighbors(node.state)
        Dim neighbor As Variant
        For Each neighbor In neighbors
            Dim action As String
            Dim state As Variant
            action = neighbor(0)
            state = neighbor(1)
            
            Dim neiRow As Integer: neiRow = state(0)
            Dim neiCol As Integer: neiCol = state(1)
            
            If Not frontier.contains_state(ws.cells(neiRow, neiCol).Address) And Not IsInCollection(explored, state) Then
                Dim child As New node
                Dim actualState As String: actualState = GetCellAddress(state)
                child.init actualState, node, action
                states.add actualState
                directions.add action
                frontier.add child
            End If
        Next neighbor
    Loop
End Sub
Sub PrintResult()
    Dim wsResult As Worksheet
    Set wsResult = wb.Worksheets(1)
    wsResult.Range("p2:p2000").ClearContents
    wsResult.Range("q2:q2000").ClearContents
    Dim i As Integer
    For i = 1 To states.Count
        If i <> states.Count Then
            wsResult.cells(i + 1, "P").Value = states(i)
            wsResult.cells(i + 1, "Q").Value = directions(i)
        Else
            wsResult.cells(i + 1, "P").Value = states(i)

        End If
    Next i
End Sub
Function GetCellAddress(state As Variant) As String
    Dim row As Integer: row = state(0)
    Dim col As Integer: col = state(1)
    GetCellAddress = cells(row, col).Address
End Function

Function IsInCollection(col As Collection, val As Variant) As Boolean
    Dim item As Variant
    Dim actualCell As Range
    
    
    
    
    
    For Each item In col
        Set actualCell = ws.Range(item)
        If actualCell.row = val(0) And actualCell.Column = val(1) Then
            IsInCollection = True
            Exit Function
        End If
    Next item
    IsInCollection = False
End Function



Private Sub Class_Initialize()
    Dim newArr As Variant
    pSolution = newArr
    
End Sub
