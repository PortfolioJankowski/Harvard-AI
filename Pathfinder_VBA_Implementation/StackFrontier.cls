VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StackFrontier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pFrontier As Collection

Public Property Get frontier() As Collection
    Set frontier = pFrontier
End Property

Public Property Let frontier(ByVal vNewValue As Collection)
    Set pFrontier = vNewValue
End Property

Public Sub add(node As node)
    Me.frontier.add node
End Sub

Public Function contains_state(state As String) As Boolean
    Dim element As node
    For Each element In Me.frontier
        If element.state = state Then
            contains_state = True
            Exit Function
        End If
    Next element
    contains_state = False
End Function

Public Function isEmpty() As Boolean
    isEmpty = (Me.frontier.Count = 0)
End Function

Public Function remove() As node
    If Me.frontier.Count = 0 Then
        MsgBox "Empty frontier"
    Else
        ' For Queue (FIFO) behavior, remove the first item
        Set remove = Me.frontier(1)
        Me.frontier.remove 1
    End If
End Function

Private Sub Class_Initialize()
    Set pFrontier = New Collection
End Sub

