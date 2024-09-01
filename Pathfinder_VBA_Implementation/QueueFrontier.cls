VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QueueFrontier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False



Private pFrontier As StackFrontier

Public Property Get frontier() As Variant
    Set frontier = pFrontier
End Property

Public Property Let frontier(ByVal vNewValue As Variant)
    Set pFrontier = vNewValue
End Property

Sub init(stack As StackFrontier)
    Set Me.frontier = stack
End Sub
