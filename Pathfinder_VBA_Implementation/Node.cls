VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Node"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pState As String
Private pParent As node
Private pAction As String

Public Property Get state() As String
    state = pState
End Property

Public Property Let state(ByVal vNewValue As String)
    pState = vNewValue
End Property

Public Property Get parent() As node
    Set parent = pParent
End Property


Public Property Set parent(ByVal vNewValue As node)
    If vNewValue Is Nothing Then
        Dim newParent As New node
        pParent = newParent
    Else
        Set pParent = vNewValue
    End If
End Property

Public Property Get action() As String
    action = pAction
End Property

Public Property Let action(ByVal vNewValue As String)
    pAction = vNewValue
End Property

Sub init(state As String, parent As node, action As String)
    Me.state = state
    Set Me.parent = parent
    Me.action = action
End Sub

Sub init_first_Node(state As String)
    Me.state = state
End Sub
