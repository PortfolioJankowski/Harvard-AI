Attribute VB_Name = "Program"
Public wb As Workbook
Public ws As Worksheet
Public directions As Collection
Public states As Collection

Sub PublicStaticVoidMain()
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets(1)
    Set directions = New Collection
    Set states = New Collection
    Dim maze As New maze
    maze.init
    maze.solve
    
End Sub
