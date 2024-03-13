Attribute VB_Name = "Alien"
Public Const NUM_ALIENS As Integer = 15
Public Const NUM_ROWS As Integer = 3
Public aliensX() As Integer
Public aliensY() As Integer
Public alienRanges() As Range

Sub InitializeAliens()
    Dim i As Integer
    ReDim aliensX(1 To NUM_ALIENS)
    ReDim aliensY(1 To NUM_ALIENS)
    ReDim alienRanges(1 To NUM_ALIENS)
    
    ' Set initial positions for each alien
    For i = 1 To NUM_ROWS
        For j = 1 To 5
            aliensX((i - 1) * 5 + j) = j * 6
            aliensY((i - 1) * 5 + j) = i * 6 ' Adjust the spacing vertically
        Next j
    Next i
End Sub

Sub DrawAliens()
    Dim i As Integer
    Dim alienCells As Range
    Dim cell As Range
    Dim color As Integer
    
    ' Define colors for each line of aliens
    Dim lineColors(1 To 3) As Long
    lineColors(1) = RGB(255, 0, 0) ' Red
    lineColors(2) = RGB(0, 255, 0) ' Green
    lineColors(3) = RGB(0, 0, 255) ' Blue
    
    For i = 1 To NUM_ALIENS
        ' Determine the color index for the current line for 5 aliens
        color = ((i - 1) \ 5) + 1
        
        ' Define the cells that make up the alien shape
        Set alienCells = Union(Range(Cells(aliensY(i), aliensX(i) - 1), Cells(aliensY(i), aliensX(i) + 1)), _
                               Range(Cells(aliensY(i) + 1, aliensX(i) - 2), Cells(aliensY(i) + 1, aliensX(i) + 2)), _
                               Range(Cells(aliensY(i) + 2, aliensX(i) - 2), Cells(aliensY(i) + 2, aliensX(i) + 2)), _
                               Range(Cells(aliensY(i) + 3, aliensX(i) - 1), Cells(aliensY(i) + 3, aliensX(i) + 1)), _
                               Range(Cells(aliensY(i) + 4, aliensX(i)), Cells(aliensY(i) + 4, aliensX(i))))
        
        ' Draw the alien shape with the corresponding color
        For Each cell In alienCells
            cell.Interior.color = lineColors(color)
        Next cell
    Next i
End Sub


