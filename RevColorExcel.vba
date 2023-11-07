/*
VBA code which can reverse the color of text and background and set all border in white color in Excel
*/

Sub ReverseColorsAndSetWhiteBorders()
    Dim rng As Range
    Dim cell As Range
    Dim tempColor As Long
    
    ' Set the range to the currently selected cells
    Set rng = Selection
    
    ' Loop through each cell in the range
    For Each cell In rng
        With cell
            ' Swap background color with font color and vice versa
            tempColor = .Interior.Color
            .Interior.Color = .Font.Color
            .Font.Color = tempColor
            
            ' Set all borders to white
            .Borders.Color = RGB(255, 255, 255) ' RGB(255, 255, 255) represents white
        End With
    Next cell
End Sub
