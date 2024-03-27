Sub AddTrianglesBasedOnColorMultipleShades()
    Dim rng As Range
    Dim cell As Range
    Dim green1 As Long, green2 As Long, green3 As Long, green4 As Long
    Dim red1 As Long, red2 As Long, red3 As Long, red4 As Long
    
    ' Define RGB values for different shades of green and red
    green1 = RGB(1, 102, 94)  ' 99.9% significance - Green
    green2 = RGB(65, 140, 134) ' 99% significance - Green
    green3 = RGB(128, 179, 175) ' 95% significance - Green
    green4 = RGB(192, 217, 215) ' 90% significance - Green

    red1 = RGB(213, 62, 79) ' 99.9% significance - Red
    red2 = RGB(224, 110, 123) ' 99% significance - Red
    red3 = RGB(234, 159, 167) ' 95% significance - Red
    red4 = RGB(245, 207, 211) ' 90% significance - Red

    ' Set the target range where you want to apply the triangles
    Set rng = ThisWorkbook.Sheets("Q39").Range("C6:L17")
    
    For Each cell In rng
        ' Check if the cell's fill is any of the green shades
        If cell.Interior.Color = green1 Or cell.Interior.Color = green2 Or cell.Interior.Color = green3 Or cell.Interior.Color = green4 Then
            cell.NumberFormat = "0%" & ChrW(9650) ' Up-pointing triangle
            cell.Font.Color = RGB(0, 176, 80) ' Green color for the triangle
            cell.Interior.Color = RGB(255, 255, 255) ' Set cell interior to white (remove shading)
        ' Check if the cell's fill is any of the red shades
        ElseIf cell.Interior.Color = red1 Or cell.Interior.Color = red2 Or cell.Interior.Color = red3 Or cell.Interior.Color = red4 Then
            cell.NumberFormat = "0%" & ChrW(9660) ' Down-pointing triangle
            cell.Font.Color = RGB(255, 0, 0) ' Red color for the triangle
            cell.Interior.Color = RGB(255, 255, 255) ' Set cell interior to white (remove shading)
        Else
            cell.Interior.Color = RGB(255, 255, 255) ' Set cell interior to white (remove shading)
            cell.NumberFormat = "0%"
        End If
        ' Center-align the contents of each cell
        cell.HorizontalAlignment = xlCenter
        cell.VerticalAlignment = xlCenter
    Next cell
End Sub