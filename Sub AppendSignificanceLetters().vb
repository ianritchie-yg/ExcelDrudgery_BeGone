Sub AddSignificanceLetters()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim percentColumn As String
    Dim significanceColumnArray() As Variant
    Dim fontSize As Integer
    Dim fillColor As String
    Dim fontColor As String
    Dim alignment As Integer
    
    ' Set the active worksheet
    Set ws = ActiveSheet
    
    ' Find the last row with data in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Get user input for the percentage column letter
    percentColumn = InputBox("Enter the column letter for percentages (e.g., C):")
    
    ' Get user input for the significance column numbers
    significanceColumns = InputBox("Enter the significance column numbers separated by commas (e.g., 3, 5):")
    significanceColumnArray = Split(significanceColumns, ",")
    
    ' Get user input for font size
    fontSize = InputBox("Enter the font size for the added letters:")
    
    ' Get user input for fill color (hex color code)
    fillColor = InputBox("Enter the fill color as a hex color code (e.g., FFFF00 for yellow):")
    
    ' Get user input for font color (hex color code)
    fontColor = InputBox("Enter the font color as a hex color code (e.g., FF0000 for red):")
    
    ' Get user input for alignment (1 = Left, 2 = Center, 3 = Right)
    alignment = CInt(InputBox("Enter the alignment (1 = Left, 2 = Center, 3 = Right):"))
    
    ' Define the range for the percentages
    Set percentRange = ws.Columns(percentColumn)
    
    ' Loop through each percentage cell
    For Each cell In percentRange.SpecialCells(xlCellTypeConstants, xlNumbers)
        ' Loop through each significance column
        For i = 0 To UBound(significanceColumnArray)
            Set significanceRange = cell.Offset(0, CLng(significanceColumnArray(i)) - CLng(Asc(UCase(percentColumn)) - 64))
            ' Check if the significance cell is not empty
            If Not IsEmpty(significanceRange.Value) Then
                ' Add the significance letter to the percentage cell
                cell.Value = cell.Value & Chr(10) & significanceRange.Value
                cell.WrapText = True
                ' Set the font size, fill color, font color, and alignment for the added letter
                cell.Characters(Start:=Len(cell.Value) - Len(significanceRange.Value) + 1, Length:=Len(significanceRange.Value)).Font.Size = fontSize
                cell.Characters(Start:=Len(cell.Value) - Len(significanceRange.Value) + 1, Length:=Len(significanceRange.Value)).Font.Color = RGB(CLng("&H" & Right(fontColor, 2)), CLng("&H" & Mid(fontColor, 3, 2)), CLng("&H" & Left(fontColor, 2)))
                cell.Interior.Color = RGB(CLng("&H" & Right(fillColor, 2)), CLng("&H" & Mid(fillColor, 3, 2)), CLng("&H" & Left(fillColor, 2)))
                ' Depending on alignment input, set cell alignment
                Select Case alignment
                    Case 1
                        cell.HorizontalAlignment = xlLeft
                    Case 2
                        cell.HorizontalAlignment = xlCenter
                    Case 3
                        cell.HorizontalAlignment = xlRight
                End Select
                ' Clear the original significance cell
                significanceRange.ClearContents
            End If
        Next i
    Next cell
    
    ' Center-align the contents of each cell
    ws.Range(percentColumn & "2:" & percentColumn & lastRow).HorizontalAlignment = xlCenter
    ws.Range(percentColumn & "2:" & percentColumn & lastRow).VerticalAlignment = xlCenter
End Sub