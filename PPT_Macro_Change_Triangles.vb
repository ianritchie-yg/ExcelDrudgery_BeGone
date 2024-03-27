Sub ChangeTriangleColorsInTables()
    Dim sld As slide
    Dim shp As shape
    Dim tbl As Table
    Dim rw As Row
    Dim cl As Cell
    Dim char As textRange
    Dim i As Integer
    
    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.HasTable Then
                Set tbl = shp.Table
                For Each rw In tbl.Rows
                    For Each cl In rw.Cells
                        If cl.shape.HasTextFrame Then
                            If cl.shape.TextFrame.HasText Then
                                For i = 1 To cl.shape.TextFrame.textRange.Characters.Count
                                    Set char = cl.shape.TextFrame.textRange.Characters(i, 1)
                                    If char.Text = ChrW(&H25B2) Then
                                        char.Font.Color.RGB = RGB(0, 128, 0) ' Green
                                    ElseIf char.Text = ChrW(&H25BC) Then
                                        char.Font.Color.RGB = RGB(128, 0, 0) ' Red
                                    End If
                                Next i
                            End If
                        End If
                    Next cl
                Next rw
            End If
        Next shp
    Next sld
End Sub

