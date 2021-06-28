Sub move()

Dim i As Integer, j As Integer, k As Integer

i = 1
For k = 1 To 36
    For j = 1 To 5
            Cells(i, 1).Select
            Selection.Cut
            Cells(k, j).Select
            ActiveSheet.Paste
    i = i + 1
    Next j
Next k

End Sub