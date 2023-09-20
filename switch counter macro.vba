Sub SwitchCounter()
    Dim LastRow As Long
    Dim i As Long
    Dim CurrentSubject As Variant
    Dim PreviousSubject As Variant
    Dim Changes As Long
    Dim OutputColumn As Long
    
    LastRow = Selection.Rows.Count
    OutputColumn = Selection.Column + Selection.Columns.Count
    
    ' Initialize the current and previous subject variables
    CurrentSubject = Cells(Selection.Row, Selection.Column).Value
    PreviousSubject = CurrentSubject
    Changes = 0
    
    For i = 2 To LastRow
        ' Update the current and previous subject variables and count changes
        CurrentSubject = Cells(Selection.Row + i - 1, Selection.Column).Value
        If CurrentSubject <> PreviousSubject Then
            ' Output the total changes for the previous subject to the new column to the right
            Cells(Selection.Row + i - 2, OutputColumn).Value = Changes
            Changes = 0
            PreviousSubject = CurrentSubject
        ElseIf Cells(Selection.Row + i - 1, Selection.Column + 1).Value <> Cells(Selection.Row + i - 2, Selection.Column + 1).Value Then
            ' Increment the changes count for the current subject
            Changes = Changes + 1
        End If
    Next i
    
    ' Output the total changes for the last subject to the new column to the right
    Cells(Selection.Row + LastRow - 1, OutputColumn).Value = Changes
End Sub
