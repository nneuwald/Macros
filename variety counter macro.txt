Sub Varietycounter()
    Dim rng As Range
    Dim uniqueNames As Collection
    Dim i As Long, j As Long
    Dim counter As Long
    Dim startRow As Long
    Dim currentSubject As String
    
    ' Set the selected range
    Set rng = Selection
    
    ' Initialize variables
    Set uniqueNames = New Collection
    counter = 0
    startRow = rng.Rows(1).Row
    
    ' Loop through each row in the selected range
    For i = 1 To rng.Rows.Count + 1
        ' Handle the subject change or last row
        If i > rng.Rows.Count Or rng.Cells(i, 1).Value <> rng.Cells(startRow - rng.Rows(1).Row + 1, 1).Value Then
            ' Output the count of unique names to the last column + 1 of your selection
            rng.Cells(startRow - rng.Rows(1).Row + 1, rng.Columns.Count + 1).Value = counter
            
            ' Reset variables for the next subject
            Set uniqueNames = New Collection
            counter = 0
            startRow = i + rng.Rows(1).Row - 1
        End If
        
        ' If within the data range, try to add the name
        If i <= rng.Rows.Count Then
            currentSubject = rng.Cells(i, 2).Value
            On Error Resume Next
            uniqueNames.Add currentSubject, CStr(currentSubject)
            
            ' If no error occurred, it's a unique name, so increment the counter
            If Err.Number = 0 Then
                counter = counter + 1
            End If
            On Error GoTo 0
        End If
    Next i
End Sub
