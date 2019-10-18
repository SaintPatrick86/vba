Public Function IstDurchgestrichen(rng As Range) As Boolean
    Dim i As Long
    With rng(1)
        For i = 1 To .Characters.Count
            If .Characters(i, 1).Font.Strikethrough Then
                IstDurchgestrichen = True
                Exit For
            End If
        Next i
    End With
End Function
