
Public Function DateiName() As String
Dim f As Office.FileDialog
Set f = Application.FileDialog(msoFileDialogFilePicker)

f.Show

If f.SelectedItems.Count > 0 Then ' Prüfen auf 'Abbrechen'-Button
    DateiName = f.SelectedItems(1)
End If
End Function

Public Function Zeilenende(tabelle As String) As String
Dim loErste As Long, loLetzte As Long

With Worksheets(tabelle)
    If .FilterMode Then .ShowAllData
    loErste = .Cells(.Rows.Count, 9).End(xlUp).Offset(1, 0).Row
    loLetzte = .Cells(.Rows.Count, 1).End(xlUp).Row
End With

Zeilenende = loLetzte

End Function
