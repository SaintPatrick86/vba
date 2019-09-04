Public Function DateiName() As String
Dim f As Office.FileDialog
Set f = Application.FileDialog(msoFileDialogFilePicker)

f.Show

If f.SelectedItems.Count > 0 Then ' Prüfen auf 'Abbrechen'-Button
    DateiName = f.SelectedItems(1)
End If
End Function
