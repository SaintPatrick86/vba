Sub import()

Dim csvpfad As String
Dim delimiter As String
Dim blatt As String
Dim zeile As String
Dim beginn As String

delimiter = ";"

blatt = "Tabelle1"

csvpfad = PMIT_Functions.DateiName

zeile = PMIT_Functions.Zeilenende(blatt)

beginn = "B" & zeile + 1

Debug.Print beginn

Debug.Print csvpfad

Call ImportCSVFromFolder(csvpfad, delimiter, blatt, beginn)

End Sub






Private Sub ImportCSVFromFolder(csvpfad As String, strCSVDelimiter As String, blatt As String, beginn As String)

    Dim wsTemp As Worksheet, wsTarget As Worksheet, curCell As Range, fso As Object, f As Object, akt As Integer

    'Legt das CSV-Trennzeichen für die Dateien fest

    Set fso = CreateObject("Scripting.Filesystemobject")

    Application.DisplayAlerts = False

    Application.ScreenUpdating = False

    'Zielarbeitsblatt für die importierten Daten

    Set wsTarget = ActiveWorkbook.Worksheets(blatt)

    'temporäres Arbeitsblatt für den Import der Daten erstellen

    Set wsTemp = ActiveWorkbook.Worksheets("TEMP")

    'Inhalt des Zusammenfassungsblattes löschen

    'wsTarget.UsedRange.Clear

    'Startausgabezelle festlegen

    Set curCell = wsTarget.Range(beginn)

            'Temporäres Sheet löschen

   wsTemp.Activate
    
    
    wsTemp.UsedRange.Clear
    
    
            'CSV-Daten in Temporäres Sheet importieren

    With ActiveSheet.QueryTables.Add(Connection:="TEXT;" & csvpfad, Destination:=wsTemp.Range("$A$1"))

          .Name = "import"
          .FieldNames = True
          .AdjustColumnWidth = True
          .RefreshPeriod = 0
          .TextFilePlatform = xlWindows
          .TextFileStartRow = 5
          .TextFileParseType = xlDelimited
          .TextFileTextQualifier = xlTextQualifierDoubleQuote
          .TextFileOtherDelimiter = strCSVDelimiter
          .Refresh BackgroundQuery:=False
          .Delete
    End With

    With wsTemp

                .UsedRange.SpecialCells(xlCellTypeVisible).Copy curCell

    End With


            'Ausgabezeile eins nach unten schieben
    
    akt = PMIT_Functions.Zeilenende(blatt)
    
    
    Set curCell = wsTarget.Cells(akt + 1, 2)


    'Spalten anpassen



    wsTarget.Columns.AutoFit


    Application.DisplayAlerts = True



    Application.ScreenUpdating = True



    MsgBox "Vorgang beendet!", vbInformation



    Set fso = Nothing

End Sub

