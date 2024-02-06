Sub Porównanie()
'
' Porównanie Makro
' Wykonuje całość porównania
'
' Klawisz skrótu: Ctrl+Shift+P
'
    Range("A1:A500").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo _
        :=Array(Array(1, 2), Array(2, 1)), TrailingMinusNumbers:=True
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveCell.FormulaR1C1 = "KK"
    Rows("1:1").Select
    Range("B1").Activate
    ActiveCell.FormulaR1C1 = "IK"
    Range("A1:B500").Select
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "POROWNANIE!R1C1:R500C2", Version:=6).CreatePivotTable TableDestination:= _
        "Arkusz1!R3C1", TableName:="Tabela przestawna1", DefaultVersion:=6
    Sheets("Arkusz1").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("Tabela przestawna1").PivotFields("KK")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Tabela przestawna1").AddDataField ActiveSheet. _
        PivotTables("Tabela przestawna1").PivotFields("IK"), "Suma z IK", xlSum
    Range("A4:B500").Select
    Selection.Copy
    Sheets("POROWNANIE").Select
    Range("D2").Select
    ActiveSheet.Paste
    Columns("A:C").Select
    Range("C1").Activate
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Kod kolektor"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Ilość kolektor"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Ilość hurt"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Różnica"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Nazwa towaru"
    Range("A2:A500").Select
    Selection.NumberFormat = "0.00"
    Selection.NumberFormat = "0.0"
    Selection.NumberFormat = "0"
    Sheets("Arkusz1").Select
    ActiveWindow.SelectedSheets.Delete
    Range("C2").Select
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],BAZA!C1:C3,3,0)"
    Selection.AutoFill Destination:=Range("C2:C500")
    Range("D2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-2]-RC[-1]"
    Range("D2").Select
    Selection.AutoFill Destination:=Range("D2:D500")
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-4],BAZA!C1:C2,2,0)"
    Selection.AutoFill Destination:=Range("E2:E500")
    Columns("A:A").EntireColumn.AutoFit
    Columns("B:B").EntireColumn.AutoFit
    Columns("B:B").EntireColumn.AutoFit
    Columns("C:C").EntireColumn.AutoFit
    Columns("D:D").EntireColumn.AutoFit
    Range("F1").Select
    Sheets("BAZA").Select
    Columns("A:C").Select
    Selection.Copy
    Sheets("POROWNANIE").Select
    Range("F1").Select
    ActiveSheet.Paste
    Range("F:F,A:A").Select
    Range("A1").Activate
    Selection.FormatConditions.AddUniqueValues
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).DupeUnique = xlDuplicate
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub
