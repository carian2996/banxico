
Sub seriesSWAPS()

Dim macro As String, derivado As String, MaxCol As String
Dim i As Integer, LastCol As Integer
Dim e As Double, d As Double, pesta As String

macro = "CalculaContraparte.xlsm"
derivado = "swaps_contra_09072014.xlsx"
pesta = "swaps_contra_09072014"

'======================== PATOS ========================

Windows(derivado).Activate
Sheets(pesta).Select
Range("B1").Select
Selection.End(xlDown).Select 'Ubico el ?ltimo registro de la fila
e = ActiveWindow.RangeSelection.Row
Range("B2:B" & e).Select
Selection.Copy

Range("B" & e + 1).Select
ActiveSheet.Paste
Selection.RemoveDuplicates Columns:=1, Header:=xlNo

Range("B" & e + 1).Select
Selection.End(xlDown).Select 'Ubico el ?ltimo registro de la fila de fechas
d = ActiveWindow.RangeSelection.Row

'ESCRIBE PATITO
Range("L" & e + 1).Select
ActiveCell.FormulaR1C1 = "PATITO"
Selection.Copy
Range("L" & e + 1 & ":L" & d).Select
ActiveSheet.Paste

'ESCRIBE CEROS EN IMPORTE
Range("J" & e + 1).Select
ActiveCell.FormulaR1C1 = "0"
Selection.Copy
Range("J" & e + 1 & ":J" & d).Select
ActiveSheet.Paste

MsgBox ("Actualiza la tabla")

'======================== o ======================== BANCOS ======================== o ======================== o
'======================== SECCION I ========================
Windows(derivado).Activate

Sheets("tabla").Select
    
    ActiveSheet.PivotTables("tabladim").PivotFields("TIPO_INST").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("tabladim").PivotFields("TIPO_INST")
        .PivotItems("BM_BD").Visible = True
        .PivotItems("CB").Visible = False
        .PivotItems("(blank)").Visible = True
    End With
    
    ActiveSheet.PivotTables("tabladim").PivotFields("SECCION").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("tabladim").PivotFields("SECCION")
        .PivotItems("I").Visible = True
        .PivotItems("II").Visible = False
        .PivotItems("III").Visible = False
        .PivotItems("IV").Visible = False
        .PivotItems("PATITO").Visible = True
        .PivotItems("(blank)").Visible = True
    End With
    ActiveSheet.PivotTables("tabladim").PivotFields("SECCION").EnableMultiplePageItems = True
    
    ActiveSheet.PivotTables("tabladim").PivotFields("TIPO_INST").EnableMultiplePageItems = True
    With ActiveSheet.PivotTables("tabladim")
        .ColumnGrand = False
        .RowGrand = False
    End With

Range("A5:A5").Select
Selection.End(xlDown).Select
e = ActiveWindow.RangeSelection.Row

    Range("A5:A" & e - 1).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
   
    Sheets.Add
    ActiveSheet.Name = "resultado"
    Range("B3").Select
    ActiveSheet.Paste
    Range("C2").Value = "IC - SEC I"

MaxCol = Cells.SpecialCells(xlLastCell).Column

'======================== SECCION II - IV ========================
Sheets("tabla").Select
    
    ActiveSheet.PivotTables("tabladim").PivotFields("TIPO_INST").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("tabladim").PivotFields("TIPO_INST")
        .PivotItems("BM_BD").Visible = True
        .PivotItems("CB").Visible = False
        .PivotItems("(blank)").Visible = True
    End With
    
    ActiveSheet.PivotTables("tabladim").PivotFields("SECCION").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("tabladim").PivotFields("SECCION")
        .PivotItems("I").Visible = False
        .PivotItems("II").Visible = True
        .PivotItems("III").Visible = True
        .PivotItems("IV").Visible = True
        .PivotItems("PATITO").Visible = True
        .PivotItems("(blank)").Visible = True
    End With
    ActiveSheet.PivotTables("tabladim").PivotFields("SECCION").EnableMultiplePageItems = True
    
    ActiveSheet.PivotTables("tabladim").PivotFields("TIPO_INST").EnableMultiplePageItems = True
    With ActiveSheet.PivotTables("tabladim")
        .ColumnGrand = False
        .RowGrand = False
    End With

Range("A5:A5").Select
Selection.End(xlDown).Select
e = ActiveWindow.RangeSelection.Row

    Range("A5:A" & e).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    
    Sheets("resultado").Select
    Cells(3, MaxCol + 2).Select
    ActiveSheet.Paste
    Cells(2, MaxCol + 3).Value = "IC - SEC II IV"
    
   Cells(4, MaxCol + 2).Select
MaxCol = Cells.SpecialCells(xlLastCell).Column
    
'======================== o ======================== CASAS DE BOLSA ======================== o ======================== o
'======================== SECCION I ========================
Sheets("tabla").Select
    
    ActiveSheet.PivotTables("tabladim").PivotFields("TIPO_INST").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("tabladim").PivotFields("TIPO_INST")
        .PivotItems("BM_BD").Visible = False
        .PivotItems("CB").Visible = True
        .PivotItems("(blank)").Visible = True
    End With
    
    ActiveSheet.PivotTables("tabladim").PivotFields("SECCION").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("tabladim").PivotFields("SECCION")
        .PivotItems("I").Visible = True
        .PivotItems("II").Visible = False
        .PivotItems("III").Visible = False
        .PivotItems("IV").Visible = False
        .PivotItems("PATITO").Visible = True
        .PivotItems("(blank)").Visible = True
    End With
    ActiveSheet.PivotTables("tabladim").PivotFields("SECCION").EnableMultiplePageItems = True
    
    ActiveSheet.PivotTables("tabladim").PivotFields("TIPO_INST").EnableMultiplePageItems = True
    With ActiveSheet.PivotTables("tabladim")
        .ColumnGrand = False
        .RowGrand = False
    End With

Range("A5:A5").Select
Selection.End(xlDown).Select
e = ActiveWindow.RangeSelection.Row

    Range("A5:A" & e).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    
    Sheets("resultado").Select
    Cells(3, MaxCol + 2).Select
    ActiveSheet.Paste
    Cells(2, MaxCol + 3).Value = "CB - SEC I"
    
   Cells(4, MaxCol + 2).Select
MaxCol = Cells.SpecialCells(xlLastCell).Column

'======================== SECCION II - IV ========================
Sheets("tabla").Select
    
    ActiveSheet.PivotTables("tabladim").PivotFields("TIPO_INST").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("tabladim").PivotFields("TIPO_INST")
        .PivotItems("BM_BD").Visible = False
        .PivotItems("CB").Visible = True
        .PivotItems("(blank)").Visible = True
    End With
    
    ActiveSheet.PivotTables("tabladim").PivotFields("SECCION").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("tabladim").PivotFields("SECCION")
        .PivotItems("I").Visible = False
        .PivotItems("II").Visible = True
        .PivotItems("III").Visible = True
        .PivotItems("IV").Visible = True
        .PivotItems("PATITO").Visible = True
        .PivotItems("(blank)").Visible = True
    End With
    ActiveSheet.PivotTables("tabladim").PivotFields("SECCION").EnableMultiplePageItems = True
    
    ActiveSheet.PivotTables("tabladim").PivotFields("TIPO_INST").EnableMultiplePageItems = True
    With ActiveSheet.PivotTables("tabladim")
        .ColumnGrand = False
        .RowGrand = False
    End With

Range("A5:A5").Select
Selection.End(xlDown).Select
e = ActiveWindow.RangeSelection.Row

    Range("A5:A" & e).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    
    Sheets("resultado").Select
    Cells(3, MaxCol + 2).Select
    ActiveSheet.Paste
    Cells(2, MaxCol + 3).Value = "CB - SEC II IV"
    
   Cells(4, MaxCol + 2).Select
MaxCol = Cells.SpecialCells(xlLastCell).Column

'======================== o ======================== SERIES ======================== o ======================== o
Rows("4").Select
Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
Rows("3").Copy
Range("A4").Select
ActiveSheet.Paste
Range("C3 ").Select
i = 1

'======================== IC - I ========================
Do While Workbooks(derivado).Sheets("resultado").Cells(3, 2 + i) <> ""
 Cells(3, 2 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("resultado").Cells(3, 2 + i).Value, Workbooks(macro).Sheets("IC").Range("$C$2:$D$30"), 2, True)
i = i + 1
Loop

Cells(3, 2 + i).EntireColumn.Delete
Cells(3, 2 + i).EntireColumn.Delete
'======================== IC - II IV ========================
Do While Workbooks(derivado).Sheets("resultado").Cells(3, 2 + i) <> ""
 Cells(3, 2 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("resultado").Cells(3, 2 + i).Value, Workbooks(macro).Sheets("IC").Range("$C$2:$D$30"), 2, True)
i = i + 1
Loop

Cells(3, 2 + i).EntireColumn.Delete
Cells(3, 2 + i).EntireColumn.Delete
'======================== CB - I  ========================
Do While Workbooks(derivado).Sheets("resultado").Cells(3, 2 + i) <> ""
 Cells(3, 2 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("resultado").Cells(3, 2 + i).Value, Workbooks(macro).Sheets("CB").Range("$B$2:$C$30"), 2, False)
i = i + 1
Loop

Cells(3, 2 + i).EntireColumn.Delete
Cells(3, 2 + i).EntireColumn.Delete
'======================== CB - II IV ========================
Do While Workbooks(derivado).Sheets("resultado").Cells(3, 2 + i) <> ""
 Cells(3, 2 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("resultado").Cells(3, 2 + i).Value, Workbooks(macro).Sheets("CB").Range("$B$2:$C$30"), 2, False)
i = i + 1
Loop

Range("B4").Select
Selection.End(xlDown).Select
e = ActiveWindow.RangeSelection.Row
Range("A4:A" & e).FormulaR1C1 = "=TEXT(RC[1], ""dd/mm/aaaa"")"

Range("A4:A" & e).Copy
Range("B4").PasteSpecial Paste:=xlPasteValues
Range("A4:A" & e).ClearContents


MsgBox ("Listo")

End Sub
