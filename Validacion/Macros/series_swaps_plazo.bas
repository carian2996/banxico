
Sub seriesSWAPS()

Dim macro As String, derivado As String, MaxCol As String
Dim i As Integer, LastCol As Integer
Dim e As Double, d As Double, pesta As String

macro = "CalculaPlazo.xlsm"
derivado = "swaps_plazo_09072014.xlsx"
pesta = "swaps_plazo_09072014"

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
Range("M" & e + 1).Select
ActiveCell.FormulaR1C1 = "PATITO"
Selection.Copy
Range("M" & e + 1 & ":M" & d).Select
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

Range("A5").Select
Selection.End(xlDown).Select
e = ActiveWindow.RangeSelection.Row

    Range("A5:A" & e - 1).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
   
    Sheets.Add
    ActiveSheet.Name = "SWAPS_IC_PLAZO"
    Range("B7").Select
    ActiveSheet.Paste
    Range("C6").Value = "IC - REC"
    
    Range("B7").Select
MaxCol = Cells.SpecialCells(xlLastCell).Column

'======================== SECCION II - IV ========================
Sheets("tabla").Select
    
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

Range("A5").Select
Selection.End(xlDown).Select
e = ActiveWindow.RangeSelection.Row

    Range("A5:A" & e - 1).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    
    Sheets("SWAPS_IC_PLAZO").Select
    Cells(7, MaxCol + 2).Select
    ActiveSheet.Paste
    Cells(6, MaxCol + 3).Value = "IC - EXT"
    
   Cells(7, MaxCol + 2).Select
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
    ActiveSheet.PivotTables("tabladim").PivotFields("TIPO_INST").EnableMultiplePageItems = True
    
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

Range("A5").Select
Selection.End(xlDown).Select
e = ActiveWindow.RangeSelection.Row

    Range("A5:A" & e - 1).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
   
    Sheets.Add
    ActiveSheet.Name = "SWAPS_CB_PLAZO"
    Range("B7").Select
    ActiveSheet.Paste
    Range("C6").Value = "CB - REC"
    
    Range("B7").Select
MaxCol = Cells.SpecialCells(xlLastCell).Column

'======================== SECCION II - IV ========================
Sheets("tabla").Select
    
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


Range("A5").Select
Selection.End(xlDown).Select
e = ActiveWindow.RangeSelection.Row

    Range("A5:A" & e - 1).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    
    Sheets("SWAPS_CB_PLAZO").Select
    Cells(7, MaxCol + 2).Select
    ActiveSheet.Paste
    Cells(6, MaxCol + 3).Value = "CB - EXT"
    
   Cells(7, MaxCol + 2).Select
MaxCol = Cells.SpecialCells(xlLastCell).Column

'======================== o ======================== SERIES ======================== o ======================== o
Sheets("SWAPS_IC_PLAZO").Select
Rows("8").Select
Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
Rows("7").Copy
Range("A8").Select
ActiveSheet.Paste
Range("C8").Select
i = 1

'======================== IC-REC ========================
Do While Workbooks(derivado).Sheets("SWAPS_IC_PLAZO").Cells(8, 2 + i) <> ""
 Cells(8, 2 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("SWAPS_IC_PLAZO").Cells(8, 2 + i).Value, Workbooks(macro).Sheets("IC").Range("$B$3:$C$16"), 2, False)
i = i + 1
Loop

Cells(8, 2 + i).EntireColumn.Delete
Cells(8, 2 + i).EntireColumn.Delete
Cells(8, 1 + i).EntireColumn.Delete
'======================== IC - EXT========================
Do While Workbooks(derivado).Sheets("SWAPS_IC_PLAZO").Cells(8, 1 + i) <> ""
 Cells(8, 1 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("SWAPS_IC_PLAZO").Cells(8, 1 + i).Value, Workbooks(macro).Sheets("IC").Range("$E$3:$F$16"), 2, False)
i = i + 1
Loop

Cells(8, i).EntireColumn.Delete

Range("B9").Select
Selection.End(xlDown).Select
e = ActiveWindow.RangeSelection.Row
Range("A9:A" & e).FormulaR1C1 = "=TEXT(RC[1], ""dd/mm/aaaa"")"

Range("A9:A" & e).Copy
Range("B9").PasteSpecial Paste:=xlPasteValues
Range("A9:A" & e).ClearContents


'======================== CB-REC ========================
Sheets("SWAPS_CB_PLAZO").Select
Rows("8").Select
Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
Rows("7").Copy
Range("A8").Select
ActiveSheet.Paste
Range("C8").Select
i = 1

Do While Workbooks(derivado).Sheets("SWAPS_CB_PLAZO").Cells(8, 2 + i) <> ""
 Cells(8, 2 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("SWAPS_CB_PLAZO").Cells(8, 2 + i).Value, Workbooks(macro).Sheets("CB").Range("$B$3:$C$16"), 2, False)
i = i + 1
Loop

Cells(8, 2 + i).EntireColumn.Delete
Cells(8, 2 + i).EntireColumn.Delete
Cells(8, 1 + i).EntireColumn.Delete
'======================== CB - EXT ========================
Do While Workbooks(derivado).Sheets("SWAPS_CB_PLAZO").Cells(8, 1 + i) <> ""
 Cells(8, 1 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("SWAPS_CB_PLAZO").Cells(8, 1 + i).Value, Workbooks(macro).Sheets("CB").Range("$E$3:$F$16"), 2, False)
i = i + 1
Loop

Cells(8, i).EntireColumn.Delete

Range("B9").Select
Selection.End(xlDown).Select
e = ActiveWindow.RangeSelection.Row
Range("A9:A" & e).FormulaR1C1 = "=TEXT(RC[1], ""dd/mm/aaaa"")"

Range("A9:A" & e).Copy
Range("B9").PasteSpecial Paste:=xlPasteValues
Range("A9:A" & e).ClearContents

MsgBox ("Listo")


End Sub
