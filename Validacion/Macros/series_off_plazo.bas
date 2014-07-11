
Sub seriesOFF()

Dim macro As String, derivado As String, MaxCol As String
Dim i As Integer, LastCol As Integer
Dim e As Double, pesta As String

macro = "CalculaPlazo.xlsm"
derivado = "off_plazo_09072014.xlsx"
pesta = "off_plazo_09072014"

'======================== PATOS ========================

Windows(derivado).Activate
Sheets(pesta).Select
Range("C1").Select
Selection.End(xlDown).Select 'Ubico el ?ltimo registro de la fila
e = ActiveWindow.RangeSelection.Row
Range("C2:C" & e).Select
Selection.Copy

Range("C" & e + 1).Select
ActiveSheet.Paste
Selection.RemoveDuplicates Columns:=1, Header:=xlNo

Range("C" & e + 1).Select
Selection.End(xlDown).Select 'Ubico el ?ltimo registro de la fila de fechas
d = ActiveWindow.RangeSelection.Row

'ESCRIBE PATITO
Range("K" & e + 1).Select
ActiveCell.FormulaR1C1 = "PATITO"
Selection.Copy
Range("K" & e + 1 & ":K" & d).Select
ActiveSheet.Paste

'ESCRIBE CEROS EN IMPORTE
Range("N" & e + 1).Select
ActiveCell.FormulaR1C1 = "0"
Selection.Copy
Range("N" & e + 1 & ":N" & d).Select
ActiveSheet.Paste

MsgBox ("Actualiza la tabla")

'======================== o ======================== BANCOS ======================== o ======================== o
'======================== FUTUROS ========================
Windows(derivado).Activate

Sheets("tabla").Select

        ActiveSheet.PivotTables("tabladim").PivotFields("TIP_INSTI").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("tabladim").PivotFields("TIP_INSTI")
        .PivotItems("BDE").Visible = True
        .PivotItems("BMU").Visible = True
        .PivotItems("CBO").Visible = False
        .PivotItems("(blank)").Visible = True
    End With
    ActiveSheet.PivotTables("tabladim").PivotFields("TIP_INSTI").EnableMultiplePageItems = True

    ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPE").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPE")
        .PivotItems("FUTURO").Visible = True
        .PivotItems("FORWARD").Visible = False
        .PivotItems("(blank)").Visible = True
        .PivotItems("PATITO").Visible = True
    End With
    ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPE").EnableMultiplePageItems = True
    
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
    ActiveSheet.Name = "FF_IC_PLAZO"
    Range("B7").Select
    ActiveSheet.Paste
    Range("C6").Value = "IC - FUT"
    
    Range("B7").Select
MaxCol = Cells.SpecialCells(xlLastCell).Column

'======================== FORWARD ========================
Sheets("tabla").Select
    
    ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPE").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPE")
        .PivotItems("FUTURO").Visible = False
        .PivotItems("FORWARD").Visible = True
        .PivotItems("(blank)").Visible = True
        .PivotItems("PATITO").Visible = True
    End With
    ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPE").EnableMultiplePageItems = True
    
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
    
    Sheets("FF_IC_PLAZO").Select
    Cells(7, MaxCol + 2).Select
    ActiveSheet.Paste
    Cells(6, MaxCol + 3).Value = "IC - FWD"
    
   Cells(7, MaxCol + 2).Select
MaxCol = Cells.SpecialCells(xlLastCell).Column
    
'======================== o ======================== CASAS DE BOLSA ======================== o ======================== o
'======================== FUTUROS ========================
Sheets("tabla").Select
    
    ActiveSheet.PivotTables("tabladim").PivotFields("TIP_INSTI").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("tabladim").PivotFields("TIP_INSTI")
        .PivotItems("BDE").Visible = False
        .PivotItems("BMU").Visible = False
        .PivotItems("CBO").Visible = True
        .PivotItems("(blank)").Visible = True
    End With
    ActiveSheet.PivotTables("tabladim").PivotFields("TIP_INSTI").EnableMultiplePageItems = True
    
    ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPE").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPE")
        .PivotItems("FORWARD").Visible = False
        .PivotItems("FUTURO").Visible = True
        .PivotItems("(blank)").Visible = True
        .PivotItems("PATITO").Visible = True
    End With
    ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPE").EnableMultiplePageItems = True

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
    ActiveSheet.Name = "FF_CB_PLAZO"
    Range("B7").Select
    ActiveSheet.Paste
    Range("C6").Value = "CB - FUT"
    
    Range("B7").Select
MaxCol = Cells.SpecialCells(xlLastCell).Column

'======================== FORWARD ========================
Sheets("tabla").Select
    
    ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPE").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPE")
        .PivotItems("FORWARD").Visible = True
        .PivotItems("FUTURO").Visible = False
        .PivotItems("(blank)").Visible = True
        .PivotItems("PATITO").Visible = True
    End With
    ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPE").EnableMultiplePageItems = True
    
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
    
    Sheets("FF_CB_PLAZO").Select
    Cells(7, MaxCol + 2).Select
    ActiveSheet.Paste
    Cells(6, MaxCol + 3).Value = "CB - FWD"
    
   Cells(7, MaxCol + 2).Select
MaxCol = Cells.SpecialCells(xlLastCell).Column


'======================== o ======================== SERIES ======================== o ======================== o
Sheets("FF_IC_PLAZO").Select
Rows("8").Select
Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
Rows("7").Copy
Range("A8").Select
ActiveSheet.Paste
Range("C8").Select
i = 1

'======================== IC-FUT ========================
Do While Workbooks(derivado).Sheets("FF_IC_PLAZO").Cells(8, 2 + i) <> ""
 Cells(8, 2 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("FF_IC_PLAZO").Cells(8, 2 + i).Value, Workbooks(macro).Sheets("IC").Range("$J$2:$K$16"), 2, False)
i = i + 1
Loop

Cells(8, 2 + i).EntireColumn.Delete
Cells(8, 2 + i).EntireColumn.Delete
Cells(8, 1 + i).EntireColumn.Delete
'======================== IC - FUT ========================
Do While Workbooks(derivado).Sheets("FF_IC_PLAZO").Cells(8, 1 + i) <> ""
 Cells(8, 1 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("FF_IC_PLAZO").Cells(8, 1 + i).Value, Workbooks(macro).Sheets("IC").Range("$N$2:$O$16"), 2, False)
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

'======================== CB - FUT ========================
Sheets("FF_CB_PLAZO").Select
Rows("8").Select
Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
Rows("7").Copy
Range("A8").Select
ActiveSheet.Paste
Range("C8").Select
i = 1

Do While Workbooks(derivado).Sheets("FF_CB_PLAZO").Cells(8, 2 + i) <> ""
 Cells(8, 2 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("FF_CB_PLAZO").Cells(8, 2 + i).Value, Workbooks(macro).Sheets("CB").Range("$J$2:$K$16"), 2, False)
i = i + 1
Loop

Cells(8, 2 + i).EntireColumn.Delete
Cells(8, 2 + i).EntireColumn.Delete
Cells(8, 1 + i).EntireColumn.Delete
'======================== CB - FWD ========================
Do While Workbooks(derivado).Sheets("FF_CB_PLAZO").Cells(8, 1 + i) <> ""
 Cells(8, 1 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("FF_CB_PLAZO").Cells(8, 1 + i).Value, Workbooks(macro).Sheets("CB").Range("$N$2:$O$16"), 2, False)
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
