Sub FwdFut()

Dim macro As String, derivado As String, MaxCol As String
Dim i As Integer, LastCol As Integer, f As Integer
Dim e As Double

macro = "CalculaPlazo.xlsm"
derivado = "off_plazo_07072014.xlsx"

'======================== o ======================== FUTUROS ======================== o ======================== o
'======================== BANCOS ========================
Windows(derivado).Activate

Sheets("tabla").Select
    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("CLASE_OPE").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("CLASE_OPE")
        .PivotItems("FUTURO").Visible = True
        .PivotItems("FORWARD").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("CLASE_OPE").EnableMultiplePageItems = True
    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("TIP_INSTI").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("TIP_INSTI")
        .PivotItems("BDE").Visible = True
        .PivotItems("BMU").Visible = True
        .PivotItems("CBO").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("TIP_INSTI").EnableMultiplePageItems = True
    With ActiveSheet.PivotTables("Tabla dinámica2")
        .ColumnGrand = False
        .RowGrand = False
    End With

Range("A5:A5").Select
Selection.End(xlDown).Select
e = ActiveWindow.RangeSelection.Row

    Range("A5:A" & e).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
   
    Sheets.Add
    ActiveSheet.Name = "resultado"
    Range("B3").Select
    ActiveSheet.Paste
    Range("C2").Value = "IC - FUT"
    
    Range("B3").Select
MaxCol = Cells.SpecialCells(xlLastCell).Column

'======================== CASAS DE BOLSA ========================
Sheets("tabla").Select
    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("CLASE_OPE").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("CLASE_OPE")
        .PivotItems("FUTURO").Visible = True
        .PivotItems("FORWARD").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("CLASE_OPE").EnableMultiplePageItems = True
    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("TIP_INSTI").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("TIP_INSTI")
        .PivotItems("CBO").Visible = True
        .PivotItems("BDE").Visible = False
        .PivotItems("BMU").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("TIP_INSTI").EnableMultiplePageItems = True
    With ActiveSheet.PivotTables("Tabla dinámica2")
        .ColumnGrand = False
        .RowGrand = False
    End With

Range("A5:A5").Select
Selection.End(xlDown).Select
e = ActiveWindow.RangeSelection.Row

    Range("B5:B" & e).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    
    Sheets("resultado").Select
    Cells(3, MaxCol + 2).Select
    ActiveSheet.Paste
    Cells(2, MaxCol + 2).Value = "CBO - FUT"
    
   Cells(4, MaxCol + 2).Select
MaxCol = Cells.SpecialCells(xlLastCell).Column
    
'======================== o ======================== FORWARDS ======================== o ======================== o
'======================== BANCOS ========================
Sheets("tabla").Select
    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("CLASE_OPE").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("CLASE_OPE")
        .PivotItems("FORWARD").Visible = True
        .PivotItems("FUTURO").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("CLASE_OPE").EnableMultiplePageItems = True
    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("TIP_INSTI").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("TIP_INSTI")
        .PivotItems("BDE").Visible = True
        .PivotItems("BMU").Visible = True
        .PivotItems("CBO").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("TIP_INSTI").EnableMultiplePageItems = True
    With ActiveSheet.PivotTables("Tabla dinámica2")
        .ColumnGrand = False
        .RowGrand = False
    End With
    
Range("A5:A5").Select
Selection.End(xlDown).Select
e = ActiveWindow.RangeSelection.Row

    Range("B5:B" & e).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    
    Sheets("resultado").Select
    Cells(3, MaxCol + 2).Select
    ActiveSheet.Paste
    Cells(2, MaxCol + 2).Value = "IC - FWD"
    
   Cells(3, MaxCol + 2).Select
MaxCol = Cells.SpecialCells(xlLastCell).Column

'======================== CASAS DE BOLSA ========================
Sheets("tabla").Select
    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("CLASE_OPE").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("CLASE_OPE")
        .PivotItems("FORWARD").Visible = True
        .PivotItems("FUTURO").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("CLASE_OPE").EnableMultiplePageItems = True
    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("TIP_INSTI").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("TIP_INSTI")
    .PivotItems("CBO").Visible = True
        .PivotItems("BDE").Visible = False
        .PivotItems("BMU").Visible = False
        .PivotItems("(blank)").Visible = False
    End With
    ActiveSheet.PivotTables("Tabla dinámica2").PivotFields("TIP_INSTI").EnableMultiplePageItems = True
    With ActiveSheet.PivotTables("Tabla dinámica2")
        .ColumnGrand = False
        .RowGrand = False
    End With
    
Range("A5:A5").Select
Selection.End(xlDown).Select
e = ActiveWindow.RangeSelection.Row

    Range("B5:B" & e).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    
    Sheets("resultado").Select
    Cells(3, MaxCol + 2).Select
    ActiveSheet.Paste
    Cells(2, MaxCol + 2).Value = "CBO - FWD"
    
   Cells(3, MaxCol + 2).Select
MaxCol = Cells.SpecialCells(xlLastCell).Column

'======================== o ======================== SERIES ======================== o ======================== o
Range("C3").Select
i = 1

'======================== FTR & IC ========================
Do While Workbooks(derivado).Sheets("resultado").Cells(3, 2 + i) <> ""
 Cells(3, 2 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("resultado").Cells(3, 2 + i).Value, Workbooks(macro).Sheets("IC").Range("$J$3:$K$16"), 2, False)
i = i + 1
Loop

Cells(3, 2 + i).EntireColumn.Delete
'======================== FTR & CBO ========================
Do While Workbooks(derivado).Sheets("resultado").Cells(3, 2 + i) <> ""
 Cells(3, 2 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("resultado").Cells(3, 2 + i).Value, Workbooks(macro).Sheets("CB").Range("$J$3:$K$16"), 2, False)
i = i + 1
Loop

Cells(3, 2 + i).EntireColumn.Delete
'======================== FWD & IC ========================
Do While Workbooks(derivado).Sheets("resultado").Cells(3, 2 + i) <> ""
 Cells(3, 2 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("resultado").Cells(3, 2 + i).Value, Workbooks(macro).Sheets("IC").Range("$N$3:$O$16"), 2, False)
i = i + 1
Loop

Cells(3, 2 + i).EntireColumn.Delete
'======================== FWD & CBO ========================
Do While Workbooks(derivado).Sheets("resultado").Cells(3, 2 + i) <> ""
 Cells(3, 2 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("resultado").Cells(3, 2 + i).Value, Workbooks(macro).Sheets("CB").Range("$J$3:$K$16"), 2, False)
i = i + 1
Loop

MsgBox ("Listo")

End Sub