Sub seriesOPTO()

Dim macro As String, ruta As String, tcfix As String, udi As String, ruta2 As String, plazo As String
Dim j As Integer
Dim rango1 As Range
Dim b As Double, e As Double, f As Double, i As Double, d As Double
Dim MaxCol As String, pesta As String


macro = "CalculaContraparte.xlsm"
derivado = "opto_contra_10072014.xlsx"
pesta = "opto_contra_10072014"


'======================== PATOS ========================

Windows(derivado).Activate
Sheets(pesta).Select
Range("E1").Select
Selection.End(xlDown).Select 'Ubico el ?ltimo registro de la fila
e = ActiveWindow.RangeSelection.Row
Range("E2:E" & e).Select
Selection.Copy

Range("E" & e + 1).Select
ActiveSheet.Paste
Selection.RemoveDuplicates Columns:=1, Header:=xlNo

Range("E" & e + 1).Select
Selection.End(xlDown).Select 'Ubico el ultimo registro de la fila de fechas
d = ActiveWindow.RangeSelection.Row

'ESCRIBE PATITO
Range("T" & e + 1).Select
ActiveCell.FormulaR1C1 = "PATITO"
Selection.Copy
Range("T" & e + 1 & ":T" & d).Select
ActiveSheet.Paste

'ESCRIBE CEROS EN IMPORTE
Range("S" & e + 1).Select
ActiveCell.FormulaR1C1 = "0"
Selection.Copy
Range("S" & e + 1 & ":S" & d).Select
ActiveSheet.Paste

MsgBox ("Actualiza la tabla")

'======================== o ======================== BANCOS ======================== o ======================== o
'======================== EMISION ========================
'=========== ESTANTDAR ===========

Windows(derivado).Activate
Sheets("tabla").Select
    
    ActiveSheet.PivotTables("tabladim").PivotFields("TIPO_INST").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("tabladim").PivotFields("TIPO_INST")
        .PivotItems("BM_BD").Visible = True
        .PivotItems("CB").Visible = False
        .PivotItems("(blank)").Visible = True
    End With
    ActiveSheet.PivotTables("tabladim").PivotFields("TIPO_INST").EnableMultiplePageItems = True
    
    ActiveSheet.PivotTables("tabladim").PivotFields("TIP_OPE").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("tabladim").PivotFields("TIP_OPE")
        .PivotItems("E").Visible = True
        .PivotItems("A").Visible = False
        .PivotItems("(blank)").Visible = True
    End With
    ActiveSheet.PivotTables("tabladim").PivotFields("TIP_OPE").EnableMultiplePageItems = True
    
    ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPER").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPER")
        .PivotItems("Estandar").Visible = True
        .PivotItems("Exotica").Visible = False
        .PivotItems("Warrant").Visible = False
        .PivotItems("(blank)").Visible = True
        .PivotItems("PATITO").Visible = True
    End With
    ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPER").EnableMultiplePageItems = True

    With ActiveSheet.PivotTables("tabladim")
        .ColumnGrand = False
        .RowGrand = False
    End With

Range("A7").Select
Selection.End(xlDown).Select
e = ActiveWindow.RangeSelection.Row

    Range("A7:A" & e - 1).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
   
    Sheets.Add
    ActiveSheet.Name = "resultado"
    Range("B3").Select
    ActiveSheet.Paste
    Range("C2").Value = "IC-EMI-EST"
    
    Range("B3").Select
MaxCol = Cells.SpecialCells(xlLastCell).Column
   
'=========== EXOTICAS ===========
Windows(derivado).Activate
Sheets("tabla").Select
    
    ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPER").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPER")
        .PivotItems("Estandar").Visible = False
        .PivotItems("Exotica").Visible = True
        .PivotItems("Warrant").Visible = False
        .PivotItems("(blank)").Visible = True
        .PivotItems("PATITO").Visible = True
    End With
    ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPER").EnableMultiplePageItems = True

Range("A7").Select
Selection.End(xlDown).Select
e = ActiveWindow.RangeSelection.Row

    Range("A7:A" & e - 1).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    
    Sheets("resultado").Select
    Cells(3, MaxCol + 2).Select
    ActiveSheet.Paste
    Cells(2, MaxCol + 3).Value = "IC-EMI-EXO"
    
   Cells(3, MaxCol + 2).Select
MaxCol = Cells.SpecialCells(xlLastCell).Column

'=========== WARRANT ===========
Windows(derivado).Activate
Sheets("tabla").Select
    
    ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPER").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPER")
        .PivotItems("Estandar").Visible = False
        .PivotItems("Exotica").Visible = False
        .PivotItems("Warrant").Visible = True
        .PivotItems("(blank)").Visible = True
        .PivotItems("PATITO").Visible = True
    End With
    ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPER").EnableMultiplePageItems = True
    
Range("A7").Select
Selection.End(xlDown).Select
e = ActiveWindow.RangeSelection.Row

    Range("A7:A" & e - 1).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    
    Sheets("resultado").Select
    Cells(3, MaxCol + 2).Select
    ActiveSheet.Paste
    Cells(2, MaxCol + 3).Value = "IC-EMI-WAR"
    
   Cells(3, MaxCol + 2).Select
MaxCol = Cells.SpecialCells(xlLastCell).Column


'======================== ADQUISICION ========================
'=========== ESTANTDAR ===========
Windows(derivado).Activate
Sheets("tabla").Select
    
    ActiveSheet.PivotTables("tabladim").PivotFields("TIP_OPE").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("tabladim").PivotFields("TIP_OPE")
        .PivotItems("E").Visible = False
        .PivotItems("A").Visible = True
        .PivotItems("(blank)").Visible = True
    End With
    ActiveSheet.PivotTables("tabladim").PivotFields("TIP_OPE").EnableMultiplePageItems = True
    
    ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPER").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPER")
        .PivotItems("Estandar").Visible = True
        .PivotItems("Exotica").Visible = False
        .PivotItems("Warrant").Visible = False
        .PivotItems("(blank)").Visible = True
        .PivotItems("PATITO").Visible = True
    End With
    ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPER").EnableMultiplePageItems = True
    
    With ActiveSheet.PivotTables("tabladim")
        .ColumnGrand = False
        .RowGrand = False
    End With

Range("A7").Select
Selection.End(xlDown).Select
e = ActiveWindow.RangeSelection.Row

    Range("A7:A" & e - 1).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    
    Sheets("resultado").Select
    Cells(3, MaxCol + 2).Select
    ActiveSheet.Paste
    Cells(2, MaxCol + 3).Value = "IC-ADQ-EST-EXT"
    
   Cells(3, MaxCol + 2).Select
MaxCol = Cells.SpecialCells(xlLastCell).Column
   
'=========== EXOTICAS ===========
Windows(derivado).Activate
Sheets("tabla").Select
    
    ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPER").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPER")
        .PivotItems("Estandar").Visible = False
        .PivotItems("Exotica").Visible = True
        .PivotItems("Warrant").Visible = False
        .PivotItems("(blank)").Visible = True
        .PivotItems("PATITO").Visible = True
    End With
    ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPER").EnableMultiplePageItems = True

Range("A7").Select
Selection.End(xlDown).Select
e = ActiveWindow.RangeSelection.Row

    Range("A7:A" & e - 1).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    
    Sheets("resultado").Select
    Cells(3, MaxCol + 2).Select
    ActiveSheet.Paste
    Cells(2, MaxCol + 3).Value = "IC-ADQ-EXO"
    
   Cells(3, MaxCol + 2).Select
MaxCol = Cells.SpecialCells(xlLastCell).Column

'=========== WARRANT ===========
Windows(derivado).Activate
Sheets("tabla").Select
    
    ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPER").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPER")
        .PivotItems("Estandar").Visible = False
        .PivotItems("Exotica").Visible = False
        .PivotItems("Warrant").Visible = True
        .PivotItems("(blank)").Visible = True
        .PivotItems("PATITO").Visible = True
    End With
    ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPER").EnableMultiplePageItems = True
    
Range("A7").Select
Selection.End(xlDown).Select
e = ActiveWindow.RangeSelection.Row

    Range("A7:A" & e - 1).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    
    Sheets("resultado").Select
    Cells(3, MaxCol + 2).Select
    ActiveSheet.Paste
    Cells(2, MaxCol + 3).Value = "IC-ADQ-WAR"
    
   Cells(3, MaxCol + 2).Select
MaxCol = Cells.SpecialCells(xlLastCell).Column

'======================== o ======================== CASAS DE BOLSA ======================== o ======================== o
'======================== EMISION ========================
'=========== ESTANTDAR ===========

Windows(derivado).Activate
Sheets("tabla").Select
    
    ActiveSheet.PivotTables("tabladim").PivotFields("TIPO_INST").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("tabladim").PivotFields("TIPO_INST")
        .PivotItems("BM_BD").Visible = False
        .PivotItems("CB").Visible = True
        .PivotItems("(blank)").Visible = True
    End With
    ActiveSheet.PivotTables("tabladim").PivotFields("TIPO_INST").EnableMultiplePageItems = True
    
    ActiveSheet.PivotTables("tabladim").PivotFields("TIP_OPE").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("tabladim").PivotFields("TIP_OPE")
        .PivotItems("E").Visible = True
        .PivotItems("A").Visible = False
        .PivotItems("(blank)").Visible = True
    End With
    ActiveSheet.PivotTables("tabladim").PivotFields("TIP_OPE").EnableMultiplePageItems = True
    
    ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPER").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPER")
        .PivotItems("Estandar").Visible = True
        .PivotItems("Exotica").Visible = False
        .PivotItems("Warrant").Visible = False
        .PivotItems("(blank)").Visible = True
        .PivotItems("PATITO").Visible = True
    End With
    ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPER").EnableMultiplePageItems = True

    With ActiveSheet.PivotTables("tabladim")
        .ColumnGrand = False
        .RowGrand = False
    End With

Range("A7").Select
Selection.End(xlDown).Select
e = ActiveWindow.RangeSelection.Row

    Range("A7:A" & e - 1).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    
    Sheets("resultado").Select
    Cells(3, MaxCol + 2).Select
    ActiveSheet.Paste
    Cells(2, MaxCol + 3).Value = "CB-EMI-EST"
    
   Cells(3, MaxCol + 2).Select
MaxCol = Cells.SpecialCells(xlLastCell).Column
   
'=========== EXOTICAS ===========
Windows(derivado).Activate
Sheets("tabla").Select
    
    ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPER").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPER")
        .PivotItems("Estandar").Visible = False
        .PivotItems("Exotica").Visible = True
        .PivotItems("Warrant").Visible = False
        .PivotItems("(blank)").Visible = True
        .PivotItems("PATITO").Visible = True
    End With
    ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPER").EnableMultiplePageItems = True

Range("A7").Select
Selection.End(xlDown).Select
e = ActiveWindow.RangeSelection.Row

    Range("A7:A" & e - 1).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    
    Sheets("resultado").Select
    Cells(3, MaxCol + 2).Select
    ActiveSheet.Paste
    Cells(2, MaxCol + 3).Value = "CB-EMI-EXO"
    
   Cells(3, MaxCol + 2).Select
MaxCol = Cells.SpecialCells(xlLastCell).Column

'=========== WARRANT ===========
Windows(derivado).Activate
Sheets("tabla").Select
    
    ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPER").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPER")
        .PivotItems("Estandar").Visible = False
        .PivotItems("Exotica").Visible = False
        .PivotItems("Warrant").Visible = True
        .PivotItems("(blank)").Visible = True
        .PivotItems("PATITO").Visible = True
    End With
    ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPER").EnableMultiplePageItems = True
    
Range("A7").Select
Selection.End(xlDown).Select
e = ActiveWindow.RangeSelection.Row

    Range("A7:A" & e - 1).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    
    Sheets("resultado").Select
    Cells(3, MaxCol + 2).Select
    ActiveSheet.Paste
    Cells(2, MaxCol + 3).Value = "CB-EMI-WAR"
    
   Cells(3, MaxCol + 2).Select
MaxCol = Cells.SpecialCells(xlLastCell).Column


'======================== ADQUISICION ========================
'=========== ESTANTDAR ===========
Windows(derivado).Activate
Sheets("tabla").Select
    
    ActiveSheet.PivotTables("tabladim").PivotFields("TIPO_INST").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("tabladim").PivotFields("TIPO_INST")
        .PivotItems("BM_BD").Visible = False
        .PivotItems("CB").Visible = True
        .PivotItems("(blank)").Visible = True
    End With
    ActiveSheet.PivotTables("tabladim").PivotFields("TIPO_INST").EnableMultiplePageItems = True
    
    ActiveSheet.PivotTables("tabladim").PivotFields("TIP_OPE").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("tabladim").PivotFields("TIP_OPE")
        .PivotItems("E").Visible = False
        .PivotItems("A").Visible = True
        .PivotItems("(blank)").Visible = True
    End With
    ActiveSheet.PivotTables("tabladim").PivotFields("TIP_OPE").EnableMultiplePageItems = True
    
    ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPER").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPER")
        .PivotItems("Estandar").Visible = True
        .PivotItems("Exotica").Visible = False
        .PivotItems("Warrant").Visible = False
        .PivotItems("(blank)").Visible = True
        .PivotItems("PATITO").Visible = True
    End With
    ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPER").EnableMultiplePageItems = True

    With ActiveSheet.PivotTables("tabladim")
        .ColumnGrand = False
        .RowGrand = False
    End With

Range("A7").Select
Selection.End(xlDown).Select
e = ActiveWindow.RangeSelection.Row

    Range("A7:A" & e - 1).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    
    Sheets("resultado").Select
    Cells(3, MaxCol + 2).Select
    ActiveSheet.Paste
    Cells(2, MaxCol + 3).Value = "CB-ADQ-EST"
    
   Cells(3, MaxCol + 2).Select
MaxCol = Cells.SpecialCells(xlLastCell).Column

'=========== EXOTICAS ===========
Windows(derivado).Activate
Sheets("tabla").Select
    
    ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPER").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPER")
        .PivotItems("Estandar").Visible = False
        .PivotItems("Exotica").Visible = True
        .PivotItems("Warrant").Visible = False
        .PivotItems("(blank)").Visible = True
        .PivotItems("PATITO").Visible = True
    End With
    ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPER").EnableMultiplePageItems = True

Range("A7").Select
Selection.End(xlDown).Select
e = ActiveWindow.RangeSelection.Row

    Range("A7:A" & e - 1).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    
    Sheets("resultado").Select
    Cells(3, MaxCol + 2).Select
    ActiveSheet.Paste
    Cells(2, MaxCol + 3).Value = "CB-ADQ-EXO"
    
   Cells(3, MaxCol + 2).Select
MaxCol = Cells.SpecialCells(xlLastCell).Column

'=========== WARRANT ===========
Windows(derivado).Activate
Sheets("tabla").Select
    
    ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPER").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPER")
        .PivotItems("Estandar").Visible = False
        .PivotItems("Exotica").Visible = False
        .PivotItems("Warrant").Visible = True
        .PivotItems("(blank)").Visible = True
        .PivotItems("PATITO").Visible = True
    End With
    ActiveSheet.PivotTables("tabladim").PivotFields("CLASE_OPER").EnableMultiplePageItems = True
    
Range("A7").Select
Selection.End(xlDown).Select
e = ActiveWindow.RangeSelection.Row

    Range("A7:A" & e - 1).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    
    Sheets("resultado").Select
    Cells(3, MaxCol + 2).Select
    ActiveSheet.Paste
    Cells(2, MaxCol + 3).Value = "CB-ADQ-WAR"
    
   Cells(3, MaxCol + 2).Select
MaxCol = Cells.SpecialCells(xlLastCell).Column

'======================== o ======================== SERIES ======================== o ======================== o
Rows("4").Select
Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
Rows("3").Copy
Range("A4").Select
ActiveSheet.Paste

Range("C3").Select
i = 1

'======================== IC-EMI-EST ========================
Do While Workbooks(derivado).Sheets("resultado").Cells(3, 2 + i) <> ""
 Cells(3, 2 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("resultado").Cells(3, 2 + i).Value, Workbooks(macro).Sheets("IC").Range("$B$36:$C$68"), 2, False)
i = i + 1
Loop

Cells(3, 2 + i).EntireColumn.Delete
Cells(3, 2 + i).EntireColumn.Delete

'======================== IC-EMI-EXO ========================
Do While Workbooks(derivado).Sheets("resultado").Cells(3, 2 + i) <> ""
 Cells(3, 2 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("resultado").Cells(3, 2 + i).Value, Workbooks(macro).Sheets("IC").Range("$E$36:$F$63"), 2, False)
i = i + 1
Loop

Cells(3, 2 + i).EntireColumn.Delete
Cells(3, 2 + i).EntireColumn.Delete

'======================== IC-EMI-WAR ========================
Do While Workbooks(derivado).Sheets("resultado").Cells(3, 2 + i) <> ""
 Cells(3, 2 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("resultado").Cells(3, 2 + i).Value, Workbooks(macro).Sheets("IC").Range("$H$36:$I$63"), 2, False)
i = i + 1
Loop

Cells(3, 2 + i).EntireColumn.Delete
Cells(3, 2 + i).EntireColumn.Delete

'======================== IC-ADQ-EST ========================
Do While Workbooks(derivado).Sheets("resultado").Cells(3, 2 + i) <> ""
 Cells(3, 2 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("resultado").Cells(3, 2 + i).Value, Workbooks(macro).Sheets("IC").Range("$B$81:$C$113"), 2, False)
i = i + 1
Loop

Cells(3, 2 + i).EntireColumn.Delete
Cells(3, 2 + i).EntireColumn.Delete

'======================== IC-ADQ-EXO ========================
Do While Workbooks(derivado).Sheets("resultado").Cells(3, 2 + i) <> ""
 Cells(3, 2 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("resultado").Cells(3, 2 + i).Value, Workbooks(macro).Sheets("IC").Range("$E$81:$F$108"), 2, False)
i = i + 1
Loop

Cells(3, 2 + i).EntireColumn.Delete
Cells(3, 2 + i).EntireColumn.Delete

'======================== IC-ADQ-WAR ========================
Do While Workbooks(derivado).Sheets("resultado").Cells(3, 2 + i) <> ""
 Cells(3, 2 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("resultado").Cells(3, 2 + i).Value, Workbooks(macro).Sheets("IC").Range("$H$81:$I$108"), 2, False)
i = i + 1
Loop

Cells(3, 2 + i).EntireColumn.Delete
Cells(3, 2 + i).EntireColumn.Delete

'======================== CB-EMI-EST ========================
Do While Workbooks(derivado).Sheets("resultado").Cells(3, 2 + i) <> ""
 Cells(3, 2 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("resultado").Cells(3, 2 + i).Value, Workbooks(macro).Sheets("CB").Range("$B$37:$C$69"), 2, False)
i = i + 1
Loop

Cells(3, 2 + i).EntireColumn.Delete
Cells(3, 2 + i).EntireColumn.Delete

'======================== CB-EMI-EXO ========================
Do While Workbooks(derivado).Sheets("resultado").Cells(3, 2 + i) <> ""
 Cells(3, 2 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("resultado").Cells(3, 2 + i).Value, Workbooks(macro).Sheets("CB").Range("$E$37:$F$64"), 2, False)
i = i + 1
Loop

Cells(3, 2 + i).EntireColumn.Delete
Cells(3, 2 + i).EntireColumn.Delete

'======================== CB-EMI-WAR ========================
Do While Workbooks(derivado).Sheets("resultado").Cells(3, 2 + i) <> ""
 Cells(3, 2 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("resultado").Cells(3, 2 + i).Value, Workbooks(macro).Sheets("CB").Range("$H$37:$I$64"), 2, False)
i = i + 1
Loop

Cells(3, 2 + i).EntireColumn.Delete
Cells(3, 2 + i).EntireColumn.Delete

'======================== CB-ADQ-EST ========================
Do While Workbooks(derivado).Sheets("resultado").Cells(3, 2 + i) <> ""
 Cells(3, 2 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("resultado").Cells(3, 2 + i).Value, Workbooks(macro).Sheets("CB").Range("$B$77:$C$109"), 2, False)
i = i + 1
Loop

Cells(3, 2 + i).EntireColumn.Delete
Cells(3, 2 + i).EntireColumn.Delete

'======================== CB-EMI-EXO ========================
Do While Workbooks(derivado).Sheets("resultado").Cells(3, 2 + i) <> ""
 Cells(3, 2 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("resultado").Cells(3, 2 + i).Value, Workbooks(macro).Sheets("CB").Range("$E$77:$F$104"), 2, False)
i = i + 1
Loop

Cells(3, 2 + i).EntireColumn.Delete
Cells(3, 2 + i).EntireColumn.Delete

'======================== CB-EMI-WAR ========================
Do While Workbooks(derivado).Sheets("resultado").Cells(3, 2 + i) <> ""
 Cells(3, 2 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("resultado").Cells(3, 2 + i).Value, Workbooks(macro).Sheets("CB").Range("$H$77:$I$104"), 2, False)
i = i + 1
Loop

Cells(3, 2 + i).EntireColumn.Delete
Cells(3, 2 + i).EntireColumn.Delete


MsgBox ("Listo")

End Sub
