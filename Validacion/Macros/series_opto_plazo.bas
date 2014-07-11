
Sub seriesOPTO()

Dim macro As String, ruta As String, tcfix As String, udi As String, ruta2 As String, plazo As String
Dim j As Integer
Dim rango1 As Range
Dim b As Double, e As Double, f As Double, i As Double, d As Double
Dim MaxCol As String, pesta As String


macro = "CalculaPlazo.xlsm"
derivado = "opto_plazo_09072014.xlsx"
pesta = "opto_plazo_09072014"


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
Range("V" & e + 1).Select
ActiveCell.FormulaR1C1 = "PATITO"
Selection.Copy
Range("V" & e + 1 & ":V" & d).Select
ActiveSheet.Paste

'ESCRIBE CEROS EN IMPORTE
Range("U" & e + 1).Select
ActiveCell.FormulaR1C1 = "0"
Selection.Copy
Range("U" & e + 1 & ":U" & d).Select
ActiveSheet.Paste

MsgBox ("Actualiza la tabla")

'======================== o ======================== BANCOS ======================== o ======================== o
'======================== EMISION ========================
'=========== ESTANTDAR ===========
' EXTRABURSATIL

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
    
    ActiveSheet.PivotTables("tabladim").PivotFields("MDO").CurrentPage = "(All)"
        With ActiveSheet.PivotTables("tabladim").PivotFields("MDO")
        .PivotItems("(blank)").Visible = True
        For i = 1 To .PivotItems.Count - 1
            .PivotItems(i).Visible = False
        Next
        .PivotItems("E").Visible = True
        EnableMultiplePageItems = True
    End With
    ActiveSheet.PivotTables("tabladim").PivotFields("MDO").EnableMultiplePageItems = True
    
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
    ActiveSheet.Name = "OPTO_IC_PLAZO"
    Range("B7").Select
    ActiveSheet.Paste
    Range("C6").Value = "IC-EMI-EST-EXT"
    
    Range("B7").Select
MaxCol = Cells.SpecialCells(xlLastCell).Column


' ORGANIZADOS
Windows(derivado).Activate
Sheets("tabla").Select
    
    ActiveSheet.PivotTables("tabladim").PivotFields("MDO").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("tabladim").PivotFields("MDO")
        .PivotItems("(blank)").Visible = True
        For i = 1 To .PivotItems.Count - 1
            .PivotItems(i).Visible = True
        Next
        .PivotItems("E").Visible = False
    End With
    ActiveSheet.PivotTables("tabladim").PivotFields("MDO").EnableMultiplePageItems = True
 
Range("A7").Select
Selection.End(xlDown).Select
e = ActiveWindow.RangeSelection.Row

    Range("A7:A" & e - 1).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    
    Sheets("OPTO_IC_PLAZO").Select
    Cells(7, MaxCol + 2).Select
    ActiveSheet.Paste
    Cells(6, MaxCol + 3).Value = "IC-EMI-EST-REC"
    
   Cells(7, MaxCol + 2).Select
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
    
    ActiveSheet.PivotTables("tabladim").PivotFields("MDO").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("tabladim").PivotFields("MDO")
        .PivotItems("E").Visible = True
        EnableMultiplePageItems = True
    End With
    ActiveSheet.PivotTables("tabladim").PivotFields("MDO").EnableMultiplePageItems = True
    
Range("A7").Select
Selection.End(xlDown).Select
e = ActiveWindow.RangeSelection.Row

    Range("A7:A" & e - 1).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    
    Sheets("OPTO_IC_PLAZO").Select
    Cells(7, MaxCol + 2).Select
    ActiveSheet.Paste
    Cells(6, MaxCol + 3).Value = "IC-EMI-EXO"
    
   Cells(7, MaxCol + 2).Select
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
    
    Sheets("OPTO_IC_PLAZO").Select
    Cells(7, MaxCol + 2).Select
    ActiveSheet.Paste
    Cells(6, MaxCol + 3).Value = "IC-EMI-WAR"
    
   Cells(7, MaxCol + 2).Select
MaxCol = Cells.SpecialCells(xlLastCell).Column


'======================== ADQUISICION ========================
'=========== ESTANTDAR ===========
' EXTRABURSATIL
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
    
    ActiveSheet.PivotTables("tabladim").PivotFields("MDO").CurrentPage = "(All)"
        With ActiveSheet.PivotTables("tabladim").PivotFields("MDO")
        .PivotItems("(blank)").Visible = True
        For i = 1 To .PivotItems.Count - 1
            .PivotItems(i).Visible = False
        Next
        .PivotItems("E").Visible = True
        EnableMultiplePageItems = True
    End With
    
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
    
    Sheets("OPTO_IC_PLAZO").Select
    Cells(7, MaxCol + 2).Select
    ActiveSheet.Paste
    Cells(6, MaxCol + 3).Value = "IC-ADQ-EST-REC"
    
   Cells(7, MaxCol + 2).Select
MaxCol = Cells.SpecialCells(xlLastCell).Column

' ORGANIZADOS
Windows(derivado).Activate
Sheets("tabla").Select
    
    ActiveSheet.PivotTables("tabladim").PivotFields("MDO").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("tabladim").PivotFields("MDO")
        .PivotItems("(blank)").Visible = True
        For i = 1 To .PivotItems.Count - 1
            .PivotItems(i).Visible = True
        Next
        .PivotItems("E").Visible = False
    End With
    ActiveSheet.PivotTables("tabladim").PivotFields("MDO").EnableMultiplePageItems = True
 
Range("A7").Select
Selection.End(xlDown).Select
e = ActiveWindow.RangeSelection.Row

    Range("A7:A" & e - 1).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    
    Sheets("OPTO_IC_PLAZO").Select
    Cells(7, MaxCol + 2).Select
    ActiveSheet.Paste
    Cells(6, MaxCol + 3).Value = "IC-ADQ-EST-REC"
    
   Cells(7, MaxCol + 2).Select
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
    
    Sheets("OPTO_IC_PLAZO").Select
    Cells(7, MaxCol + 2).Select
    ActiveSheet.Paste
    Cells(6, MaxCol + 3).Value = "IC-ADQ-EXO"
    
   Cells(7, MaxCol + 2).Select
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
    
    Sheets("OPTO_IC_PLAZO").Select
    Cells(7, MaxCol + 2).Select
    ActiveSheet.Paste
    Cells(6, MaxCol + 3).Value = "IC-ADQ-WAR"
    
   Cells(7, MaxCol + 2).Select
MaxCol = Cells.SpecialCells(xlLastCell).Column

'======================== o ======================== CASAS DE BOLSA ======================== o ======================== o
'======================== EMISION ========================
'=========== ESTANTDAR ===========
' EXTRABURSATIL

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
    
    ActiveSheet.PivotTables("tabladim").PivotFields("MDO").CurrentPage = "(All)"
        With ActiveSheet.PivotTables("tabladim").PivotFields("MDO")
        .PivotItems("(blank)").Visible = True
        For i = 1 To .PivotItems.Count - 1
            .PivotItems(i).Visible = False
        Next
        .PivotItems("E").Visible = True
        EnableMultiplePageItems = True
    End With
    ActiveSheet.PivotTables("tabladim").PivotFields("MDO").EnableMultiplePageItems = True
    
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
    ActiveSheet.Name = "OPTO_CB_PLAZO"
    Range("B7").Select
    ActiveSheet.Paste
    Range("C6").Value = "CB-EMI-EST-EXT"
    
    Range("B7").Select
MaxCol = Cells.SpecialCells(xlLastCell).Column


' ORGANIZADOS
Windows(derivado).Activate
Sheets("tabla").Select
    
    ActiveSheet.PivotTables("tabladim").PivotFields("MDO").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("tabladim").PivotFields("MDO")
        .PivotItems("(blank)").Visible = True
        For i = 1 To .PivotItems.Count - 1
            .PivotItems(i).Visible = True
        Next
        .PivotItems("E").Visible = False
    End With
    ActiveSheet.PivotTables("tabladim").PivotFields("MDO").EnableMultiplePageItems = True
 
Range("A7").Select
Selection.End(xlDown).Select
e = ActiveWindow.RangeSelection.Row

    Range("A7:A" & e - 1).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    
    Sheets("OPTO_CB_PLAZO").Select
    Cells(7, MaxCol + 2).Select
    ActiveSheet.Paste
    Cells(6, MaxCol + 3).Value = "CB-EMI-EST-REC"
    
   Cells(7, MaxCol + 2).Select
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
    
    ActiveSheet.PivotTables("tabladim").PivotFields("MDO").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("tabladim").PivotFields("MDO")
        .PivotItems("E").Visible = True
        EnableMultiplePageItems = True
    End With
    ActiveSheet.PivotTables("tabladim").PivotFields("MDO").EnableMultiplePageItems = True
    
Range("A7").Select
Selection.End(xlDown).Select
e = ActiveWindow.RangeSelection.Row

    Range("A7:A" & e - 1).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    
    Sheets("OPTO_CB_PLAZO").Select
    Cells(7, MaxCol + 2).Select
    ActiveSheet.Paste
    Cells(6, MaxCol + 3).Value = "CB-EMI-EXO"
    
   Cells(7, MaxCol + 2).Select
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
    
    Sheets("OPTO_CB_PLAZO").Select
    Cells(7, MaxCol + 2).Select
    ActiveSheet.Paste
    Cells(6, MaxCol + 3).Value = "CB-EMI-WAR"
    
   Cells(7, MaxCol + 2).Select
MaxCol = Cells.SpecialCells(xlLastCell).Column


'======================== ADQUISICION ========================
'=========== ESTANTDAR ===========
' EXTRABURSATIL

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
    
    ActiveSheet.PivotTables("tabladim").PivotFields("MDO").CurrentPage = "(All)"
        With ActiveSheet.PivotTables("tabladim").PivotFields("MDO")
        .PivotItems("(blank)").Visible = True
        For i = 1 To .PivotItems.Count - 1
            .PivotItems(i).Visible = False
        Next
        .PivotItems("E").Visible = True
        EnableMultiplePageItems = True
    End With
    ActiveSheet.PivotTables("tabladim").PivotFields("MDO").EnableMultiplePageItems = True
    
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
    
    Sheets("OPTO_CB_PLAZO").Select
    Cells(7, MaxCol + 2).Select
    ActiveSheet.Paste
    Cells(6, MaxCol + 3).Value = "CB-ADQ-EST-EXT"
    
   Cells(7, MaxCol + 2).Select
MaxCol = Cells.SpecialCells(xlLastCell).Column


' ORGANIZADOS
Windows(derivado).Activate
Sheets("tabla").Select
    
    ActiveSheet.PivotTables("tabladim").PivotFields("MDO").CurrentPage = "(All)"
    With ActiveSheet.PivotTables("tabladim").PivotFields("MDO")
        .PivotItems("(blank)").Visible = True
        For i = 1 To .PivotItems.Count - 1
            .PivotItems(i).Visible = True
        Next
        .PivotItems("E").Visible = False
    End With
    ActiveSheet.PivotTables("tabladim").PivotFields("MDO").EnableMultiplePageItems = True
 
Range("A7").Select
Selection.End(xlDown).Select
e = ActiveWindow.RangeSelection.Row

    Range("A7:A" & e - 1).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    
    Sheets("OPTO_CB_PLAZO").Select
    Cells(7, MaxCol + 2).Select
    ActiveSheet.Paste
    Cells(6, MaxCol + 3).Value = "CB-ADQ-EST-REC"
    
   Cells(7, MaxCol + 2).Select
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
    
    Sheets("OPTO_CB_PLAZO").Select
    Cells(7, MaxCol + 2).Select
    ActiveSheet.Paste
    Cells(6, MaxCol + 3).Value = "CB-ADQ-EXO"
    
   Cells(7, MaxCol + 2).Select
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
    
    Sheets("OPTO_CB_PLAZO").Select
    Cells(7, MaxCol + 2).Select
    ActiveSheet.Paste
    Cells(6, MaxCol + 3).Value = "CB-ADQ-WAR"
    
   Cells(7, MaxCol + 2).Select
MaxCol = Cells.SpecialCells(xlLastCell).Column

'======================== o ======================== SERIES ======================== o ======================== o
Sheets("OPTO_IC_PLAZO").Select
Rows("8").Select
Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
Rows("7").Copy
Range("A8").Select
ActiveSheet.Paste
Range("C8").Select
i = 1

'======================== IC-EMI-EST-EXT ========================
Do While Workbooks(derivado).Sheets("OPTO_IC_PLAZO").Cells(8, 2 + i) <> ""
 Cells(8, 2 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("OPTO_IC_PLAZO").Cells(8, 2 + i).Value, Workbooks(macro).Sheets("IC").Range("$B$22:$C$36"), 2, False)
i = i + 1
Loop

Cells(8, 2 + i).EntireColumn.Delete
Cells(8, 2 + i).EntireColumn.Delete
Cells(8, 1 + i).EntireColumn.Delete

'======================== IC-EMI-EST-ORG ========================
Do While Workbooks(derivado).Sheets("OPTO_IC_PLAZO").Cells(8, 1 + i) <> ""
 Cells(8, 1 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("OPTO_IC_PLAZO").Cells(8, 1 + i).Value, Workbooks(macro).Sheets("IC").Range("$E$22:$F$36"), 2, False)
i = i + 1
Loop

Cells(8, 1 + i).EntireColumn.Delete
Cells(8, 1 + i).EntireColumn.Delete
Cells(8, i).EntireColumn.Delete

'======================== IC-EMI-EXO ========================
Do While Workbooks(derivado).Sheets("OPTO_IC_PLAZO").Cells(8, 1 + i) <> ""
 Cells(8, 1 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("OPTO_IC_PLAZO").Cells(8, 1 + i).Value, Workbooks(macro).Sheets("IC").Range("$H$22:$I$36"), 2, False)
i = i + 1
Loop

Cells(8, 2 + i).EntireColumn.Delete
Cells(8, 1 + i).EntireColumn.Delete
Cells(8, i).EntireColumn.Delete

'======================== IC-EMI-WAR ========================
Do While Workbooks(derivado).Sheets("OPTO_IC_PLAZO").Cells(8, 1 + i) <> ""
 Cells(8, 1 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("OPTO_IC_PLAZO").Cells(8, 1 + i).Value, Workbooks(macro).Sheets("IC").Range("$K$22:$L$36"), 2, False)
i = i + 1
Loop

Cells(8, 2 + i).EntireColumn.Delete
Cells(8, 1 + i).EntireColumn.Delete
Cells(8, i).EntireColumn.Delete

'======================== IC-ADQ-EST-EXT ========================
Do While Workbooks(derivado).Sheets("OPTO_IC_PLAZO").Cells(8, 1 + i) <> ""
 Cells(8, 1 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("OPTO_IC_PLAZO").Cells(8, 1 + i).Value, Workbooks(macro).Sheets("IC").Range("$B$41:$C$55"), 2, False)
i = i + 1
Loop

Cells(8, 2 + i).EntireColumn.Delete
Cells(8, 1 + i).EntireColumn.Delete
Cells(8, i).EntireColumn.Delete

'======================== IC-ADQ-EST-ORG ========================
Do While Workbooks(derivado).Sheets("OPTO_IC_PLAZO").Cells(8, 1 + i) <> ""
 Cells(8, 1 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("OPTO_IC_PLAZO").Cells(8, 1 + i).Value, Workbooks(macro).Sheets("IC").Range("$E$41:$F$55"), 2, False)
i = i + 1
Loop

Cells(8, 2 + i).EntireColumn.Delete
Cells(8, 1 + i).EntireColumn.Delete
Cells(8, i).EntireColumn.Delete

'======================== IC-ADQ-EXO ========================
Do While Workbooks(derivado).Sheets("OPTO_IC_PLAZO").Cells(8, 1 + i) <> ""
 Cells(8, 1 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("OPTO_IC_PLAZO").Cells(8, 1 + i).Value, Workbooks(macro).Sheets("IC").Range("$H$41:$I$55"), 2, False)
i = i + 1
Loop

Cells(8, 2 + i).EntireColumn.Delete
Cells(8, 1 + i).EntireColumn.Delete
Cells(8, i).EntireColumn.Delete

'======================== IC-ADQ-WAR ========================
Do While Workbooks(derivado).Sheets("OPTO_IC_PLAZO").Cells(8, 1 + i) <> ""
 Cells(8, 1 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("OPTO_IC_PLAZO").Cells(8, 1 + i).Value, Workbooks(macro).Sheets("IC").Range("$K$41:$L$55"), 2, False)
i = i + 1
Loop

Cells(8, 1 + i).EntireColumn.Delete

Range("B9").Select
Selection.End(xlDown).Select
e = ActiveWindow.RangeSelection.Row
Range("A9:A" & e).FormulaR1C1 = "=TEXT(RC[1], ""dd/mm/aaaa"")"

Range("A9:A" & e).Copy
Range("B9").PasteSpecial Paste:=xlPasteValues
Range("A9:A" & e).ClearContents

'======================== CB-EMI-EST-EXT ========================
Sheets("OPTO_CB_PLAZO").Select
Rows("8").Select
Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
Rows("7").Copy
Range("A8").Select
ActiveSheet.Paste
Range("C8").Select
i = 1

Do While Workbooks(derivado).Sheets("OPTO_CB_PLAZO").Cells(8, 2 + i) <> ""
 Cells(8, 2 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("OPTO_CB_PLAZO").Cells(8, 2 + i).Value, Workbooks(macro).Sheets("CB").Range("$B$22:$C$36"), 2, False)
i = i + 1
Loop

Cells(8, 2 + i).EntireColumn.Delete
Cells(8, 2 + i).EntireColumn.Delete
Cells(8, 1 + i).EntireColumn.Delete

'======================== CB-EMI-EST-ORG ========================
Do While Workbooks(derivado).Sheets("OPTO_CB_PLAZO").Cells(8, 1 + i) <> ""
 Cells(8, 1 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("OPTO_CB_PLAZO").Cells(8, 1 + i).Value, Workbooks(macro).Sheets("CB").Range("$E$22:$F$36"), 2, False)
i = i + 1
Loop

Cells(8, 2 + i).EntireColumn.Delete
Cells(8, 2 + i).EntireColumn.Delete
Cells(8, 1 + i).EntireColumn.Delete

'======================== CB-EMI-EXO ========================
Do While Workbooks(derivado).Sheets("OPTO_CB_PLAZO").Cells(8, 1 + i) <> ""
 Cells(8, 1 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("OPTO_CB_PLAZO").Cells(8, 1 + i).Value, Workbooks(macro).Sheets("CB").Range("$H$22:$I$36"), 2, False)
i = i + 1
Loop

Cells(8, 2 + i).EntireColumn.Delete
Cells(8, 2 + i).EntireColumn.Delete
Cells(8, 1 + i).EntireColumn.Delete

'======================== CB-EMI-WAR ========================
Do While Workbooks(derivado).Sheets("OPTO_CB_PLAZO").Cells(8, 1 + i) <> ""
 Cells(8, 1 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("OPTO_CB_PLAZO").Cells(8, 1 + i).Value, Workbooks(macro).Sheets("CB").Range("$K$22:$L$36"), 2, False)
i = i + 1
Loop

Cells(8, 2 + i).EntireColumn.Delete
Cells(8, 2 + i).EntireColumn.Delete
Cells(8, 1 + i).EntireColumn.Delete

'======================== CB-ADQ-EST-EXT ========================
Do While Workbooks(derivado).Sheets("OPTO_CB_PLAZO").Cells(8, 1 + i) <> ""
 Cells(8, 1 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("OPTO_CB_PLAZO").Cells(8, 1 + i).Value, Workbooks(macro).Sheets("CB").Range("$B$22:$C$36"), 2, False)
i = i + 1
Loop

Cells(8, 2 + i).EntireColumn.Delete
Cells(8, 2 + i).EntireColumn.Delete
Cells(8, 1 + i).EntireColumn.Delete

'======================== CB-ADQ-EST-ORG ========================
Do While Workbooks(derivado).Sheets("OPTO_CB_PLAZO").Cells(8, 1 + i) <> ""
 Cells(8, 1 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("OPTO_CB_PLAZO").Cells(8, 1 + i).Value, Workbooks(macro).Sheets("CB").Range("$E$22:$F$36"), 2, False)
i = i + 1
Loop

Cells(8, 2 + i).EntireColumn.Delete
Cells(8, 2 + i).EntireColumn.Delete
Cells(8, 1 + i).EntireColumn.Delete

'======================== CB-ADQ-EXO ========================
Do While Workbooks(derivado).Sheets("OPTO_CB_PLAZO").Cells(8, 1 + i) <> ""
 Cells(8, 1 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("OPTO_CB_PLAZO").Cells(8, 1 + i).Value, Workbooks(macro).Sheets("CB").Range("$H$22:$I$36"), 2, False)
i = i + 1
Loop

Cells(8, 2 + i).EntireColumn.Delete
Cells(8, 2 + i).EntireColumn.Delete
Cells(8, 1 + i).EntireColumn.Delete

'======================== CB-ADQ-WAR ========================
Do While Workbooks(derivado).Sheets("OPTO_CB_PLAZO").Cells(8, 1 + i) <> ""
 Cells(8, 1 + i).Select
      ActiveCell = Application.VLookup(Workbooks(derivado).Sheets("OPTO_CB_PLAZO").Cells(8, 1 + i).Value, Workbooks(macro).Sheets("CB").Range("$K$22:$L$36"), 2, False)
i = i + 1
Loop
Cells(8, 1 + i).EntireColumn.Delete

Range("B9").Select
Selection.End(xlDown).Select
e = ActiveWindow.RangeSelection.Row
Range("A9:A" & e).FormulaR1C1 = "=TEXT(RC[1], ""dd/mm/aaaa"")"

Range("A9:A" & e).Copy
Range("B9").PasteSpecial Paste:=xlPasteValues
Range("A9:A" & e).ClearContents

MsgBox ("Listo")
End Sub

