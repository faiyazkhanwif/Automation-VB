Sub Macro2() 

' 

' Macro2 Macro 

' 

 

' 

    'Bug fix 

    Dim wsfin As Worksheet 

    Set wsfin = Workbooks("How to load HCP contracts to JDE F4801.xlsm").Worksheets("Load Template") 

    Dim last_row As Long 

    last_row = wsfin.Cells(Rows.Count, 1).End(xlUp).Row 

    'MsgBox (last_row) 

     

    If last_row = 2 Then 

        Rows(2).Copy Rows(3) 

        'MsgBox ("added for bug fix") 

        'Rows(3).EntireRow.Delete 

    End If 

     

    Range("O2").Select 

    ActiveCell.FormulaR1C1 = "=LEFT(RC[-13],30)" 

    Range("O2").Select 

    Selection.AutoFill Destination:=Range("O2:O" & Range("B" & Rows.Count).End(xlUp).Row) 

    Range(Selection, Selection.End(xlDown)).Select 

    Range("P2").Select 

    ActiveCell.FormulaR1C1 = _ 

        "=VLOOKUP(RC[-10],'Map ECLM BU to JDE BU'!C[-15]:C[-14],2,FALSE)" 

    Range("P2").Select 

    Selection.AutoFill Destination:=Range("P2:P" & Range("B" & Rows.Count).End(xlUp).Row) 

    Range(Selection, Selection.End(xlDown)).Select 

    Selection.Copy 

    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _ 

        :=False, Transpose:=False 

    Application.CutCopyMode = False 

    Selection.Copy 

    Range("F2").Select 

    ActiveSheet.Paste 

    Application.CutCopyMode = False 

    With Selection 

        .HorizontalAlignment = xlLeft 

        .VerticalAlignment = xlBottom 

        .WrapText = False 

        .Orientation = 0 

        .AddIndent = False 

        .IndentLevel = 0 

        .ShrinkToFit = False 

        .ReadingOrder = xlContext 

        .MergeCells = False 

    End With 

    Range("O2").Select 

    Range(Selection, Selection.End(xlDown)).Select 

    Selection.Copy 

    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _ 

        :=False, Transpose:=False 

    Application.CutCopyMode = False 

    Selection.Copy 

    Range("B2").Select 

    ActiveSheet.Paste 

    Range("C2:E2").Select 

    Range(Selection, Selection.End(xlDown)).Select 

    Application.CutCopyMode = False 

    Selection.ClearContents 

    Range("C2").Select 

    ActiveCell.FormulaR1C1 = "A" 

    Range("D2").Select 

    ActiveCell.FormulaR1C1 = "G" 

    Range("E2").Select 

    With Selection.Interior 

        .Pattern = xlNone 

        .TintAndShade = 0 

        .PatternTintAndShade = 0 

    End With 

    Selection.Borders(xlDiagonalDown).LineStyle = xlNone 

    Selection.Borders(xlDiagonalUp).LineStyle = xlNone 

    Selection.Borders(xlEdgeLeft).LineStyle = xlNone 

    Selection.Borders(xlEdgeTop).LineStyle = xlNone 

    Selection.Borders(xlEdgeBottom).LineStyle = xlNone 

    Selection.Borders(xlEdgeRight).LineStyle = xlNone 

    Selection.Borders(xlInsideVertical).LineStyle = xlNone 

    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone 

    ActiveCell.FormulaR1C1 = "H" 

    Range("C2").Select 

    Selection.AutoFill Destination:=Range("C2:C" & Range("B" & Rows.Count).End(xlUp).Row) 

    Range(Selection, Selection.End(xlDown)).Select 

    Range("D2").Select 

    Selection.AutoFill Destination:=Range("D2:D" & Range("B" & Rows.Count).End(xlUp).Row) 

    Range(Selection, Selection.End(xlDown)).Select 

     Range("E2").Select 

    Selection.AutoFill Destination:=Range("E2:E" & Range("B" & Rows.Count).End(xlUp).Row) 

    Range(Selection, Selection.End(xlDown)).Select 

    Range("H2:J2").Select 

    Range(Selection, Selection.End(xlDown)).Select 

    Application.CutCopyMode = False 

    Selection.ClearContents 

    Range("H2").Select 

    Application.CutCopyMode = False 

    ActiveCell.FormulaR1C1 = "=TODAY()-3" 

    Range("H2").Select 

    Selection.AutoFill Destination:=Range("H2:H" & Range("B" & Rows.Count).End(xlUp).Row) 

    Range(Selection, Selection.End(xlDown)).Select 

    Range("I2").Select 

    ActiveCell.FormulaR1C1 = "=TODAY()-3" 

    Selection.AutoFill Destination:=Range("I2:I" & Range("B" & Rows.Count).End(xlUp).Row) 

    Range(Selection, Selection.End(xlDown)).Select 

    Range("J2").Select 

    ActiveCell.FormulaR1C1 = "=TODAY()-3" 

    Selection.AutoFill Destination:=Range("J2:J" & Range("B" & Rows.Count).End(xlUp).Row) 

    Range(Selection, Selection.End(xlDown)).Select 

    Range("O2:P2").Select 

    Range(Selection, Selection.End(xlDown)).Select 

    Application.CutCopyMode = False 

    Selection.ClearContents 

  '  Sheets("Instructions").Select 

  '  Range("B76").Select 

  '  Selection.Hyperlinks(1).Follow NewWindow:=False, AddHistory:=True 

  '  Sheets("Load Template").Select 

   

   'Bug fix 

     If last_row = 2 Then 

        'Rows(2).Copy Rows(3) 

        'MsgBox ("added for bug fix") 

        Rows(3).EntireRow.Delete 

     End If 

      

End Sub 

 

 