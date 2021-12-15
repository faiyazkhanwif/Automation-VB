Sub automate() 

    Dim wsmaster As Worksheet, wslulist As Worksheet, wsulist As Worksheet 

    Set wsmaster = Workbooks("Cost Center Master File.xlsx").Worksheets("Sheet1") 

    Set wslulist = Workbooks("lookup_value_list.xlsx").Worksheets("sheet1") 

    Set wsulist = Workbooks("user_list.xlsx").Worksheets("sheet1") 

     

    'i. Create a new sheet (“File”) in wsmaster. 

    Dim st As String 

    st = "File" 

 

    wsmaster.Activate 

    Sheets.Add(After:=wsmaster).Name = st 

     

    Dim nsheet As Worksheet 

    Set nsheet = Workbooks("Cost Center Master File.xlsx").Worksheets(st) 

     

    'ii. Copy A,B,AI,AJ,AK,AL columns and paste in to new sheet. 

    wsmaster.Activate 

    Dim cnt As Long 

    cnt = Cells(wsmaster.Rows.Count, "A").End(xlUp).Row 

     

    wsmaster.Range("A1:B" & cnt).Copy 

    nsheet.Activate 

    nsheet.Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats 

     

    wsmaster.Activate 

    wsmaster.Range("AI1:AL" & cnt).Copy 

    nsheet.Activate 

    nsheet.Range("C1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats 

     

    'iii. Filter and delete blank data from nsheet (Based on column D and F) 

    nsheet.Activate 

    Dim lastRow As Long 

    lastRow = Cells(nsheet.Rows.Count, "A").End(xlUp).Row 

     

    With nsheet 

        .AutoFilterMode = False 

        With .Range("D1:F" & lastRow) 

            .AutoFilter Field:=1, Criteria1:="" 

            .AutoFilter Field:=3, Criteria1:="" 

            .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete 

        End With 

        .AutoFilterMode = False 

    End With 

    Application.CutCopyMode = False 

     

     

    'iv. Create a new sheet (“User”) in Main file. 

    Dim st1 As String 

    st1 = "User" 

 

    nsheet.Activate 

    Sheets.Add(After:=nsheet).Name = st1 

     

    Dim nsheet1 As Worksheet 

    Set nsheet1 = Workbooks("Cost Center Master File.xlsx").Worksheets(st1) 

     

     

    'v. Copy all columns and paste in to nsheet1. 

    wsulist.Activate 

    Dim cnt1 As Long 

    cnt1 = Cells(wsulist.Rows.Count, "A").End(xlUp).Row 

     

    wsulist.Range("A1:F" & cnt1).Copy 

    nsheet1.Activate 

    nsheet1.Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats 

     

     

    'vi. Insert blank columns in File before vlookup 

    nsheet.Activate 

    nsheet.Range("D1").EntireColumn.Insert 

    nsheet.Range("F1").EntireColumn.Insert 

    nsheet.Range("H1").EntireColumn.Insert 

    nsheet.Range("J1").EntireColumn.Insert 

     

    'vii. Do vlookup and populate column D and F in File 

    Dim i As Integer 

     

    nsheet.Activate 

    Dim cnt2 As Long 

    cnt2 = Cells(nsheet.Rows.Count, "A").End(xlUp).Row 

 

    For i = 2 To cnt2 

        nsheet.Range("D" & i).Formula = "=VLOOKUP(C" & i & ",'User'!A:A,1,False)" 

        nsheet.Range("F" & i).Formula = "=VLOOKUP(E" & i & ",'User'!B:B,1,False)" 

        nsheet.Range("H" & i).Formula = "=VLOOKUP(G" & i & ",'User'!A:A,1,False)" 

        nsheet.Range("J" & i).Formula = "=VLOOKUP(I" & i & ",'User'!B:B,1,False)" 

    Next i 

 

    nsheet.Range("D1").Interior.Color = RGB(61, 245, 98) 

    nsheet.Range("F1").Interior.Color = RGB(61, 245, 98) 

    nsheet.Range("H1").Interior.Color = RGB(61, 245, 98) 

    nsheet.Range("J1").Interior.Color = RGB(61, 245, 98) 

 

    nsheet.Range("D1") = "Vlookup with Full Name" 

    nsheet.Range("F1") = "Vlookup with Employee Number" 

    nsheet.Range("H1") = "Vlookup with Full Name" 

    nsheet.Range("J1") = "Vlookup with Employee Number" 

         

    Application.CutCopyMode = False 

     

     

    'viii. Filter out the N/A value for D, F, H and J columns, remove the N/A to blank. 

    '   D 

    Dim rng1 As Range 

    Set rng1 = nsheet.Range("D1:D" & cnt2) 

     

    'Loop all the cells in range 

    For Each cell In rng1 

        If Application.WorksheetFunction.IsNA(cell) Then 

            cell.Value = "" 

        End If 

    Next 

    Application.CutCopyMode = False 

    '   F 

    Dim rng2 As Range 

    Set rng2 = nsheet.Range("F1:F" & cnt2) 

     

    'Loop all the cells in range 

    For Each cell In rng2 

        If Application.WorksheetFunction.IsNA(cell) Then 

            cell.Value = "" 

        End If 

    Next 

    Application.CutCopyMode = False 

    '   H 

    Dim rng3 As Range 

    Set rng3 = nsheet.Range("H1:H" & cnt2) 

     

    'Loop all the cells in range 

    For Each cell In rng3 

        If Application.WorksheetFunction.IsNA(cell) Then 

            cell.Value = "" 

        End If 

    Next 

    Application.CutCopyMode = False 

    '   J 

    Dim rng4 As Range 

    Set rng4 = nsheet.Range("J1:J" & cnt2) 

     

    'Loop all the cells in range 

    For Each cell In rng4 

        If Application.WorksheetFunction.IsNA(cell) Then 

            cell.Value = "" 

        End If 

    Next 

    Application.CutCopyMode = False 

     

    MsgBox ("If employee number is blank but the name is there, then please find out the employee number base on the name and fill in into the 'vlookup with employee number' columns") 

     

End Sub 