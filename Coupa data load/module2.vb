Sub automate2() 

    Dim wsmaster As Worksheet, wslulist As Worksheet, wsulist As Worksheet 

    Set wsmaster = Workbooks("Cost Center Master File.xlsx").Worksheets("Sheet1") 

    Set wslulist = Workbooks("lookup_value_list.xlsx").Worksheets("sheet1") 

    Set wsulist = Workbooks("user_list.xlsx").Worksheets("sheet1") 

     

    Dim nsheet As Worksheet 

    Set nsheet = Workbooks("Cost Center Master File.xlsx").Worksheets("File") 

     

    'i. Delete unnecessary columns from nsheet 

    nsheet.Activate 

    nsheet.Range("K1").Interior.Color = RGB(61, 245, 98) 

    nsheet.Range("L1").Interior.Color = RGB(61, 245, 98) 

    nsheet.Range("K1") = "Finance Partner" 

    nsheet.Range("L1") = "Business Primary Owner" 

     

    Dim cnt As Long 

    cnt = Cells(nsheet.Rows.Count, "A").End(xlUp).Row 

     

    nsheet.Range("F2:F" & cnt).Copy 

    nsheet.Range("K2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats 

     

    nsheet.Range("J2:J" & cnt).Copy 

    nsheet.Range("L2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats 

     

    nsheet.Range("C:C,D:D,E:E,F:F,G:G,H:H,I:I,J:J").Delete 

     

    [A:A].Select 

    With Selection 

        .NumberFormat = "General" 

        .Value = .Value 

    End With 

    '[C:C].Select 

    'With Selection 

    '    .NumberFormat = "General" 

    '    .Value = .Value 

    'End With 

    '[D:D].Select 

    'With Selection 

    '    .NumberFormat = "General" 

    '    .Value = .Value 

    'End With 

     

    Application.CutCopyMode = False 

     

    'ii. Copy lookup value sheet inside the master file in a new sheet 

    Dim str As String 

    str = "LUsheet" 

 

    nsheet.Activate 

    Sheets.Add(After:=nsheet).Name = str 

     

    Dim lusheet As Worksheet 

    Set lusheet = Workbooks("Cost Center Master File.xlsx").Worksheets(str) 

     

    wslulist.Activate 

    Dim lucnt As Long 

    lucnt = Cells(wslulist.Rows.Count, "A").End(xlUp).Row 

     

    wslulist.Range("A1:I" & lucnt).Copy 

     

    lusheet.Activate 

    lusheet.Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats 

     

    Application.CutCopyMode = False 

     

    'iii. Prepare nsheet for vlookup 

    Dim ncnt As Long 

    ncnt = Cells(nsheet.Rows.Count, "A").End(xlUp).Row 

     

    nsheet.Range("C1:C" & ncnt).Copy 

    nsheet.Range("E1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats 

    nsheet.Range("C:C").Delete 

    nsheet.Range("D1").Interior.Color = RGB(61, 245, 98) 

     

    lusheet.Activate 

    lusheet.Range("I1").EntireColumn.Insert 

     

    [F:F].Select 

    With Selection 

        .NumberFormat = "General" 

        .Value = .Value 

    End With 

     

    'iv. Do vlookup in Lusheet 

    Dim cntnew As Long 

    cntnew = Cells(lusheet.Rows.Count, "A").End(xlUp).Row 

     

    For i = 2 To cntnew 

        lusheet.Range("I" & i).Formula = "=VLOOKUP(F" & i & ",'File'!A:C,3,False)" 

        lusheet.Range("K" & i).Formula = "=VLOOKUP(F" & i & ",'File'!A:D,4,False)" 

    Next i 

     

    lusheet.Range("I1") = "Primary Business Owner (new)" 

    lusheet.Range("K1") = "Finance Partner (new)" 

     

    'v. Filter out the primary business owner (New) is #N/A AND the primary business owner is blank – Replace the #N/A with blank. 

    With lusheet 

        .AutoFilterMode = False 

        With .Range("H1:I" & cntnew) 

            .AutoFilter Field:=1, Criteria1:="" 

            .AutoFilter Field:=2, Criteria1:="=#N/A" 

        End With 

        '.AutoFilterMode = False 

    End With 

    Application.CutCopyMode = False 

     

    Dim cl As Range, borng As Range 

    Set borng = lusheet.Range("I2:I" & cntnew) 

    For Each cl In borng.SpecialCells(xlCellTypeVisible) 

        cl = "" 

    Next cl 

     

     

    'vi. Filter out the finance partner (New) #N/A AND the finance partner is blank – Replace the #N/A with blank. 

    With lusheet 

        .AutoFilterMode = False 

        With .Range("J1:K" & cntnew) 

            .AutoFilter Field:=1, Criteria1:="" 

            .AutoFilter Field:=2, Criteria1:="=#N/A" 

        End With 

        '.AutoFilterMode = False 

    End With 

    Application.CutCopyMode = False 

     

    Dim cl1 As Range, fprng As Range 

    Set fprng = lusheet.Range("K2:K" & cntnew) 

    For Each cl1 In fprng.SpecialCells(xlCellTypeVisible) 

        cl1 = "" 

    Next cl1 

     

    lusheet.AutoFilterMode = False 

     

    MsgBox ("Please complete the last four steps manually") 

     

End Sub 