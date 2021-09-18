Sub automate() 

    Dim wsSrc As Worksheet, wsDest As Worksheet, wsSrcsn As Worksheet, wsDestsn As Worksheet, wsDestlc As Worksheet 

    Set wsSrc = Workbooks("1_License_usage_report_Aug11_to_Aug26th_with_periodic_review_count.xlsx").Worksheets("Sheet2") 

    Set wsDest = Workbooks("User Report for License Exchange and Deactivation.xlsm").Worksheets("Ignite") 

    Set wsSrcsn = Workbooks("sys_user.xlsx").Worksheets("Page 1") 

    Set wsDestsn = Workbooks("User Report for License Exchange and Deactivation.xlsm").Worksheets("SNOW") 

    Set wsDestlc = Workbooks("User Report for License Exchange and Deactivation.xlsm").Worksheets("Licenses") 

     

    'i. Clear wsDest and wsDestsn 

    'ii. Copy and paste values from wsSrcsn to wsDestsn 

    'iii. Copy and paste values from wsSrc to wsDest 

    'iv. Add Formulas to wsDest 

    'v. Clear Licenses 

    'vi. Copy from wsDest to wsDestlc 

     

    'i 

    wsDest.Activate 

    wsDest.Range("A2:Y" & wsDest.Rows.Count).ClearContents 

    wsDestsn.Activate 

    wsDestsn.Range("A2:N" & wsDestsn.Rows.Count).ClearContents 

     

    'ii 

    wsSrcsn.Activate 

    Dim cntsn As Long 

    cntsn = Cells(wsSrcsn.Rows.Count, "A").End(xlUp).Row 

    wsSrcsn.Range("A2:N" & cntsn).Copy 

    wsDestsn.Activate 

    wsDestsn.Range("A2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats 

     

    'iii 

    wsSrc.Activate 

    Dim cnt As Long 

    cnt = Cells(wsSrc.Rows.Count, "A").End(xlUp).Row 

    wsSrc.Range("A2:L" & cnt).Copy 

    wsDest.Activate 

    wsDest.Range("F2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats 

    wsSrc.Activate 

    wsSrc.Range("O2:V" & cnt).Copy 

    wsDest.Activate 

    wsDest.Range("R2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats 

     

     

    'iv. 

    Dim i As Integer 

    wsDest.Activate 

    For i = 2 To cnt 

        wsDest.Range("A" & i).Formula = "=INDEX(SNOW!$A:$ZZ,MATCH(Ignite!H" & i & ",SNOW!B:B,0),13)" 

        wsDest.Range("B" & i).Formula = "=INDEX(SNOW!$A:$ZZ,MATCH(Ignite!H" & i & ",SNOW!B:B,0),12)" 

        wsDest.Range("C" & i).Formula = "=INDEX(SNOW!$A:$ZZ,MATCH(Ignite!H" & i & ",SNOW!B:B,0),9)" 

        wsDest.Range("D" & i).Formula = "=INDEX(SNOW!$A:$ZZ,MATCH(Ignite!H" & i & ",SNOW!B:B,0),7)" 

        wsDest.Range("E" & i).Formula = "=INDEX(SNOW!$A:$ZZ,MATCH(Ignite!H" & i & ",SNOW!B:B,0),3)" 

    Next i 

    'MsgBox "Done" 

     

    'v. 

    wsDestlc.Activate 

    wsDestlc.Range("A2:T" & wsDestlc.Rows.Count).ClearContents 

     

    'vi. 

    wsDest.Activate 

    Dim cntign As Long 

    cntign = Cells(wsDest.Rows.Count, "F").End(xlUp).Row 

     

    wsDest.Range("C2:D" & cntign).Copy 

    wsDestlc.Activate 

    wsDestlc.Range("A2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats 

     

    wsDest.Activate 

    wsDest.Range("F2:F" & cntign).Copy 

    wsDestlc.Activate 

    wsDestlc.Range("C2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats 

     

    wsDest.Activate 

    wsDest.Range("H2:I" & cntign).Copy 

    wsDestlc.Activate 

    wsDestlc.Range("D2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats 

     

    wsDest.Activate 

    wsDest.Range("K2:Y" & cntign).Copy 

    wsDestlc.Activate 

    wsDestlc.Range("F2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats 

     

    Application.CutCopyMode = False 

    MsgBox "Done" 

     

End Sub 

 

 