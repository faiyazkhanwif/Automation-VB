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

    'vii. Filter and remove redundant from wsDestlc data based on status 

    'viii. Copy to new sheet with custom name reflecting date (*2) 

    'ix. Filter data based on Pivot requirements (*2) 

    'x. Format table (*2) 

     

    'i.------- 

    wsDest.Activate 

    wsDest.Range("A2:Y" & wsDest.Rows.Count).ClearContents 

    wsDestsn.Activate 

    wsDestsn.Range("A2:N" & wsDestsn.Rows.Count).ClearContents 

     

    'ii.------- 

    wsSrcsn.Activate 

    Dim cntsn As Long 

    cntsn = Cells(wsSrcsn.Rows.Count, "A").End(xlUp).Row 

    wsSrcsn.Range("A2:N" & cntsn).Copy 

    wsDestsn.Activate 

    wsDestsn.Range("A2").PasteSpecial Paste:=xlPasteValuesAndNumberFormats 

     

    'iii.------- 

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

     

     

    'iv.------- 

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

     

    'v.------- 

    wsDestlc.Activate 

    wsDestlc.Range("A2:T" & wsDestlc.Rows.Count).ClearContents 

     

    'vi.------- 

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

     

     

    'vii.------- 

    wsDestlc.Activate 

    Dim lastRow As Long 

    lastRow = Cells(wsDestlc.Rows.Count, "C").End(xlUp).Row 

     

    With wsDestlc 

        .AutoFilterMode = False 

        With .Range("G1:G" & lastRow) 

            .AutoFilter Field:=1, Criteria1:="Inactive" 

            .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete 

        End With 

        .AutoFilterMode = False 

    End With 

    Application.CutCopyMode = False 

     

     

    'viii.------- 

    Dim dt As String 

    dt = Format(Date, "yyyy/mm/dd") 

     

    Dim s1 As String 

    s1 = Replace(dt, "/", "") 

     

    wsDestsn.Activate 

    Dim sheetname As String 

    sheetname = ">45days_Reminder_" & s1 

    Sheets.Add(After:=wsDestsn).Name = sheetname 

     

    Dim nsheet As Worksheet 

    Set nsheet = Workbooks("User Report for License Exchange and Deactivation.xlsm").Worksheets(sheetname) 

 

    wsDestlc.Activate 

    Dim cntnew As Long 

    cntnew = Cells(wsDestlc.Rows.Count, "C").End(xlUp).Row 

    wsDestlc.Range("A1:T" & cntnew).Copy 

    nsheet.Activate 

    nsheet.Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats 

     

    Application.CutCopyMode = False 

     

     

    'ix.------- 

    ' user created remove n 

    nsheet.Activate 

    Dim lrforuc As Long 

    lrforuc = Cells(nsheet.Rows.Count, "C").End(xlUp).Row 

     

    With nsheet 

        .AutoFilterMode = False 

        With .Range("J1:J" & lrforuc) 

            .AutoFilter Field:=1, Criteria1:="n" 

            .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete 

        End With 

        .AutoFilterMode = False 

    End With 

    Application.CutCopyMode = False 

     

    ' doc owner remove y 

    nsheet.Activate 

    Dim lrfordo As Long 

    lrfordo = Cells(nsheet.Rows.Count, "C").End(xlUp).Row 

     

    With nsheet 

        .AutoFilterMode = False 

        With .Range("L1:L" & lrfordo) 

            .AutoFilter Field:=1, Criteria1:="y" 

            .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete 

        End With 

        .AutoFilterMode = False 

    End With 

    Application.CutCopyMode = False 

     

    ' last login remove n 

    nsheet.Activate 

    Dim lrforll As Long 

    lrforll = Cells(nsheet.Rows.Count, "C").End(xlUp).Row 

     

    With nsheet 

        .AutoFilterMode = False 

        With .Range("K1:K" & lrforll) 

            .AutoFilter Field:=1, Criteria1:="n" 

            .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete 

        End With 

        .AutoFilterMode = False 

    End With 

    Application.CutCopyMode = False 

     

    ' Default group remove dba,labeling,PCS,regulatory 

     

    'dba 

    nsheet.Activate 

    Dim lrfordgdba As Long 

    lrfordgdba = Cells(nsheet.Rows.Count, "C").End(xlUp).Row 

     

    With nsheet 

        .AutoFilterMode = False 

        With .Range("E1:E" & lrfordgdba) 

            .AutoFilter Field:=1, Criteria1:="dba" 

            .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete 

        End With 

        .AutoFilterMode = False 

    End With 

    Application.CutCopyMode = False 

     

    'Labeling 

    nsheet.Activate 

    Dim lrfordglb As Long 

    lrfordglb = Cells(nsheet.Rows.Count, "C").End(xlUp).Row 

     

    With nsheet 

        .AutoFilterMode = False 

        With .Range("E1:E" & lrfordglb) 

            .AutoFilter Field:=1, Criteria1:="Labeling" 

            .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete 

        End With 

        .AutoFilterMode = False 

    End With 

    Application.CutCopyMode = False 

     

    'PCS 

    nsheet.Activate 

    Dim lrfordgpcs As Long 

    lrfordgpcs = Cells(nsheet.Rows.Count, "C").End(xlUp).Row 

     

    With nsheet 

        .AutoFilterMode = False 

        With .Range("E1:E" & lrfordgpcs) 

            .AutoFilter Field:=1, Criteria1:="PCS" 

            .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete 

        End With 

        .AutoFilterMode = False 

    End With 

    Application.CutCopyMode = False 

     

    'Regulatory 

    nsheet.Activate 

    Dim lrfordgrg As Long 

    lrfordgrg = Cells(nsheet.Rows.Count, "C").End(xlUp).Row 

     

    With nsheet 

        .AutoFilterMode = False 

        With .Range("E1:E" & lrfordgrg) 

            .AutoFilter Field:=1, Criteria1:="Regulatory" 

            .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete 

        End With 

        .AutoFilterMode = False 

    End With 

    Application.CutCopyMode = False 

     

    ' Title remove dir,vp 

    'dir remove 

    nsheet.Activate 

    Dim lrfortt As Long 

    lrfortt = Cells(nsheet.Rows.Count, "C").End(xlUp).Row 

     

    Dim rngtt As Range 

    Dim postt As Integer 

    Set rngtt = nsheet.Range("A1:A" & lrfortt) 

     

    For i = rngtt.Cells.Count To 1 Step -1 

        postt = InStr(LCase(rngtt.Item(i).Value), LCase("dir")) 

        If postt > 0 Then 

            rngtt.Item(i).EntireRow.Delete 

        End If 

    Next i 

    Application.CutCopyMode = False 

    'vp remove 

    Dim lrfortt2 As Long 

    lrfortt2 = Cells(nsheet.Rows.Count, "C").End(xlUp).Row 

     

    Dim rngtt2 As Range 

    Dim postt2 As Integer 

    Set rngtt2 = nsheet.Range("A1:A" & lrfortt2) 

     

    For i = rngtt2.Cells.Count To 1 Step -1 

        postt2 = InStr(LCase(rngtt2.Item(i).Value), LCase("VP")) 

        If postt2 > 0 Then 

            rngtt2.Item(i).EntireRow.Delete 

        End If 

    Next i 

    Application.CutCopyMode = False 

     

    'Name remove data migration 

    Dim lrfornm As Long 

    lrfornm = Cells(nsheet.Rows.Count, "C").End(xlUp).Row 

     

    Dim rngnm As Range 

    Dim posnm As Integer 

    Set rngnm = nsheet.Range("C1:C" & lrfornm) 

     

    For i = rngnm.Cells.Count To 1 Step -1 

        posnm = InStr(LCase(rngnm.Item(i).Value), LCase("data migration")) 

        If posnm > 0 Then 

            rngnm.Item(i).EntireRow.Delete 

        End If 

    Next i 

    Application.CutCopyMode = False 

     

     

    'x 

    Dim tbl As ListObject 

    Dim rng As Range 

    Dim lrfortable As Long 

    lrfortable = Cells(nsheet.Rows.Count, "C").End(xlUp).Row 

    Set rng = nsheet.Range("A1:T" & lrfortable) 

 

    Set tbl = nsheet.ListObjects.Add(xlSrcRange, rng, , xlYes) 

    tbl.TableStyle = "TableStyleMedium2" 

     

    Application.CutCopyMode = False 

    MsgBox "Done" 

     

End Sub 

 

 

 

 