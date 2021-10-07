Sub automate() 

    Dim wsSrc As Worksheet, wsDest As Worksheet, wsSrcsn As Worksheet, wsDestsn As Worksheet, wsDestlc As Worksheet 

    Set wsSrc = Workbooks("License_usage_report.xlsx").Worksheets("License_usage_report") 

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

        If InStr(1, rngtt.Item(i).Text, "#N/A") = 0 Then 'Data type exception handler 

            postt = InStr(LCase(rngtt.Item(i).Value), LCase("dir")) 

            If postt > 0 Then 

                rngtt.Item(i).EntireRow.Delete 

            End If 

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

        If InStr(1, rngtt.Item(i).Text, "#N/A") = 0 Then 'Data type exception handler 

            postt2 = InStr(LCase(rngtt2.Item(i).Value), LCase("VP")) 

            If postt2 > 0 Then 

                rngtt2.Item(i).EntireRow.Delete 

            End If 

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

     

     

    'Name remove shr-ignite-dynatrace 

    Dim lrfornm1 As Long 

    lrfornm1 = Cells(nsheet.Rows.Count, "C").End(xlUp).Row 

     

    Dim rngnm1 As Range 

    Dim posnm1 As Integer 

    Set rngnm1 = nsheet.Range("C1:C" & lrfornm1) 

     

    For i = rngnm1.Cells.Count To 1 Step -1 

        posnm1 = InStr(LCase(rngnm1.Item(i).Value), LCase("shr-ignite-dynatrace")) 

        If posnm1 > 0 Then 

            rngnm1.Item(i).EntireRow.Delete 

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

     

     

     

    'viii, ix, x for second report 

     

    'viii(*2)-------- 

    Dim dta2c As String 

    dta2c = Format(Date, "yyyy/mm/dd") 

     

    Dim sa2c As String 

    sa2c = Replace(dta2c, "/", "") 

     

    Dim wsDest45 As Worksheet 

    Set wsDest45 = Workbooks("User Report for License Exchange and Deactivation.xlsm").Worksheets("SNOW") 

    wsDest45.Activate 

     

    Dim sheetnamea2c As String 

    sheetnamea2c = "A2C_Reminder_" & sa2c 

    Sheets.Add(After:=wsDest45).Name = sheetnamea2c 

     

    Dim nsheeta2c As Worksheet 

    Set nsheeta2c = Workbooks("User Report for License Exchange and Deactivation.xlsm").Worksheets(sheetnamea2c) 

 

    wsDestlc.Activate 

    Dim cntnewa2c As Long 

    cntnewa2c = Cells(wsDestlc.Rows.Count, "C").End(xlUp).Row 

    wsDestlc.Range("A1:T" & cntnewa2c).Copy 

    nsheeta2c.Activate 

    nsheeta2c.Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats 

     

    Application.CutCopyMode = False 

     

     

    'ix.(*2) 

    'license level 

    nsheeta2c.Activate 

    Dim lrforlla2c As Long 

    lrforlla2c = Cells(nsheeta2c.Rows.Count, "C").End(xlUp).Row 

     

    With nsheeta2c 

        .AutoFilterMode = False 

        With .Range("F1:F" & lrforlla2c) 

            .AutoFilter Field:=1, Criteria1:="<>Author" 

            .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete 

        End With 

        .AutoFilterMode = False 

    End With 

    Application.CutCopyMode = False 

     

    'status 

    nsheeta2c.Activate 

    Dim lrforsa2c As Long 

    lrforsa2c = Cells(nsheeta2c.Rows.Count, "C").End(xlUp).Row 

     

    With nsheeta2c 

        .AutoFilterMode = False 

        With .Range("G1:G" & lrforsa2c) 

            .AutoFilter Field:=1, Criteria1:="<>Active" 

            .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete 

        End With 

        .AutoFilterMode = False 

    End With 

    Application.CutCopyMode = False 

     

    'user created 

    nsheeta2c.Activate 

    Dim lrforuca2c As Long 

    lrforuca2c = Cells(nsheeta2c.Rows.Count, "C").End(xlUp).Row 

     

    With nsheeta2c 

        .AutoFilterMode = False 

        With .Range("J1:J" & lrforuca2c) 

            .AutoFilter Field:=1, Criteria1:="<>y" 

            .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete 

        End With 

        .AutoFilterMode = False 

    End With 

    Application.CutCopyMode = False 

     

     

    'doc owner 

    nsheeta2c.Activate 

    Dim lrfordoa2c As Long 

    lrfordoa2c = Cells(nsheeta2c.Rows.Count, "C").End(xlUp).Row 

     

    With nsheeta2c 

        .AutoFilterMode = False 

        With .Range("L1:L" & lrfordoa2c) 

            .AutoFilter Field:=1, Criteria1:="<>n" 

            .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete 

        End With 

        .AutoFilterMode = False 

    End With 

    Application.CutCopyMode = False 

     

     

    'part rev 

    nsheeta2c.Activate 

    Dim lrforpra2c As Long 

    lrforpra2c = Cells(nsheeta2c.Rows.Count, "C").End(xlUp).Row 

     

    With nsheeta2c 

        .AutoFilterMode = False 

        With .Range("M1:M" & lrforpra2c) 

            .AutoFilter Field:=1, Criteria1:="<>0" 

            .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete 

        End With 

        .AutoFilterMode = False 

    End With 

    Application.CutCopyMode = False 

     

    'doc rev 

    nsheeta2c.Activate 

    Dim lrfordra2c As Long 

    lrfordra2c = Cells(nsheeta2c.Rows.Count, "C").End(xlUp).Row 

     

    With nsheeta2c 

        .AutoFilterMode = False 

        With .Range("N1:N" & lrfordra2c) 

            .AutoFilter Field:=1, Criteria1:="<>0" 

            .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete 

        End With 

        .AutoFilterMode = False 

    End With 

    Application.CutCopyMode = False 

     

    'ECR rev 

    nsheeta2c.Activate 

    Dim lrforecrra2c As Long 

    lrforecrra2c = Cells(nsheeta2c.Rows.Count, "C").End(xlUp).Row 

     

    With nsheeta2c 

        .AutoFilterMode = False 

        With .Range("P1:P" & lrforecrra2c) 

            .AutoFilter Field:=1, Criteria1:="<>0" 

            .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete 

        End With 

        .AutoFilterMode = False 

    End With 

    Application.CutCopyMode = False 

     

    'ECN rev 

    nsheeta2c.Activate 

    Dim lrforecnra2c As Long 

    lrforecnra2c = Cells(nsheeta2c.Rows.Count, "C").End(xlUp).Row 

     

    With nsheeta2c 

        .AutoFilterMode = False 

        With .Range("Q1:Q" & lrforecnra2c) 

            .AutoFilter Field:=1, Criteria1:="<>0" 

            .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete 

        End With 

        .AutoFilterMode = False 

    End With 

    Application.CutCopyMode = False 

     

    'DHF rev 

    nsheeta2c.Activate 

    Dim lrfordhfra2c As Long 

    lrfordhfra2c = Cells(nsheeta2c.Rows.Count, "C").End(xlUp).Row 

     

    With nsheeta2c 

        .AutoFilterMode = False 

        With .Range("R1:R" & lrfordhfra2c) 

            .AutoFilter Field:=1, Criteria1:="<>0" 

            .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete 

        End With 

        .AutoFilterMode = False 

    End With 

    Application.CutCopyMode = False 

     

    'datasets 

    nsheeta2c.Activate 

    Dim lrfordatara2c As Long 

    lrfordatara2c = Cells(nsheeta2c.Rows.Count, "C").End(xlUp).Row 

     

    With nsheeta2c 

        .AutoFilterMode = False 

        With .Range("S1:S" & lrfordatara2c) 

            .AutoFilter Field:=1, Criteria1:="<>0" 

            .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete 

        End With 

        .AutoFilterMode = False 

    End With 

    Application.CutCopyMode = False 

     

    'number of workflows 

    nsheeta2c.Activate 

    Dim lrfornfra2c As Long 

    lrfornfra2c = Cells(nsheeta2c.Rows.Count, "C").End(xlUp).Row 

     

    With nsheeta2c 

        .AutoFilterMode = False 

        With .Range("T1:T" & lrfornfra2c) 

            .AutoFilter Field:=1, Criteria1:="<>0" 

            .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete 

        End With 

        .AutoFilterMode = False 

    End With 

    Application.CutCopyMode = False 

     

     

    ' Default group remove dba,labeling,PCS,regulatory 

     

    'dba 

    nsheeta2c.Activate 

    Dim lrfordbaa2c As Long 

    lrfordbaa2c = Cells(nsheeta2c.Rows.Count, "C").End(xlUp).Row 

     

    With nsheeta2c 

        .AutoFilterMode = False 

        With .Range("E1:E" & lrfordbaa2c) 

            .AutoFilter Field:=1, Criteria1:="dba" 

            .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete 

        End With 

        .AutoFilterMode = False 

    End With 

    Application.CutCopyMode = False 

     

    'Labeling 

    nsheeta2c.Activate 

    Dim lrforlba2c As Long 

    lrforlba2c = Cells(nsheeta2c.Rows.Count, "C").End(xlUp).Row 

     

    With nsheeta2c 

        .AutoFilterMode = False 

        With .Range("E1:E" & lrforlba2c) 

            .AutoFilter Field:=1, Criteria1:="Labeling" 

            .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete 

        End With 

        .AutoFilterMode = False 

    End With 

    Application.CutCopyMode = False 

     

    'PCS 

    nsheeta2c.Activate 

    Dim lrforpcsa2c As Long 

    lrforpcsa2c = Cells(nsheeta2c.Rows.Count, "C").End(xlUp).Row 

     

    With nsheeta2c 

        .AutoFilterMode = False 

        With .Range("E1:E" & lrforpcsa2c) 

            .AutoFilter Field:=1, Criteria1:="PCS" 

            .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete 

        End With 

        .AutoFilterMode = False 

    End With 

    Application.CutCopyMode = False 

     

    'Regulatory 

    nsheeta2c.Activate 

    Dim lrforrega2c As Long 

    lrforrega2c = Cells(nsheeta2c.Rows.Count, "C").End(xlUp).Row 

     

    With nsheeta2c 

        .AutoFilterMode = False 

        With .Range("E1:E" & lrforrega2c) 

            .AutoFilter Field:=1, Criteria1:="Regulatory" 

            .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete 

        End With 

        .AutoFilterMode = False 

    End With 

    Application.CutCopyMode = False 

     

     

    'x(*2) 

    Dim tbla2c As ListObject 

    Dim rnga2c As Range 

    Dim lrfortablea2c As Long 

    lrfortablea2c = Cells(nsheeta2c.Rows.Count, "C").End(xlUp).Row 

    Set rnga2c = nsheeta2c.Range("A1:T" & lrfortablea2c) 

 

    Set tbla2c = nsheeta2c.ListObjects.Add(xlSrcRange, rnga2c, , xlYes) 

    tbla2c.TableStyle = "TableStyleMedium2" 

     

    Application.CutCopyMode = False 

     

    MsgBox "Reports have been generated" 

     

End Sub 

 

 

 

 