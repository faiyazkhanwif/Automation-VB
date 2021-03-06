Sub fullautotest() 

 

    Dim wsCopy As Worksheet, wsDest As Worksheet, wsfin As Worksheet 

    Dim lCopyLastRow As Long, lDestLastRow As Long 

     

    'Set variables for copy and destination sheets 

     

    Dim WS As Worksheet 

    For Each WS In Workbooks("Report for Jim Team.xlsx").Worksheets 

        WS.Name = "Sheet1" 

    Next WS 

     

    Set wsCopy = Workbooks("Report for Jim Team.xlsx").Worksheets("Sheet1") 

    Set wsDest = Workbooks("CUMULATIVE DAILY FILE.xlsx").Worksheets("Sheet1") 

    Set wsfin = Workbooks("How to load HCP contracts to JDE F4801.xlsm").Worksheets("Load Template") 

     

    '1. Find last used row in the copy range based on data in column A 

    lCopyLastRow = wsCopy.Cells(wsCopy.Rows.Count, 1).End(xlUp).Row 

     

    '2. Find first blank row in the destination range based on data in column A 

    'Offset property moves down 1 row 

    lDestLastRow = wsDest.Cells(wsDest.Rows.Count, 4).End(xlUp).Offset(1, 0).Row 

     

    '3. Copy & Paste Data 

    wsCopy.Range("A2:C" & lCopyLastRow).Copy _ 

    wsDest.Range("A" & lDestLastRow) 

     

    'weekly or daily check 

    Dim dt As String 

    sDayName = Format(Format(Date, "mm/dd/yy"), "dddd") 

 

    Dim word As String 

     

    If sDayName = "Monday" Then 

        dt = Format(Date - 3, "mm/dd/yy") 

        word = "weekly" 

    Else 

        dt = Format(Date - 1, "mm/dd/yy") 

        word = "daily" 

    End If 

 

    Dim fcfordt As Long, lcfordt As Long 

    fcfordt = lDestLastRow 

    lcfordt = wsDest.Cells(wsDest.Rows.Count, 1).End(xlUp).Row 

    wsDest.Range(wsDest.Cells(fcfordt, 5), wsDest.Cells(lcfordt, 5)).Value = dt 

    wsDest.Range(wsDest.Cells(fcfordt, 4), wsDest.Cells(lcfordt, 4)).Value = word 

     

    For i = fcfordt To lcfordt 

        'wsDest.Cells(i, 6).Value = "N/A" 

        Dim MyVal As Variant 

        Dim res As Variant 

        'On Error Resume Next 

        MyVal = Application.VLookup(wsDest.Cells(i, 2), wsDest.Range("B5:B" & i - 1), 1, False) 

        If IsError(MyVal) Then 

            res = "#N/A" 

        Else 

            res = MyVal 

        End If 

        wsDest.Cells(i, 6).Value = Application.WorksheetFunction.IfError(MyVal, "#N/A") 

        'On Error GoTo 0 

         

    Next i 

     

    Dim r As Long 

    r = 2 

    wsfin.Rows("2:" & wsfin.Rows.Count).ClearContents 

    For j = fcfordt To lcfordt 

        Dim val As Variant 

        val = wsDest.Range("F" & j).Value 

        'MsgBox (val) 

        If IsNumeric(val) Then 

             

            'wsfin.Range("A" & r).Value = wsDest.Range("B" & j).Value 

            'wsfin.Cells(r, 2).Value = wsDest.Cells(j, 3).Value 

            'wsfin.Cells(r, 6).Value = wsDest.Cells(j, 1).Value 

            'r = r + 1 

        Else 

            wsfin.Range("F" & r) = wsDest.Cells(j, 1) 

            wsfin.Range("A" & r) = wsDest.Cells(j, 2) 

            wsfin.Range("B" & r) = wsDest.Cells(j, 3) 

            r = r + 1 

            'dt = Format(Date - 1, "mm/dd/yy") 

            'word = "daily" 

        End If 

    Next j 

     

    'MsgBox ("Done!") 

    Dim lr As Long 

    lr = wsfin.Cells(Rows.Count, 1).End(xlUp).Row 

     

    If lr = 1 Then 

        MsgBox ("All duplicates, no data to be loaded in JDE.") 

    Else 

        If sDayName = "Monday" Then 

            Application.Run "Macro2" 

        Else 

            Application.Run "Macro1" 

        End If 

 

        Application.Run "deleteredundant" 

     

        Application.Run "generateCSV" 

    End If 

 

    'MsgBox ("Done!") 

End Sub 

 

 

 

 

 

 