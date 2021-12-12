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

     

    MsgBox ("Done") 

End Sub 