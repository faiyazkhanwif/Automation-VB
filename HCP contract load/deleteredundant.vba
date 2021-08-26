Sub deleteredundant() 

    Dim wsfin As Worksheet 

    Set wsfin = Workbooks("How to load HCP contracts to JDE F4801.xlsm").Worksheets("Load Template") 

    Dim last_row As Long 

    last_row = wsfin.Cells(Rows.Count, 1).End(xlUp).Row 

    For i = 2 To last_row 

        If wsfin.Cells(i, "F").Value = "JPM" Or wsfin.Cells(i, "F").Value = 1 Then 

            Rows(i).EntireRow.Delete 

            i = i - 1 

        End If 

        'If i < 2 Then 

        '    i = i + (2 - i) 

        'End If 

    Next i 

    Columns("A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete 

    'MsgBox ("Deleted redundant data!") 

End Sub 

 

 