Sub generate() 

    Dim wsmaster As Worksheet 

    Set wsmaster = Workbooks("temp.xlsx").Worksheets("temp") 

     

    Dim i As Long 

 

    wsmaster.Activate 

 

    Dim cnt2 As Long 

 

    cnt2 = Cells(wsmaster.Rows.Count, "Q").End(xlUp).Row 

    MsgBox (cnt2) 

      

    'Base logic 

    'if value in column s = copy that to column t 

    'if no value in column s = copy previous cell of column t 

    ' try till 256 

     

    For i = 5 To cnt2 

        Dim val As Integer 

        val = Len(Range("S" & i).Value) 

        If val > 0 Then 

            wsmaster.Range("S" & i).Copy 

            wsmaster.Range("T" & i).PasteSpecial Paste:=xlPasteValuesAndNumberFormats 

            'MsgBox (Len(Range("S" & i).Value)) 

        Else 

            'MsgBox (i - 1) 

            wsmaster.Range("T" & i).Formula = "=T" & i - 1 

        End If 

    Next i 

    MsgBox ("Done") 

End Sub 