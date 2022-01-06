Sub generateruleID() 

    Dim wsmaster As Worksheet 

    Set wsmaster = Workbooks("temp.xlsx").Worksheets("temp") 

     

    Dim i As Long 

 

    wsmaster.Activate 

 

    Dim cnt2 As Long 

 

    cnt2 = 3071 

    MsgBox (cnt2) 

      

    'Base logic 

    'if value in column s = copy that to column t 

    'if no value in column s = copy previous cell of column t 

    ' try till 256 

     

    For i = 2 To cnt2 

        Dim val As Integer 

        val = Len(Range("A" & i).Value) 

        Dim data As String 

        If val > 0 Then 

            data = wsmaster.Range("A" & i).Value 

            'wsmaster.Range("S" & i).Copy 

            'wsmaster.Range("T" & i).PasteSpecial Paste:=xlPasteValuesAndNumberFormats 

            'MsgBox (Len(Range("S" & i).Value)) 

        Else 

            'MsgBox (i - 1) 

            'wsmaster.Range("T" & i).Formula = "=T" & i - 1 

            wsmaster.Range("A" & i).Value = data 

        End If 

    Next i 

    MsgBox ("Done") 

End Sub 

 