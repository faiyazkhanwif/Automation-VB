Sub findsecondobjectdesc() 

 

    Dim wsmaster As Worksheet 

    Set wsmaster = Workbooks("temp.xlsx").Worksheets("temp") 

     

     

    Dim wssec As Worksheet 

    Set wssec = Workbooks("PROD_SOD Rules Matrix as of Dec 17 2021.xlsx").Worksheets("Sheet1") 

 

    Dim range1 As Range, range2 As Range 

     

     

    wsmaster.Activate 

    cnt1 = Cells(wsmaster.Rows.Count, "A").End(xlUp).Row 

    MsgBox (cnt1) 

     

    wssec.Activate 

    cnt2 = Cells(wssec.Rows.Count, "A").End(xlUp).Row 

    MsgBox (cnt2) 

     

     

    wsmaster.Activate 

    Set range1 = wsmaster.Range("H2:H" & cnt1) 

     

    wssec.Activate 

    Set range2 = wssec.Range("H2:H" & cnt2) 

     

    For Each c1 In range1 

        For Each c2 In range2 

            If c1.Value = c2.Value Then 

                c1.Offset(0, 1).Value = c2.Offset(0, 1).Value 

            End If 

        Next c2 

    Next c1 

     

    MsgBox ("Done") 

End Sub 

 

 