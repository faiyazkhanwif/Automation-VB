Sub extractdataPER() 

 

    Dim ws As Worksheet 

 

    Set ws = Workbooks("2021 JDE Incident Trending_JAPAC (Jan-Aug2021).xlsm").Worksheets("Incidents") 

 

    Dim rng As Range, cell As Range 

     

    lRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row 

    'lRow = 900 

     

    Set rng = Range("P2:P" & lRow) 

     

    Dim val As String 

    Dim ss As String 

    Dim ind As Integer 

    Dim ind1 As Integer 

    Dim i As Integer 

     

    i = 2 

     

    ss = "Program Error Reported" 

    ss1 = vbLf 

     

    For Each cell In rng 

        val = cell.Value 

        If InStr(val, ss) <> 0 Then 

            ind = InStr(val, ss) + 23 

            'ind1 = InStr(val, ss1) - 24 

            ind1 = InStr(val, ss1) 

            If (ind1 <= ind) Then 

                Dim x As Long 

                x = 1 

                Do While ind1 <= ind 

                    ind1 = InStr(x + 1, val, ss1, vbTextCompare) 

                    'MsgBox "The value of i is : " & i 

                    x = x + 1 

                Loop 

            End If 

            'MsgBox ind1 

            ind1 = ind1 - ind - 1 + 1 

            valforT = Mid(val, ind, ind1) 

            valforT = Replace(valforT, vbLf, "") 

            'valforT = Replace(valforT, " ", "") 

            valforT = Trim(valforT) 

            'MsgBox valforT 

            Range("T" & i).Value = valforT 

        End If 

        i = i + 1 

    Next cell 

     

    MsgBox "Program Error Reported Column has been populated" 

     

     

End Sub 

 