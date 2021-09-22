Sub extractdataRCC() 

 

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

    Dim ln As Long 

    i = 2 

    ss = "Root Cause Category" 

    'ss1 = vbLf 

     

    'declare a variant array 

    Dim rcc(1 To 13) As String 

  

    'populate the array 

    rcc(1) = "Training" 

    rcc(2) = "Integrity" 

    rcc(3) = "Request" 

    rcc(4) = "Integration" 

    rcc(5) = "Change" 

    rcc(6) = "Customization" 

    rcc(7) = "Localization" 

    rcc(8) = "Std SW Bug" 

    rcc(9) = "Security" 

    rcc(10) = "Master Data" 

    rcc(11) = "Infrastructure" 

    rcc(12) = "Process" 

    rcc(13) = "Configuration" 

  

    'declare a variant to hold the array element 

    Dim item As Variant 

     

    For Each cell In rng 

        val = cell.Value 

        ln = Len(val) 

        If InStr(val, ss) <> 0 Then 

            ind = InStr(val, ss) 

            valforU = Mid(val, ind, ln - ind + 1) 

            'loop through the entire array 

            For Each item In rcc 

                If InStr(valforU, item) <> 0 Then 

                    Range("U" & i).Value = item 

                    Exit For 

                End If 

            Next item 

            'Range("U" & i).Value = valforU 

        End If 

        i = i + 1 

    Next cell 

     

    MsgBox "RCC-Final Column has been populated" 

     

     

End Sub 

 

 

 