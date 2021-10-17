Sub automate() 

    Dim ws As Worksheet 

    Set ws = Workbooks("Opcon Incidents Trending.xlsm").Worksheets("Incident") 

 

    Dim cell As Object 

    Dim count As Integer 

    Dim str As String 

    count = 0 

     

    For Each cell In Selection 

        Dim Result() As String 

        Dim program As String 

        Dim schedule As String 

        count = count + 1 

         

        str = cell 

         

        If InStr(str, ":") > 0 Then 

            Result() = Split(str, ":") 

             

            program = Trim(Result(1)) 

            schedule = Trim(Result(0)) 

             

            Dim rownum As Integer 

            rownum = cell.Row 

            'MsgBox cell.Row 

            'MsgBox program 

            'MsgBox schedule 

            ws.Cells(rownum, 14).Value = program 

            ws.Cells(rownum, 15).Value = schedule 

        End If 

    Next cell 

    Application.CutCopyMode = False 

    MsgBox count & " item(s) have been splitted" 

 

End Sub 