Sub Macro1()

    Dim wsfin As Worksheet
    Set wsfin = Workbooks("Training_Progress_Summary.xlsx").Worksheets("Training Progress Summary")
    
    wsfin.Activate
    Dim last_row As Long
    last_row = wsfin.Cells(Rows.Count, 1).End(xlUp).Row
    'MsgBox (last_row)
    
    'Current date
    Dim dtToday As Date
    dtToday = Now()

    wsfin.Activate
    Range("P8").EntireColumn.Insert
    Range("P8").Value = "Days Due"
    
    Dim i As Integer
    For i = 9 To last_row 'Check starts from row 9
        If Range("H" & i).Value = "Completed (Equivalent)" Or Range("H" & i).Value = "Completed" Then
            Range("P" & i).Value = "Completed"
            Range("P" & i).Interior.Color = RGB(159, 255, 128)
        Else
            'MsgBox "Hi"
            If Len(Range("O" & i).Value) = 0 Then

            Else
                Dim dtDue As Date
                dtDue = Range("O" & i).Value
                Dim days As Long
                days = DateDiff("D", dtToday, dtDue)
                hours = DateDiff("H", dtToday, dtDue)
                If days <= 0 Then
                    If hours <= 0 Then
                        Range("P" & i).Value = "Overdue"
                        Range("P" & i).Interior.Color = RGB(255, 102, 102)
                    Else
                        Range("P" & i).NumberFormat = "0"
                        Range("P" & i).Value = "1"
                        Range("P" & i).Interior.Color = RGB(255, 187, 153)
                    End If
                Else
                    Range("P" & i).NumberFormat = "0"
                    Range("P" & i).Value = days
                    If (Range("P" & i).Value) <= 7 Then
                        Range("P" & i).Interior.Color = RGB(255, 160, 122)
                    ElseIf (Range("P" & i).Value) <= 14 Then
                        Range("P" & i).Interior.Color = RGB(255, 255, 102)
                    Else
                        Range("P" & i).Interior.Color = RGB(159, 255, 128)
                    End If
                End If
            End If
        End If
    Next i
    
    For i = 9 To last_row
        If Range("P" & i).Value <> "Overdue" Then
            If Range("P" & i).Value = "Completed" Then
                Range("P" & i).EntireRow.Delete
                i = i - 1
            ElseIf Len(Range("P" & i).Value) = 0 And Len(Range("B" & i).Value) <> 0 Then
                Range("P" & i).EntireRow.Delete
                i = i - 1
            Else
                If Range("P" & i).Value > 14 Then
                    Range("P" & i).EntireRow.Delete
                    i = i - 1
                End If
            End If
        End If
    Next i
    
    Range("A:A,B:B,C:C,E:E,F:F,I:I,J:J,K:K,L:L,M:M,N:N,Q:Q,R:R").Delete
    MsgBox ("Done")
    
End Sub

