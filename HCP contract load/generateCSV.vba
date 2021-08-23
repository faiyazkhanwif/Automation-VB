Sub generateCSV() 

    Dim wsfin As Worksheet 

    Set wsfin = Workbooks("How to load HCP contracts to JDE F4801.xlsm").Worksheets("Load Template") 

     

    'weekly or daily check 

    Dim dt As String 

    sDayName = Format(Format(Date, "mm/dd/yy"), "dddd") 

     

    If sDayName = "Monday" Then 

        dt = Format(Date - 3, "yyyy/mm/dd") 

    Else 

        dt = Format(Date - 1, "yyyy/mm/dd") 

    End If 

     

    Dim s1 As String 

    s1 = Replace(dt, "/", "") 

     

    Dim filename As String 

    filename = "RFA_" & s1 

     

    wsfin.Copy 

    ActiveWorkbook.SaveAs filename:=ThisWorkbook.Path & "\" & filename, FileFormat:=xlCSV, CreateBackup:=False 

    ActiveWorkbook.Close 

     

    Application.DisplayAlerts = True 

     

    MsgBox ("CSV File has been generated!") 

     

End Sub 

 