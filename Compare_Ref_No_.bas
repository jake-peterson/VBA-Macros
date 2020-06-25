Attribute VB_Name = "Compare_Ref_No_Healy"
Sub CompareColumns()
Application.ScreenUpdating = False

    Dim RCA_ref_no As Variant
    RCA_ref_no = Workbooks("Remote Care Assignments 07 FEB 2019.xlsm").Worksheets(1).Range("B2:B" & Workbooks("Remote Care Assignments 07 FEB 2019.xlsm").Worksheets(1).Range("B" & rows.Count).End(xlUp).row)
    
    Dim CallLogs As Variant, logs As Integer, Calls As New Collection
    CallLogs = Workbooks("Call log.xlsm").Worksheets(1).Range("A2:E" & Workbooks("Call log.xlsm").Worksheets(1).Range("E" & rows.Count).End(xlUp).row)
    
    logs = 0
    While logs < (UBound(CallLogs) - LBound(CallLogs) + 1)
        If CallLogs(logs + 1, 5) = "Closed - FCR" Or CallLogs(logs + 1, 5) = "Open - FCR" Then
            Calls.Add (CallLogs(logs + 1, 1))
        End If
        logs = logs + 1
    Wend
    Debug.Print Calls.Count
    
    
    
    Dim x, y, match As Boolean, Todays_Date As Date, Count As Integer, Missing_Logs As New Collection
    Count = 0
    Todays_Date = Date
    For Each x In Calls
        If IsDate(Left(x, 10)) Then
            If CDate(Left(x, 10)) = (Todays_Date - 3) Then
                match = False
                For Each y In RCA_ref_no
                    If CStr(x) = CStr(y) Then
                        Count = Count + 1
                        match = True
                        Missing_Logs.Add (CStr(x))
                    End If
                Next y
            End If
        End If
    Next
    Debug.Print Missing_Logs.Count
    
    Dim NewSheet As Worksheet, i As Integer
    Set NewSheet = Sheets.Add(After:=Sheets(Worksheets.Count))
    NewSheet.Name = "Duplicate Reference Numbers"
    Worksheets(2).Cells(1, 1).Value = "Secondary Ref No"
    Dim k As Integer
    For i = 0 To Missing_Logs.Count - 1
        Worksheets(2).Cells((i + 2), 1).Value = Missing_Logs(i + 1)
    Next i
    
Application.ScreenUpdating = True
End Sub


