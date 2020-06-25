Attribute VB_Name = "AddNextWeek_WFH"
Sub AddNextWeekMN()
    
    'Speeds up code
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim oldformuladate, newformuladate, newstartdate, newenddate As String, rng As Variant
    
    'Sets rng equal to last weeks formulas for C - N
    rng = Workbooks("WFH Metrics Formulas (MN).xlsm").Sheets(2).Range("C2:N18").Formula
    
    'Calculates the new start/end dates
    oldformuladate = Format(Workbooks("WFH Metrics Formulas (MN).xlsm").Sheets(2).Cells(2, 2).Value, "MMM-dd-yyyy")
    newformuladate = Format(DateAdd("d", 7, oldformuladate), "MMM-dd-yyyy")
    newstartdate = Format(DateAdd("d", 1, oldformuladate), "mm/dd/yy")
    newenddate = Format(DateAdd("d", 6, newstartdate), "mm/dd/yy")
    
    'Adds rows to sheet
    Workbooks("WFH Metrics Formulas (MN).xlsm").Sheets(2).rows("2:18").Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromRightOrBelow
    
    'Changes dates of A & B
    Workbooks("WFH Metrics Formulas (MN).xlsm").Sheets(2).Range("A2:A18").Value = newstartdate
    Workbooks("WFH Metrics Formulas (MN).xlsm").Sheets(2).Range("B2:B18").Value = newenddate
    'Sets C-N equal to last weeks formulas
    Workbooks("WFH Metrics Formulas (MN).xlsm").Sheets(2).Range("C2:N18") = rng
    'Changes old date to new date inside formula
    Workbooks("WFH Metrics Formulas (MN).xlsm").Sheets(2).rows("2:18").Replace what:=oldformuladate, replacement:=newformuladate
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub

Sub AddNextWeekCA()
    
    'Speeds up code
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim oldformuladate, newformuladate, newstartdate, newenddate As String, rng As Variant
    
    'Sets rng equal to last weeks formulas for C - N
    rng = Workbooks("WFH Metrics Formulas (CA).xlsm").Sheets(2).Range("C2:N37").Formula
    
    'Calculates the new start/end dates
    oldformuladate = Format(Workbooks("WFH Metrics Formulas (CA).xlsm").Sheets(2).Cells(2, 2).Value, "MMM-dd-yyyy")
    newformuladate = Format(DateAdd("d", 7, oldformuladate), "MMM-dd-yyyy")
    newstartdate = Format(DateAdd("d", 1, oldformuladate), "mm/dd/yy")
    newenddate = Format(DateAdd("d", 6, newstartdate), "mm/dd/yy")
    
    'Adds rows to sheet
    Workbooks("WFH Metrics Formulas (CA).xlsm").Sheets(2).rows("2:37").Insert Shift:=xlShiftDown, CopyOrigin:=xlFormatFromRightOrBelow
    
    'Changes dates of A & B
    Workbooks("WFH Metrics Formulas (CA).xlsm").Sheets(2).Range("A2:A37").Value = newstartdate
    Workbooks("WFH Metrics Formulas (CA).xlsm").Sheets(2).Range("B2:B37").Value = newenddate
    'Sets C-N equal to last weeks formulas
    Workbooks("WFH Metrics Formulas (CA).xlsm").Sheets(2).Range("C2:N37") = rng
    'Changes old date to new date inside formula
    Workbooks("WFH Metrics Formulas (CA).xlsm").Sheets(2).rows("2:37").Replace what:=oldformuladate, replacement:=newformuladate
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub













