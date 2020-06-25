Attribute VB_Name = "Timer_General"
Sub Timer()
    Dim StartTime As Double
    Dim SecondsElapsed As Double
    StartTime = Timer
    
    SecondsElapsed = Timer - StartTime
    MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation
    
End Sub
