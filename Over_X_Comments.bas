Attribute VB_Name = "Over_X_Comments"
Sub Over_X_Comments()

    Dim overfile As String, Over_90, Over_90_MN As Variant, i As Integer
    
    overfile = "Over 90 Days Comment_" & Format(Date, "yyyy-mm-dd") & ".xlsx"
    
    
    Over_90 = Workbooks(overfile).Worksheets("Over 90 Comments").Range("K3:L" & Workbooks(overfile).Worksheets("Over 90 Comments").Range("K" & rows.Count).End(xlUp).row)
    Over_90_MN = Workbooks(overfile).Worksheets("Minnesota").Range("K3:L" & Workbooks(overfile).Worksheets("Minnesota").Range("K" & rows.Count).End(xlUp).row)
    
    i = 1
    While i < (UBound(Over_90) - LBound(Over_90) + 2)
        Over_90(i, 2) = CheckComment(Over_90(i, 2), "\d{1,2}\s\w{3,9}\s\d{4}")
        Workbooks(overfile).Worksheets("Over 90 Comments").Range("M" & i + 2).Value = Over_90(i, 2)
        i = i + 1
    Wend
    
    i = 1
    While i < (UBound(Over_90_MN) - LBound(Over_90_MN) + 2)
        Over_90_MN(i, 2) = CheckComment(Over_90_MN(i, 2), "\d{1,2}\s\w{3,9}\s\d{4}")
        Workbooks(overfile).Worksheets("Minnesota").Range("M" & i + 2).Value = Over_90_MN(i, 2)
        i = i + 1
    Wend

End Sub
Function CheckComment(Statement, datepattern As String)
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    With regex
        .Pattern = datepattern
        .IgnoreCase = True
        .Global = True
        .MultiLine = True
    End With
    
    Set matches = regex.Execute(Statement)
    For Each match In matches
        If IsDate(match) Then
            If CDate(match) > DateAdd("d", -90, Date) And InStr(Statement, "greater than") > 0 Then
                CheckComment = "Complete"
                Exit Function
            End If
        End If
        Next match
    CheckComment = "Need Comment"
End Function

