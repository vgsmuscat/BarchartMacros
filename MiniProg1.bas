Attribute VB_Name = "MiniProg1"
'This macro will should work for days, week, month & year in columns
'Created 26 Dec 17
'put returnText as "tb" (Time Based for returning networking days instead of other text
Dim TW As Double, cd As Range, sd As Range, fd As Range, LOH As Range, jp As Integer
Function Barcharts(StartDate As Range, FinishDate As Range, currentdate As Range, returnText As Variant, Optional ListofHolidays As Range)
    On Error GoTo debg
    Set sd = StartDate
    Set fd = FinishDate
    Set cd = currentdate
    Set LOH = ListofHolidays
    rt = returnText
    jp = cd.Offset(0, 1) - cd - 1 'Jump Period Days
    
    CDEoP = cd + jp 'Current Date End of Jump Period
    TWD = Application.WorksheetFunction.NetworkDays_Intl(sd, fd, 11, LOH) 'Total Working Day
    'sdSOP = sd + jp  'Start Period date of SD
    'fdSOP = fd + jp 'Finish Period Date period end
    
    'Activity within the period
    If (sd >= cd And fd <= CDEoP) Then
        
        'WDFJ = WORKING DAY FOR PERIOD
         WDFM = Application.WorksheetFunction.NetworkDays_Intl(sd, fd, 11, LOH)

    ' Activity Starting before the CD but ending within the Jump Period
    ElseIf (cd >= sd And CDEoP <= fd) Then
        
        'TWD = TOTAL WORKING DAY
        WDFM = Application.WorksheetFunction.NetworkDays_Intl(cd, CDEoP, 11, LOH)

    ' For SD between the period & FD greater then SDEoP
    ElseIf (sd >= cd And sd <= CDEoP) Then
    
        WDFM = Application.WorksheetFunction.NetworkDays_Intl(sd, CDEoP, 11, LOH)
    
    ' for FD between the period and SD less than the FD
    ElseIf (fd >= cd And fd <= CDEoP) Then
    
        WDFM = Application.WorksheetFunction.NetworkDays_Intl(cd, fd, 11, LOH)
    ' Month not bewteen SD & FD
    Else
        WDFM = 0
    End If
    
    If rt = "tb" And WDFM > 0 Then
        Barcharts = WDFM / TWD
    ElseIf WDFM > 0 Then
        Barcharts = rt
    Else
        Barcharts = 0
    End If


debg:
    If Err.Number <> 0 Then
        Barcharts = CVErr(xlErrValue)
    End If
    
End Function

