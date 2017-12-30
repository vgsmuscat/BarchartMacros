Attribute VB_Name = "MiniProg2"

'Created 26 Dec 17
'Last Modified 28 Dec 17 12:14 PM

'Balance Work
'Work on Barchart is totally balance, it is a variant of Barchats in Miniporg1
'Challages are as follow:
   'Immediate Challenges
    '1. The code has been developed taking below 6 points limitation.
    '2. However the Code does not give proper output if the duration of activity is too long due to round of problem in interpolation in curved activity
    '3. Revise the
    '3. Remove the comment colon for printing debugging
    
   'Major Challenges
    '1. How to take care of moving CD
    '2 How to take care of project Start Date
    '3 How to take care of array output as per moving CD
    '4 Array output to be in line with the First Current date
    '5 While individual calculation to be as per project Start Date & to take care of Jump
    '6 All calculation in array of activity(), hard coded size for 3 years.
    


'This macro will should work for days, week, month & year in columns


'All the function and sub are related and the variable are interexhanged so take care
'Barchart function may work independed of other subs.
Dim TW As Double, fcd As Range, sd As Range, fd As Range, LOH As Range, jp As Integer
Dim CurveRange As Variant 'Curve Range in Active Workbook based on curve Sheet format
Dim dwCurves() As Variant ' Day Wise Curves based on Curve Sheet format
Dim pCurves() As Variant 'Prorata Curves based on Curve Sheet format
Function Barchart(StartDate As Range, FinishDate As Range, currentdate As Range, Optional curveNo, Optional ListofHolidays As Range)
    'Lenghth of total span is hard coded
    Dim activity() As Variant
    Dim cn As Integer
    Dim curve() As Variant
    Dim curvedActivity() As Variant
    On Error GoTo debg
    
    Set sd = StartDate
    Set fd = FinishDate
    Set fcd = currentdate 'First Current Date
    Set LOH = ListofHolidays
    If IsEmpty(cn) Then
        cn = 1
    Else
        cn = curveNo
    End If
    aSize = 900 'Activity Size
    
    ReDim activity(1 To aSize)
    firstPeriodSet = False ' This is used for interpolation in curves. To Start the interpolation to curve
    For i = 1 To UBound(activity)
        Set cd = fcd.Offset(0, i - 1)
        jp = cd.Offset(0, 1) - cd - 1 'Jump Period Days
        CDEoP = cd + jp 'Current Date End of Jump Period
        twd = Application.WorksheetFunction.NetworkDays_Intl(sd, fd, 11, LOH) 'Total Working Day
        'sdSOP = sd + jp  'Start Period date of SD
        'fdSOP = fd + jp 'Finish Period Date period end
        
        'Activity within the period
        If (sd >= cd And fd <= CDEoP) Then
            
            'WDFP = WORKING DAY FOR PERIOD
             WDFP = Application.WorksheetFunction.NetworkDays_Intl(sd, fd, 11, LOH)
    
        ' Activity Starting before the CD but ending within the Jump Period
        ElseIf (cd >= sd And CDEoP <= fd) Then
            
            'TWD = TOTAL WORKING DAY
            WDFP = Application.WorksheetFunction.NetworkDays_Intl(cd, CDEoP, 11, LOH)
    
        ' For SD between the period & FD greater then SDEoP
        ElseIf (sd >= cd And sd <= CDEoP) Then
        
            WDFP = Application.WorksheetFunction.NetworkDays_Intl(sd, CDEoP, 11, LOH)
        
        ' for FD between the period and SD less than the FD
        ElseIf (fd >= cd And fd <= CDEoP) Then
        
            WDFP = Application.WorksheetFunction.NetworkDays_Intl(cd, fd, 11, LOH)
        ' Month not bewteen SD & FD
        Else
            WDFP = 0
        End If
       
        If WDFP > 0 Then
            If firstPeriodSet = False Then
                firstperiod = i
                firstPeriodSet = True
            End If
            activity(i) = WDFP / twd * 100
            lastPeriod = i ' This is to capture the last column where the working days are entered, this is to help_
                           ' in rounding of to 100 the curved figures below
            'Debug.Print "activity(" & i & ") = " & activity(i)
        Else
            activity(i) = 0
            'Debug.Print "activity(" & i & ") = " & activity(i)
            
        End If
    Next
    
    
    ReDim curvedActivity(1 To aSize)
    curve = GetCurve(cn, twd)
    aSum = 0 'Activity Cumulative Percentage sum
    ''Debug.Print "First Period = " & firstperiod & ", Last Period = " & lastPeriod
    For i = firstperiod To lastPeriod
        
        cSum = 0 'Curve Cumulative Percentage sum
        If i <> lastPeriod Then
            If i = firstperiod Then
                aSumLast = 0
            Else
                aSumLast = aSum + 1
            End If
            aSum = Round(aSum + activity(i), 0) ' This round off create problem for long duration activity as it exhaust the range before_
                                                ' _completion of activity
            If aSum > 100 Then aSum = 100 'Adjustment made for rounding becoming more than 100
            'Interpolation from Curve
            For j = aSumLast To aSum
                cSum = cSum + curve(j)
            Next
            
            'Assign to curvedActivity
            curvedActivity(i) = cSum
            'Debug.Print "for " & aSumLast & " to " & aSum; " - curvedActivity(" & i & ") = " & cSum
        Else    'For Last Period in activity for rounding off
            If aSumLast < 100 Then aSumLast = aSum + 1
            aSum = 100
            
            'Interpolation from Curve
            For j = aSumLast To aSum
                cSum = cSum + curve(j)
            Next
            
            'Assign to curvedActivity
            curvedActivity(i) = cSum
            'Debug.Print "for " & aSumLast & " to " & aSum; " - curvedActivity(" & i & ") = " & cSum
        End If
    Next
        
    
    'CheckInExcel activity, Application.Worksheets("Sheet2") 'Will not work as you cannot modify excel in edit mode
    Barchart = curvedActivity()
debg:
    If Err.Number <> 0 Then
        'Debug.Print "i = " & i & " -- " & Err.Description
        activity(i) = CVErr(xlErrValue) 'Check can be 0
        

        
    End If
    
End Function


'For this Curve sheet should be part of the current file
'Name the Range in the Curve sheet including the Curve No and its name as "Curves"
'Note size here starts as 0 as in p6 curves 0 day can have some value.
Function GetCurve(curveNo As Integer, totalWorkingDay As intger)
    Dim curve() As Variant
    On Error GoTo debg
    csize = 100 'Curve Size (Default)
    If twd > 100 Then csaize = totalWorkingDay
    ReDim curve(0 To csize)
    cn = curveNo
    arrayInitialized = LBound(dwCurves)
    If arrayInitialized = "Not Initialized" Then
        MakeDayWiseCurves   'Note Day Wise cure is of size 0 to 100 - think and start from here critical
    End If
    For i = 0 To csize
        curve(i) = dwCurves(cn + 1, i + 3)
        'Debug.Print "curve" & cn & "(" & i & ") = " & curve(i); ""
    Next
    GetCurve = curve()
debg:
    If Err.Number <> 0 Then
        'Debug.Print Err.Description
        arrayInitialized = "Not Initialized"
        Resume Next
    End If
End Function
Sub LoadCurves()
    CurveRange = ThisWorkbook.Names("Curves").RefersToRange
End Sub


Sub CheckInExcel(arr() As Variant, Optional Sh As Worksheet)
    On Error GoTo debg
    Dim dimension As Variant
    If Sh Is Nothing Then
        MsgBox "Array will print in activesheet"
        Set cel = ActiveCell
        
    Else
        Set cel = Sh.Range("A1")
    End If
    dimension = UBound(arr, 2)
    If dimension <> "Single Dimension" Then
        For i = LBound(arr) To UBound(arr)
            For j = LBound(arr, 1) To UBound(arr, 2)
                cel.Offset(i, j) = arr(i, j)
            Next
        Next
    ElseIf dimension = "Single Dimension" Then
        For i = LBound(arr) To UBound(arr)
            cel.Select
            cel.Offset(1, i) = arr(i)
            
        Next
    End If
debg:
    If Err.Number <> 0 Then
        'Debug.Print Err.Description & " at i = " & i & ", j = " & j
        dimension = "Single Dimension"
        Resume Next
    End If
End Sub
'This is based on the Curve sheet
'First row consist of header
'First two column and last two column data is not touched
'Data for each column is repeated 5 times to make it 0 - 100 days system
Sub MakeDayWiseCurves()
    
    Application.SendKeys "^g ^a {DEL}" ' Clear immediate window
    
    Dim dwcX As Integer 'Day Wise Curve Array Size X
    Dim dwcY As Integer 'Day Wise Curve Array Size Y
    dwcX = 30
    dwcY = 104
    ReDim dwCurves(1 To dwcX, 1 To dwcY)
    If IsEmpty(CurveRange) Then LoadCurves
    For curveRow = LBound(CurveRange) To UBound(CurveRange)
        dayCol = 1
        Day1 = 1
        
        For curveCol = LBound(CurveRange, 2) To UBound(CurveRange, 2)
            If Not IsError(Application.Match(curveCol, Array(1, 2, 3, 24), 0)) Then
                dwCurves(curveRow, dayCol) = CurveRange(curveRow, curveCol)
                'Debug.Print "Item ("; curveRow; ","; dayCol; ")", dwCurves(curveRow, dayCol)
                dayCol = dayCol + 1
            Else
                If curveRow = 1 Then
                For i = 1 To 5
                    dwCurves(curveRow, dayCol) = Day1
                    'Debug.Print "Item ("; curveRow; ","; dayCol; ")", dwCurves(curveRow, dayCol)
                    Day1 = Day1 + 1
                    dayCol = dayCol + 1
                Next
                Else
                    For i = 1 To 5
                        dwCurves(curveRow, dayCol) = CurveRange(curveRow, curveCol)
                        'Debug.Print "Item ("; curveRow; ","; dayCol; ")", dwCurves(curveRow, dayCol)
                        dayCol = dayCol + 1
                    Next
                End If
            End If
        Next
    Next
    dwCurves = MakeCurves_prorata(dwCurves)
    
    'CheckInExcel dwCurves
End Sub
'use curve data based on the curve data sheet
'first two column consist of Curve No. and Description
'First row consist of header
'Last Column (No. 104) consist of total of values of Column 3 to Column 103
Function MakeCurves_prorata(arr() As Variant)
    pCurves = arr ' Prorata Curves based on curve Sheet format
    For curveRow = 2 To UBound(pCurves)
        sumRow = 0
        For curveCol = 3 To 103
            sumRow = sumRow + pCurves(curveRow, curveCol)
        Next
        If sumRow > 0 Then
            For curveCol = 3 To 103
                pCurves(curveRow, curveCol) = pCurves(curveRow, curveCol) / sumRow
                'Debug.Print "Item ("; curveRow; ","; curveCol; ")", pCurves(curveRow, curveCol)
                
            Next
        End If
    Next
    MakeCurves_prorata = pCurves
End Function


