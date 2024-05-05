Attribute VB_Name = "main"
Sub Solve()
    'This sub executes the assignment operations to station task.
    'This sub is your Main function that you should edit.
    
    ''''''''''''''''''''''''''
    'clear output
    Call ClearOutput
    Range("B4").Interior.ColorIndex = 15
    'clear output
    
    ' now we define integers for our code
    Dim k           As Integer
    k = Worksheets("INPUT").Cells(2, "c")
    Dim i           As Integer
    i = 4        'rows for the first while
    Dim j           As Integer
    j = 8        'column for the second while
    Dim m           As Integer
    m = 4        'row for the second while
    Dim m1          As Integer
    m1 = 4        'row for the second while
    Dim n           As Integer
    n = 8        'row for the second while
    Dim l           As Integer
    l = 3        'row for the second while
    Dim algo4       As Boolean
    algo4 = True
    Dim algo5       As Boolean
    algo5 = True
    Dim algo6       As Boolean
    algo6 = False
    Dim c           As Integer
    c = 0
    
    'checks if its 1
    While Worksheets("INPUT").Cells(7, i) <> ""        ' runs on all the jobs
        If Worksheets("INPUT").Cells(7, i) <> Worksheets("INPUT").Cells(7, i - 1) Then        ' check if all the times are equal
        algo4 = False
    End If
    i = i + 1
Wend

''checks if its 2
Dim counterforeachmission As Integer
counterforeachmission = 0
While Worksheets("INPUT").Cells(7, m1 + 1) <> ""        ' runs on the slant
    If Worksheets("INPUT").Cells(j, m) <> 1 Then        ' check if we have a chain
    algo5 = False
End If
For for1 = 0 To counterforeachmission
    If Worksheets("INPUT").Cells(9 + for1, m1 + 2) = 1 Then
        algo5 = False
    End If
Next for1
m = m + 1
j = j + 1
m1 = m1 + 1
counterforeachmission = counterforeachmission + 1
Wend

'checks if its 3

Dim counter         As Integer
counter = 0
Dim counter2        As Integer
counter2 = 0

While Worksheets("INPUT").Cells(7, l) <> ""        ' this loop count the number of mission and precedence constraints
    counter = counter + 1
    For x = 1 To l - 3
        If Worksheets("INPUT").Cells(8, l) = 1 Then
            counter2 = counter2 + 1
        End If
    Next x
    l = l + 1
Wend

If counter < 11 And counter2 < 26 Then
    algo6 = True
End If

'mission 1

If algo4 = True Then
    Dim temp        As Integer
    temp = WorksheetFunction.Ceiling((counter / k), 1)        ' for the LB
    Dim LB          As Integer
    LB = temp * Worksheets("input").Cells(7, "c")
    c = LB
    Dim temp2       As Integer
    temp2 = 1
    Dim temptime    As Integer
    temptime = 0
    Dim counter3    As Integer
    counter3 = 0
    While counter3 < k
        temptime = 0
        For f = 1 To temp
            Worksheets("OUTPUT").Cells(f + 10, counter3 + 2) = f + (temp * counter3)        '  Places the tasks in the stations
            
            Dim temptime1 As Integer
            temptime1 = Worksheets("INPUT").Cells(7, 2 + f)
            temptime = temptime + temptime1
            Worksheets("OUTPUT").Cells(41, counter3 + 2) = temptime        ' count total time in each station
            If f + (temp * counter3) = counter Then
                Exit For
            End If
        Next f
        
        counter3 = counter3 + 1
    Wend
    'Next s
    
    Worksheets("OUTPUT").Cells(3, "b") = c
    Worksheets("OUTPUT").Cells(5, "b") = LB
    If c = LB Then
        Worksheets("OUTPUT").Cells(4, "b") = "yes"
         Range("B4").Interior.ColorIndex = 4
    Else
        Worksheets("OUTPUT").Cells(4, "b") = "Not necessarily "
         Range("B4").Interior.ColorIndex = 45
    End If
End If

'mission2

If algo4 = False And algo5 = True Then
    Call Submission2        ' calls the sub for the chain
End If

'mission 3
If algo6 = True And algo4 = False And algo5 = False Then
    Call ex3(counter)        ' calls the solver for the linear programing
End If

'mission 4

If algo6 = False And algo5 = False And algo6 = False Then        '

Dim weight()        As Integer        ' the array that will have all the wights
ReDim weight(1 To counter) As Integer
Dim LB4             As Integer
Dim y               As Integer
y = 3
Dim maxt            As Integer
Dim counter4        As Integer
counter4 = 0
maxt = Worksheets("INPUT").Cells(7, y)
While Worksheets("INPUT").Cells(7, y) <> ""
    weight(y - 2) = Worksheets("INPUT").Cells(7, y)        'creats the array for the wights
    counter4 = counter4 + Worksheets("INPUT").Cells(7, y)
    If Worksheets("INPUT").Cells(7, y + 1) > Worksheets("INPUT").Cells(7, y) Then
        maxt = Worksheets("INPUT").Cells(7, y + 1)
    End If
    y = y + 1
Wend
If maxt > WorksheetFunction.Ceiling((counter4 / k), 1) Then        ' chacks if the maxt is the LB
LB4 = maxt
Else
    LB4 = WorksheetFunction.Ceiling((counter4 / k), 1)        'end of check for LB
End If
Dim cformission4    As Integer
cformission4 = LB4
Dim rows            As Integer
Dim columns         As Integer
rows = 7
columns = counter + 2
Dim counter5        As Integer
counter5 = 0
While Worksheets("INPUT").Cells(rows, columns) <> "Duration"        'runs backward
    For q = 1 To counter - 1        'runs on all the kdimuyot
        If Worksheets("INPUT").Cells(rows + q, columns) = 1 Then
            weight(q) = weight(q) + weight(columns - 2)
        End If
    Next q
    columns = columns - 1
Wend
Dim sortedNumbers() As Integer
sortedNumbers() = ArrangeArrayPositionsDescending(weight)        ' arrange the array by its weights
Dim thereisarrange  As Boolean
thereisarrange = False

Dim beforeitsarrange() As Integer
ReDim beforeitsarrange(1 To counter) As Integer

Dim r               As Integer
r = 3
While Worksheets("INPUT").Cells(7, r) <> ""
    beforeitsarrange(r - 2) = Worksheets("INPUT").Cells(7, r)        'creats the array for the wights
    r = r + 1
Wend

Dim arrayafterallsorted() As Integer
arrayafterallsorted = ArrangeArray(sortedNumbers(), beforeitsarrange())

'now we have 2 arrays 1 with the right order of the mission and one of the mission time arrange by its wighet

Dim shibuz          As Boolean
shibuz = False
Dim d               As Integer
d = 0
While shibuz <> True
    shibuz = FindBalanceSolution(k, cformission4 + d, sortedNumbers, arrayafterallsorted)        ' sends to a function that checks if we can arrange in each station by the order of the wights - if we can we arrange else we send again with c+1
    d = d + 1
Wend
Worksheets("OUTPUT").Cells(3, "b") = cformission4 + d - 1
Worksheets("OUTPUT").Cells(5, "b") = LB4
If cformission4 + d - 1 = LB4 Then
    Worksheets("OUTPUT").Cells(4, "b") = "Yes"
    Range("B4").Interior.ColorIndex = 4
Else
    Worksheets("OUTPUT").Cells(4, "b") = "Not necessarily "
    Range("B4").Interior.ColorIndex = 45
End If
''''''''''''''''''''''''''
End If

End Sub

Sub ClearInput()
    '
    ' clear_Input
    '
    Range("C7:AF37,C2").Select
    Selection.ClearContents
    Range("C2:C4").Select
    Selection.ClearContents
    Range("A2").Select
    
End Sub
Sub ClearOutput()
    '
    ' ClearOutput
    
    Sheets("OUTPUT").Select
    Range("B3:B5").Select
    Selection.ClearContents
    Range("B11:AE41").Select
    Selection.ClearContents
    Range("A2").Select
End Sub

Public Function ArrangeArrayPositionsDescending(ByRef inputArray() As Integer) As Variant
    Dim positions() As Integer
    Dim sortedPositions() As Integer
    Dim i           As Integer
    
    ' Create an array to store the positions
    ReDim positions(1 To UBound(inputArray)) As Integer
    
    ' Assign positions to the array
    For i = 1 To UBound(inputArray)
        positions(i) = i
    Next i
    
    ' Sort the positions array based on the values in the input array
    sortedPositions = SortArrayPositionsDescending(positions, inputArray)
    
    ' Return the sorted positions array
    ArrangeArrayPositionsDescending = sortedPositions
End Function

Function SortArrayPositionsDescending(ByRef positions() As Integer, ByRef inputArray() As Integer) As Integer()
    Dim i           As Integer, j As Integer
    Dim temp        As Integer
    
    ' Perform Bubble Sort on the positions array in descending order based on the input array values
    For i = LBound(positions) To UBound(positions) - 1
        For j = i + 1 To UBound(positions)
            If inputArray(positions(j)) > inputArray(positions(i)) Then
                ' Swap positions
                temp = positions(i)
                positions(i) = positions(j)
                positions(j) = temp
            End If
        Next j
    Next i
    
    ' Return the sorted positions array
    SortArrayPositionsDescending = positions
End Function

Function ArrangeArray(ByRef arr1() As Integer, ByRef arr2() As Integer) As Integer()
    
    Dim lowerBound  As Long
    lowerBound = LBound(arr1)
    
    Dim upperBound  As Long
    upperBound = UBound(arr1)
    
    Dim arraylength As Long
    arraylength = upperBound - lowerBound + 2
    
    Dim arr3()      As Integer
    ReDim arr3(1 To (arraylength - 1))
    For i = 1 To arraylength - 1
        For j = 1 To arraylength - 1
            If arr1(i) = j Then
                arr3(i) = arr2(j)
            End If
        Next j
    Next i
    
    ArrangeArray = arr3
    
End Function

Function getarraylength(ByRef arr1() As Integer) As Integer
    Dim lowerBound  As Long
    lowerBound = LBound(arr1)
    
    Dim upperBound  As Long
    upperBound = UBound(arr1)
    
    Dim arraylength As Long
    arraylength = upperBound - lowerBound + 1
    
    getarraylength = arraylength
End Function

Function FindBalanceSolution(ByRef k As Integer, ByRef c As Integer, ByRef operations() As Integer, ByRef times() As Integer) As Boolean
    'clear output
    Call ClearOutput
    'clear output
    Dim arraylength As Integer
    arraylength = getarraylength(operations)
    Dim countertotaltimes As Integer
    countertotaltimes = 0
    For n = 1 To arraylength
        countertotaltimes = countertotaltimes + times(n)
    Next n
    Dim opertionhasbeen() As Boolean
    ReDim opertionhasbeen(1 To getarraylength(operations))
    Dim counterforshibuz As Integer
    counterforshibuz = 0
    Dim counterforshibuzperstation As Integer
    counterforshibuzperstation = 0
    Dim j           As Integer
    j = 1
    Dim m           As Integer
    m = 0
    
    For i = 1 To k
        m = j - 2
        counterforshibuzperstation = 0
        While counterforshibuzperstation + times(j) <= c
            
            counterforshibuzperstation = counterforshibuzperstation + times(j)
            counterforshibuz = counterforshibuz + times(j)
            If opertionhasbeen(j) <> True Then
                Worksheets("OUTPUT").Cells(10 + j - m - 1, i + 1) = operations(j)
                Worksheets("OUTPUT").Cells(41, i + 1) = Worksheets("OUTPUT").Cells(41, i + 1) + times(j)
                opertionhasbeen(j) = True
                j = j + 1
            End If
            
            If j > arraylength Then
                Exit For
            End If
        Wend
        
    Next i
    If (counterforshibuz = countertotaltimes) Then
        FindBalanceSolution = True
    Else
        FindBalanceSolution = False
    End If
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''ex2''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub Submission2()
    Dim k           As Integer
    Dim counter     As Integer
    Dim i           As Integer
    
    k = Worksheets("INPUT").Cells(2, "C")
    counter = 0
    i = 3
    
    ' Step 1: Find LB
    Dim totalWorkingTime As Double
    Dim longestJob  As Double
    Dim LB          As Integer
    
    totalWorkingTime = 0
    longestJob = 0
    
    ' Calculate total working time and find the longest job
    While Worksheets("INPUT").Cells(7, i) <> ""
        totalWorkingTime = totalWorkingTime + Worksheets("INPUT").Cells(7, i).Value
        If Worksheets("INPUT").Cells(7, i).Value > longestJob Then
            longestJob = Worksheets("INPUT").Cells(7, i).Value
        End If
        counter = counter + 1
        i = i + 1
    Wend
    
    i = i - 1
    ' Determine LB
    'LB = Application.WorksheetFunction.Max(Math.Ceiling(totalWorkingTime / K), longestJob)
    LB = Application.WorksheetFunction.Max(Application.WorksheetFunction.RoundUp(totalWorkingTime / k, 0), longestJob)
    
    ' Step 2: Find UB
    Dim UB          As Integer
    Dim totalWorkingTimeAllJobs As Integer
    
    ' Calculate total working time of all jobs
    totalWorkingTimeAllJobs = 0
    For i = 3 To counter + 2
        totalWorkingTimeAllJobs = totalWorkingTimeAllJobs + Worksheets("INPUT").Cells(7, i).Value
    Next i
    
    i = i - 1
    ' Set UB
    UB = totalWorkingTimeAllJobs
    
    ' Step 3: Place jobs in stations
    Dim c           As Integer
    ' Dim success   As Boolean
    
    c = LB
    
    '***START PUTING JOB IN STATION
    
    Worksheets("OUTPUT").Cells(5, "B") = c
    Dim ColumnINPUT As Integer
    ColumnINPUT = 3
    Dim RowOUTPUT   As Integer
    RowOUTPUT = 11
    Dim ColumnOUTPUT As Integer
    ColumnOUTPUT = 2
    Dim kCheak      As Integer
    kCheak = 1
    Dim timeLeft    As Integer
    timeLeft = c
    Dim CountIfShibuz As Integer
    CountIfShibuz = 0
    Dim WeHaveShibuz As Boolean
    WeHaveShibuz = False
    
    While WeHaveShibuz = False
    
        While Worksheets("INPUT").Cells(7, ColumnINPUT) <> "" And k >= kCheak
            If timeLeft >= Worksheets("INPUT").Cells(7, ColumnINPUT).Value Then
                Worksheets("OUTPUT").Cells(RowOUTPUT, ColumnOUTPUT).Value = Worksheets("INPUT").Cells(6, ColumnINPUT).Value
                Worksheets("OUTPUT").Cells(41, ColumnOUTPUT).Value = Worksheets("OUTPUT").Cells(41, ColumnOUTPUT).Value + Worksheets("INPUT").Cells(7, ColumnINPUT).Value
                RowOUTPUT = RowOUTPUT + 1
                timeLeft = timeLeft - Worksheets("INPUT").Cells(7, ColumnINPUT).Value
                CountIfShibuz = CountIfShibuz + 1
                ColumnINPUT = ColumnINPUT + 1
                
            Else
                kCheak = kCheak + 1
                ColumnOUTPUT = ColumnOUTPUT + 1
                RowOUTPUT = 11
                timeLeft = c
            End If
            
            If CountIfShibuz = counter Then
                WeHaveShibuz = True
            End If
            
        Wend
        c = c + 1
    Wend
    
    Worksheets("OUTPUT").Cells(4, "B") = "yes"
     Range("B4").Interior.ColorIndex = 4
    Worksheets("OUTPUT").Cells(3, "B") = c - 1
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''ex3''''''''''''''''''''''''''''''''''''''''''''''''''

Function checkforstations()        ' put 1 in the solver if we open a station 0 if not
    
    k = Worksheets("INPUT").Cells(2, 3)
    Dim i           As Integer
    i = 7
    Dim j           As Integer
    j = 1
    
    For i = 7 To 16
        If j <= k Then
            Worksheets("SOLVER").Cells(i, 15) = 1
            j = j + 1
        Else
            Worksheets("SOLVER").Cells(i, 15) = 0
        End If
    Next i
    
End Function

Function ex3(ByRef numofmission As Integer)        ' main function for 3
    
    Call clearthesolver
    
    Call checkforstations
    
    Worksheets("SOLVER").Cells(5, 4) = numofmission
    Worksheets("SOLVER").Cells(4, 4) = Worksheets("INPUT").Cells(2, 3)
    
    Call Solver
    
    Call insertLinearToOutput
    
    Sheets("OUTPUT").Cells(3, 2) = Worksheets("SOLVER").Cells(3, 4)
    Sheets("OUTPUT").Cells(4, 2) = "YES"
    Range("B4").Interior.ColorIndex = 4
    
    Call sumforeachstation
    
End Function

Sub clearthesolver()
    '
    ' clearLinear Macro
    Sheets("SOLVER").Select
    Range("C7:L16").Select
    Selection.ClearContents
    Range("O7:O16").Select
    Selection.ClearContents
    Range("D5").Select
    Selection.ClearContents
    
End Sub

Sub Solver()
    ' Solver Macro
    '
    Sheets("SOLVER").Select
    SolverOk SetCell:="$D$3", MaxMinVal:=2, ValueOf:=0, ByChange:= _
             "$C$7:$L$16,$D$3", Engine:=2, EngineDesc:="Simplex LP"
    SolverOk SetCell:="$D$3", MaxMinVal:=2, ValueOf:=0, ByChange:= _
             "$C$7:$L$16,$D$3", Engine:=2, EngineDesc:="Simplex LP"
    SolverSolve UserFinish:=True
    
    Sheets("OUTPUT").Select
End Sub

Sub insertLinearToOutput()
    Dim i           As Integer
    Dim j           As Integer        '
    Dim k           As Integer        '
    k = Worksheets("INPUT").Cells(2, 3)
    Dim row         As Integer        'station from OUTPUT
    row = 11
    Dim col         As Integer        'jobs from OUTPUT
    col = 2
    
    For i = 1 To k
        j = 1
        
        For j = 1 To 10
            If Worksheets("SOLVER").Cells(6 + i, 2 + j) = 1 Then
                Worksheets("OUTPUT").Cells(row, col) = j
                row = row + 1
            End If
        Next j
        row = 11
        col = col + 1
    Next i
    
End Sub

Sub sumforeachstation()
    k = Worksheets("INPUT").Cells(2, 3)
    Dim col         As Integer
    Dim sum         As Integer
    Dim i           As Integer
    Dim flag        As Boolean
    
    sum = 0
    flag = True
    col = 3
    i = 11
    j = 2
    
    While j <= k + 1
        While i < 41 And flag = True
            If Not IsEmpty(Worksheets("OUTPUT").Cells(i, j).Value) Then
                col = Worksheets("OUTPUT").Cells(i, j).Value
                i = i + 1
                Worksheets("OUTPUT").Cells(41, j) = sum + Worksheets("INPUT").Cells(7, col + 2).Value
                sum = Worksheets("OUTPUT").Cells(41, j)
            Else: flag = False
            End If
        Wend
        sum = 0
        j = j + 1
        i = 11
        flag = True
    Wend
End Sub

