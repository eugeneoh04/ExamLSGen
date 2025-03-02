Option Explicit

Sub GenList()
    'Setup input'
    Dim wsInput As Worksheet, wsOutput As Worksheet
    Dim numCol As Integer, numRow As Integer
    On Error Resume Next
    numCol = Application.InputBox(prompt:="How many categories of examples?", Type:=1)
    numRow = Application.InputBox(prompt:="How many examples in a question?", Type:=1)
    If numCol <= 0 Or numRow <= 0 Then
        MsgBox "Invalid input. Please enter positive integers.", vbCritical
        Exit Sub
    End If

    On Error GoTo 0

    Dim qText As String
    qText = Application.InputBox(prompt:="Question prompt?", Type:=2)
    
    Set wsInput = ThisWorkbook.Sheets("Sheet1")
    
    'Setup output'
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("Combinations").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set wsOutput = ThisWorkbook.Sheets.Add(After:=wsInput)
    wsOutput.Name = "Combinations"
    
    'Array variables'
    Dim colItems() As Variant
    Dim colPointer() As Long
    ReDim colItems(1 To numCol)
    ReDim colPointer(1 To numCol)
    Dim col As Integer, lastRow As Long, i As Long, outputRow As Long
    outputRow = 1
    
    'Shuffle examples in each category'
    If MsgBox("Shuffle examples in each category?", vbYesNo) = vbYes Then
        For col = 1 To numCol
            lastRow = wsInput.Cells(wsInput.Rows.Count, col).End(xlUp).Row
            If lastRow = 0 Then lastRow = 1
            colItems(col) = Application.WorksheetFunction.Transpose(wsInput.Range(wsInput.Cells(1, col), wsInput.Cells(lastRow, col)).Value)
            If Not IsArray(colItems(col)) Then
                Dim tempArr() As Variant
                ReDim tempArr(1 To 1)
                tempArr(1) = colItems(col)
                colItems(col) = tempArr
            End If
            ShuffleArray colItems(col)
            colPointer(col) = 1
        Next col
    End If
    
    'Main loop: repeat as examples'
    Do While True
        'Check categories for available examples'
        Dim availableCol As Collection
        Set availableCol = New Collection
        
        For col = 1 To numCol
            If colPointer(col) <= UBound(colItems(col)) Then
                availableCol.Add col
            End If
        Next col
        
        'If not enough available categories, exit loop'
        If availableCol.Count < numRow Then Exit Do
        
        'Sort columns in descending order based on available elements'
        ReDim tempArr(1 To availableCol.Count, 1 To 2)
        
        For i = 1 To availableCol.Count
            Dim boxIndex As Integer
            boxIndex = availableCol(i)
            tempArr(i, 1) = boxIndex
            tempArr(i, 2) = UBound(colItems(boxIndex)) - colPointer(boxIndex) + 1
        Next i
        
        Call SortDescending(tempArr)
        
        Dim topCol() As Integer
        ReDim topCol(1 To numRow)
        For i = 1 To numRow
            topCol(i) = tempArr(i, 1)
        Next i
        
        'Find smallest column count'
        Dim minCount As Long
        minCount = UBound(colItems(topCol(1))) - colPointer(topCol(1)) + 1
        For i = 2 To numRow
            Dim avail As Long
            avail = UBound(colItems(topCol(i))) - colPointer(topCol(i)) + 1
            If avail < minCount Then minCount = avail
        Next i
        
        'Create combination and output'
        Dim roundIndex As Long
        Dim combo() As Variant
        ReDim combo(1 To numRow)
        
        For roundIndex = 1 To minCount
            Dim comboText As String
            comboText = qText & vbCrLf
            
            For i = 1 To numRow
                comboText = comboText & i & ". " & colItems(topCol(i))(colPointer(topCol(i))) & vbCrLf
                colPointer(topCol(i)) = colPointer(topCol(i)) + 1
            Next i
            
            wsOutput.Cells(outputRow, 1).Value = comboText
            outputRow = outputRow + 1
        Next roundIndex
    Loop
    
    MsgBox "Finished generation"
End Sub

Sub ShuffleArray(ByRef arr As Variant)
    Dim i As Long, j As Long
    Dim temp As Variant
    Dim ub As Long
    ub = UBound(arr)
    Randomize
    For i = ub To 2 Step -1
        j = Int(Rnd * i) + 1
        temp = arr(i)
        arr(i) = arr(j)
        arr(j) = temp
    Next i
End Sub

Sub SortDescending(ByRef arr As Variant)
    Dim i As Long, j As Long
    Dim temp1 As Variant, temp2 As Variant
    For i = LBound(arr, 1) To UBound(arr, 1) - 1
        For j = i + 1 To UBound(arr, 1)
            If arr(i, 2) < arr(j, 2) Then
                temp1 = arr(i, 1)
                temp2 = arr(i, 2)
                arr(i, 1) = arr(j, 1)
                arr(i, 2) = arr(j, 2)
                arr(j, 1) = temp1
                arr(j, 2) = temp2
            End If
        Next j
    Next i
End Sub

