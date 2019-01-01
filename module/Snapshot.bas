Public Sub SnapshotGenerate()
    Call ChangeSnapshotDate
    Call CopyArchiveSnapshotValues
    Call CopyValuesFromShippedData
    Call CopyValuesFromUnshippedData
    Call ChangePivotTableFilter(Sheets("7.Pull Forward 50 s region"))
    Call ChangePivotTableFilter(Sheets("8.Pull Forward Customers"))
    ' Sheets("Snapshot").Activate
    MsgBox "Snapshot generation completed. Please valiate data before sending", vbOKOnly
End Sub

Public Sub CurrentEndDateOfFiscalMonth()
    MsgBox "End date of current fiscal month is " & EndDateOfFiscalMonth(Date), vbOKOnly
End Sub

Private Sub ChangeSnapshotDate()
    Sheets("Snapshot").Cells.Range("A1").Value = Date
End Sub

Private Sub CopyArchiveSnapshotValues()
    Sheets("Snapshot").Activate
    
    Range("D22:D647").Copy
    Range("I22").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    Range("L22:L647").Copy
    Range("O22").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    Range("S22:S647").Copy
    Range("U22").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
End Sub

Private Sub CopyValuesFromShippedData()
    Dim destinationWs As Worksheet
    Set destinationWs = ActiveWorkbook.Sheets("Shipped data")
    
    Dim wbName As String, wb As Workbook, ws As Worksheet
    
    Application.ScreenUpdating = False
    fName = ActiveWorkbook.Path & "/shipped" & GetCognosFileSuffix(Date) & ".xlsx"
    
    destinationWs.Columns("A:Q").Clear
    If Dir(fName) = "" Then
        MsgBox "File " & fName & " does not exist"
    Else
        Set wb = Workbooks.Open(fName)
        Set ws = wb.Sheets(1)
        If Not ws Is Nothing Then
            ws.Columns("A:Q").Copy Destination:=destinationWs.Columns("A:Q")
            wb.Close
        Else
            MsgBox ws.Name & " does not exist"
        End If
    End If
    Application.ScreenUpdating = True
    
    Call AutoFillRange(2, "R")
    Call AutoFillRange(2, "S")
    Call AutoFillRange(2, "T")
    Call AutoFillRange(2, "U")
End Sub

Private Sub CopyValuesFromUnshippedData()
    Dim destinationWs As Worksheet
    Set destinationWs = Sheets("Unshipped data")
    Dim wbName As String, wb As Workbook, ws As Worksheet
    Application.ScreenUpdating = False
    fName = ActiveWorkbook.Path & "/unshipped" & GetCognosFileSuffix(Date) & ".xlsx"
    
    destinationWs.Columns("A:N").ClearContents
    If Dir(fName) = "" Then
        MsgBox "File " & fName & " does not exist"
    Else
        Set wb = Workbooks.Open(fName)
        Set ws = wb.Sheets(1)
        If Not ws Is Nothing Then
            ws.Columns("A:N").Copy Destination:=destinationWs.Columns("A:N")
            wb.Close
        Else
            MsgBox ws.Name & " does not exist"
        End If
    End If
    Application.ScreenUpdating = True
    
    Call AutoFillRange(2, "O")
    Call AutoFillRange(2, "P")
    Call AutoFillRange(2, "Q")
    Call AutoFillRange(2, "R")
    
End Sub

Private Function GetCognosFileSuffix(desiredDate As Date) As String
    GetCognosFileSuffix = Day(desiredDate) & "." & Month(desiredDate)
End Function

Private Sub AutoFillRange(sourceRow As Integer, column As String)
    Range(column & sourceRow).Select
    Dim rowCount As Long
    rowCount = ActiveSheet.UsedRange.Rows.Count
    Selection.AutoFill Destination:=Range(column & sourceRow & ":" & column & rowCount)
End Sub

Private Sub ChangePivotTableFilter(ws As Worksheet)
    Dim pi As PivotItem
    Dim filterDate As Date
    Dim endOfFiscalMonth As Date: endOfFiscalMonth = EndDateOfFiscalMonth(Date)
    ws.PivotTables("PivotTable3").RefreshTable
    
    With ws.PivotTables("PivotTable3").PivotFields("Order Line Due Date")
        .ClearAllFilters
        
        For Each pi In .PivotItems
            If IsDate(pi.Value) Then
                filterDate = CDate(pi.Value)
                If DateDiff("d", endOfFiscalMonth, filterDate) > 0 And pi.Visible = False Then
                    pi.Visible = True
                ElseIf DateDiff("d", endOfFiscalMonth, filterDate) <= 0 And pi.Visible = True Then
                    pi.Visible = False
                End If
            End If
        Next pi
    End With
    ws.PivotTables("PivotTable3").RefreshTable
End Sub

Private Function EndDateOfFiscalMonth(desiredDate As Date)
    Dim lastSaturdayOfCurrentMonth As Date: lastSaturdayOfCurrentMonth = LastSaturdayOfTheMonth(desiredDate)
    If DateDiff("d", lastSaturdayOfCurrentMonth, desiredDate) > 0 Then
        ' Should be the last Saturday of next month
        desiredDate = WorksheetFunction.EoMonth(desiredDate, 0) + 1
        EndDateOfFiscalMonth = LastSaturdayOfTheMonth(desiredDate)
    Else
            Debug.Print desiredDate & " and " & lastSaturdayOfCurrentMonth & "ok"
        EndDateOfFiscalMonth = lastSaturdayOfCurrentMonth
    End If
End Function

Private Function LastSaturdayOfTheMonth(desiredDate As Date) As Date
    Dim endOfMonthDate As Date
    endOfMonthDate = WorksheetFunction.EoMonth(desiredDate, 0)
    LastSaturdayOfTheMonth = endOfMonthDate - Weekday(DateAdd("d", 7, endOfMonthDate))
End Function
