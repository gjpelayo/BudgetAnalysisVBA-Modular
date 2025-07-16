Attribute VB_Name = "modNavigationUI"
Sub RunGrantBudgetMacros()
    Call SplitGrantDataByCategory
    Call PopulateSummaryFromTotals
    Call BudgetForecastReport
    Call ColorCodeTabs
    Call AddNavigationHyperlinks
End Sub

Sub ResetBudgetSheets()
    Dim ws As Worksheet
    Application.ScreenUpdating = False
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Data Entry" And ws.Name <> "Budget Entry" Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws
    Application.ScreenUpdating = True
    MsgBox "Reset complete. Only 'Data Entry' and 'Budget Entry' sheets remain."
End Sub

Sub ColorCodeTabs()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        Select Case ws.Name
            Case "Data Entry", "Budget Entry"
                ws.Tab.Color = RGB(0, 112, 192) ' Blue for data sheets
            Case "Summary Report", "Budget Forecast"
                ws.Tab.Color = RGB(0, 176, 80) ' Green for reports
            Case Else
                If ws.Index > 2 Then ' Assume others are category sheets
                    ws.Tab.Color = RGB(255, 192, 0) ' Orange for categories
                End If
        End Select
    Next ws
End Sub

Sub AddNavigationHyperlinks()
    Dim wsSummary As Worksheet
    Dim tbl As ListObject
    Dim lastRow As Long, i As Long
    Dim catName As String
    Dim ws As Worksheet

    On Error Resume Next
    Set wsSummary = ThisWorkbook.Sheets("Summary Report")
    Set tbl = wsSummary.ListObjects("MonthlySpendingTable")
    On Error GoTo 0

    If tbl Is Nothing Then Exit Sub

    lastRow = tbl.ListRows.Count
    For i = 1 To lastRow
        catName = tbl.DataBodyRange(i, 1).Value
        Dim safeName As String
        safeName = Application.WorksheetFunction.Clean(Replace(Replace(Replace(catName, "/", "-"), "*", ""), ":", ""))
        safeName = Left(safeName, 31)
        If WorksheetExists(safeName) Then
            wsSummary.Hyperlinks.Add _
                Anchor:=tbl.DataBodyRange(i, 1), _
                Address:="", _
                SubAddress:="'" & safeName & "'!A1", _
                TextToDisplay:=catName
        End If
    Next i

    ' Add hyperlinks in Budget Forecast if it exists
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Budget Forecast")
    On Error GoTo 0

    If Not ws Is Nothing Then
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        For i = 2 To lastRow
            catName = ws.Cells(i, 1).Value
            Dim safeName2 As String
            safeName2 = Application.WorksheetFunction.Clean(Replace(Replace(Replace(catName, "/", "-"), "*", ""), ":", ""))
            safeName2 = Left(safeName2, 31)
            If WorksheetExists(safeName2) Then
                ws.Hyperlinks.Add _
                    Anchor:=ws.Cells(i, 1), _
                    Address:="", _
                    SubAddress:="'" & safeName2 & "'!A1", _
                    TextToDisplay:=catName
            End If
        Next i
    End If

    ' Add return links to each category sheet
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Summary Report" And WorksheetExists("Summary Report") Then
            If ws.Name <> "Data Entry" And ws.Name <> "Budget Entry" And ws.Name <> "Budget Forecast" And ws.Name <> "Home" Then
                With ws
                    .Hyperlinks.Add Anchor:=.Cells(2, Columns.Count).End(xlToLeft).Offset(0, 2), _
                        Address:="", SubAddress:="'Data Entry'!A1", TextToDisplay:="Return to Home"
                End With
            End If
        End If
    Next ws
End Sub
