Attribute VB_Name = "modReporting"
Sub BudgetForecastReport()
    Dim wsSummary As Worksheet, wsBudget As Worksheet, wsForecast As Worksheet
    Dim lastRow As Long, i As Long, j As Long
    Dim category As String
    Dim avgSpend As Double, monthCount As Integer, projectedSpend As Double
    Dim remainingBudget As Double, projectedBalance As Double
    Dim dict As Object
    Dim fiscalYearEnd As Variant, todayDate As Date
    Dim monthsRemaining As Integer

    Set dict = CreateObject("Scripting.Dictionary")
    Set wsSummary = ThisWorkbook.Sheets("Summary Report")
    Set wsBudget = ThisWorkbook.Sheets("Budget Entry")

    ' Ask user for fiscal year end date
    fiscalYearEnd = InputBox("Enter your fiscal year end date (MM/DD/YYYY):", "Fiscal Year End")
    If Not IsDate(fiscalYearEnd) Then
        MsgBox "Invalid date format. Report cancelled."
        Exit Sub
    End If
    todayDate = Date
    monthsRemaining = DateDiff("m", todayDate, CDate(fiscalYearEnd)) + 1
    If monthsRemaining < 1 Then monthsRemaining = 1

    ' Create or clear Forecast sheet
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Budget Forecast").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsForecast = ThisWorkbook.Sheets.Add(After:=wsSummary)
    wsForecast.Name = "Budget Forecast"

    ' Build budget dictionary from Budget Entry
    lastRow = wsBudget.Cells(wsBudget.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        category = wsBudget.Cells(i, 2).Value
        If category <> "" And IsNumeric(wsBudget.Cells(i, 9).Value) Then
            remainingBudget = wsBudget.Cells(i, 9).Value
            dict(category) = remainingBudget
        End If
    Next i

    ' Set headers
    With wsForecast
        .Cells(1, 1).Value = "Category"
        .Cells(1, 2).Value = "Remaining Budget"
        .Cells(1, 3).Value = "Avg Monthly Spend"
        .Cells(1, 4).Value = "Projected Year Spend"
        .Cells(1, 5).Value = "Projected Over/(Under)"
    End With

    ' Process summary spending data
    lastRow = wsSummary.Cells(wsSummary.Rows.Count, 1).End(xlUp).Row
    j = 2
    For i = 2 To lastRow
        category = wsSummary.Cells(i, 1).Value
        If category <> "" Then
            monthCount = wsSummary.ListObjects(1).ListColumns.Count - 2 ' minus Category & Total
            avgSpend = 0
            Dim spendColCount As Integer: spendColCount = 0
            For colOut = 2 To monthCount + 1
                If IsNumeric(wsSummary.Cells(i, colOut).Value) Then
                    avgSpend = avgSpend + wsSummary.Cells(i, colOut).Value
                    spendColCount = spendColCount + 1
                End If
            Next colOut
            If spendColCount > 0 Then
                avgSpend = avgSpend / spendColCount
            Else
                avgSpend = 0
            End If
            projectedSpend = avgSpend * monthsRemaining

            If dict.exists(category) Then
                remainingBudget = dict(category)
                projectedBalance = remainingBudget - projectedSpend
                ' Skip Total Revenue if all forecast values are 0
                If Not (LCase(Trim(category)) = "total revenue" And remainingBudget = 0 And avgSpend = 0 And projectedSpend = 0) Then
                wsForecast.Cells(j, 1).Value = category
                wsForecast.Cells(j, 2).Value = remainingBudget
                wsForecast.Cells(j, 3).Value = avgSpend
                wsForecast.Cells(j, 4).Value = projectedSpend
                wsForecast.Cells(j, 5).Value = projectedBalance
                j = j + 1
            End If
            End If
        End If
    Next i

    With wsForecast.Range("B2:E" & j)
        .NumberFormat = "$#,##0.00"
        For Each cell In .Cells
            If IsNumeric(cell.Value) And cell.Value < 0 Then
                cell.Font.Color = vbRed
            Else
                cell.Font.Color = vbBlack
            End If
        Next cell
    End With

    wsForecast.Range("A1").CurrentRegion.Columns.AutoFit
    wsForecast.ListObjects.Add(xlSrcRange, wsForecast.Range("A1").CurrentRegion, , xlYes).Name = "ForecastTable"
    Dim forecastLastRow As Long
    forecastLastRow = wsForecast.Cells(wsForecast.Rows.Count, 1).End(xlUp).Row + 1
    wsForecast.Cells(forecastLastRow, 1).Value = "Total"
    For i = 2 To 5
        wsForecast.Cells(forecastLastRow, i).Formula = "=SUM(" & wsForecast.Cells(2, i).Address & ":" & wsForecast.Cells(forecastLastRow - 1, i).Address & ")"
        wsForecast.Cells(forecastLastRow, i).NumberFormat = "$#,##0.00"
        If wsForecast.Cells(forecastLastRow, i).Value < 0 Then
            wsForecast.Cells(forecastLastRow, i).Font.Color = vbRed
        End If
    Next i

    MsgBox "Budget Forecast report created."
End Sub
