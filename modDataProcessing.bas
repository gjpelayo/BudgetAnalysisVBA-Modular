Attribute VB_Name = "modDataProcessing"
Sub SplitGrantDataByCategory()
    Dim wsRaw As Worksheet
    Dim wsSummary As Worksheet
    Dim catSheet As Worksheet
    Dim lastRow As Long, i As Long
    Dim sumRow As Long
    Dim category As String, key As Variant
    Dim tCount As Long
    Dim tSum As Double
    Dim dict As Object
    Dim ws As Worksheet
    Dim destRow As Long
    Dim safeSheetName As String

    Set dict = CreateObject("Scripting.Dictionary")
    Application.ScreenUpdating = False

    ' Set Data Entry sheet
    Set wsRaw = ThisWorkbook.Sheets("Data Entry")

        lastRow = wsRaw.Cells(wsRaw.Rows.Count, 1).End(xlUp).Row

    ' Delete old category sheets (excluding Home, Data Entry, Budget Entry, Summary)
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Home" And ws.Name <> "Data Entry" And ws.Name <> "Budget Entry" Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
    Next ws

    ' Headers for Summary
    sumRow = 2

    ' Loop through Data Entry
    For i = 2 To lastRow ' assuming row 1 is headers
        If LCase(Left(Trim(wsRaw.Cells(i, 1).Text), 6)) = "total:" Then GoTo NextRow

        category = Trim(wsRaw.Cells(i, 2).Text)
        If category = "" Or LCase(Left(category, 6)) = "total:" Then GoTo NextRow

        key = category
        safeSheetName = Application.WorksheetFunction.Clean(Replace(Replace(Replace(key, "/", "-"), "*", ""), ":", ""))
        safeSheetName = Left(safeSheetName, 31) ' Sheet name limit

        If Not dict.exists(safeSheetName) Then
            Set catSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            On Error Resume Next
            catSheet.Name = safeSheetName
            If Err.Number <> 0 Then
                Err.Clear
                safeSheetName = "Sheet" & Format(Now, "hhmmss") & Rnd() * 1000
                catSheet.Name = safeSheetName
            End If
            On Error GoTo 0
            dict.Add safeSheetName, catSheet
            ' Add headers
            wsRaw.Rows(1).Copy Destination:=catSheet.Rows(1)
            ' Create table
            Dim lastCol As Long
            lastCol = catSheet.Cells(1, catSheet.Columns.Count).End(xlToLeft).Column
            Dim tblRange As Range
            Set tblRange = catSheet.Range(catSheet.Cells(1, 1), catSheet.Cells(1, lastCol))
            If catSheet.ListObjects.Count = 0 Then
                catSheet.ListObjects.Add(xlSrcRange, tblRange, , xlYes).Name = "tbl_" & safeSheetName
            End If
            ' Add back-link to Summary
            With catSheet
                .Hyperlinks.Add Anchor:=.Cells(1, Columns.Count).End(xlToLeft).Offset(0, 2), _
                    Address:="", SubAddress:="'Summary Report'!A1", TextToDisplay:="Return to Summary"
            End With
        Else
            Set catSheet = dict(safeSheetName)
        End If

        ' Find next empty row in category sheet and copy data
        destRow = catSheet.Cells(catSheet.Rows.Count, 1).End(xlUp).Row + 1
        wsRaw.Rows(i).Copy Destination:=catSheet.Rows(destRow)
        ' Update table range
        On Error Resume Next
        With catSheet.ListObjects(1)
            .Resize catSheet.Range(catSheet.Cells(1, 1), catSheet.Cells(catSheet.Cells(catSheet.Rows.Count, 1).End(xlUp).Row, lastCol))
        End With
        On Error GoTo 0

        ' No placeholder needed anymore

        ' Turn off wrap and autofit
        With catSheet.Cells
            .WrapText = False
            .EntireColumn.AutoFit
        End With
NextRow:
    Next i

    ' Add total row to each category sheet
    Dim catWs As Worksheet
    For Each key In dict.Keys
        Set catWs = dict(key)
        Dim amtCol As Long
        Dim headerCell As Range

        ' Find the Amount column by header
        Set headerCell = catWs.Rows(1).Find(What:="Amount", LookIn:=xlValues, LookAt:=xlWhole)
        If Not headerCell Is Nothing Then
            amtCol = headerCell.Column
            Dim lastDataRow As Long
            lastDataRow = catWs.Cells(catWs.Rows.Count, 1).End(xlUp).Row
            catWs.Cells(lastDataRow + 1, 2).Value = "Total: " & key
            catWs.Cells(lastDataRow + 1, amtCol).Formula = "=SUM(" & Cells(2, amtCol).Address(False, False) & ":" & Cells(lastDataRow, amtCol).Address(False, False) & ")"
            catWs.Cells(lastDataRow + 1, amtCol).NumberFormat = "$#,##0.00"
            catWs.Cells(lastDataRow + 1, amtCol).Font.Bold = True
        End If
    Next key

    Application.ScreenUpdating = True
    ' Color-code tabs
    For Each ws In ThisWorkbook.Sheets
        Select Case ws.Name
            Case "Data Entry", "Budget Entry"
                ws.Tab.Color = RGB(0, 112, 192) ' Blue for data sheets
            Case "Summary Report", "Budget Forecast", "Summary Report"
                ws.Tab.Color = RGB(0, 176, 80) ' Green for reports
            Case Else
                If ws.Index > 2 Then ' Assume others are category sheets
                    ws.Tab.Color = RGB(255, 192, 0) ' Orange for categories
                End If
        End Select
    Next ws

    MsgBox "Grant data split complete!"
End Sub

Sub PopulateSummaryFromTotals()
    Dim wsRaw As Worksheet
    Dim wsMonthly As Worksheet
    Dim lastRow As Long, i As Long
    Dim dateVal As Variant, category As String, amount As Double
    Dim monthKey As String, catKey As String
    Dim dict As Object, catList As Object, monthList As Object
    Dim rowOut As Long, colOut As Long
    Dim month As Variant, cat As Variant
    Dim monthsSorted() As String

    Set dict = CreateObject("Scripting.Dictionary")
    Set catList = CreateObject("Scripting.Dictionary")
    Set monthList = CreateObject("Scripting.Dictionary")

    ' Include all categories from Budget Entry
    Dim wsBudget As Worksheet
    Set wsBudget = ThisWorkbook.Sheets("Budget Entry")
    Dim budgetLastRow As Long
    budgetLastRow = wsBudget.Cells(wsBudget.Rows.Count, 1).End(xlUp).Row
    Dim bCat As String
    For i = 2 To budgetLastRow
        bCat = Trim(wsBudget.Cells(i, 2).Value)
        If bCat <> "" And Not catList.exists(bCat) Then catList.Add bCat, True
    Next i
    Set wsRaw = ThisWorkbook.Sheets("Data Entry")

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Monthly Spending").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsMonthly = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsMonthly.Name = "Summary Report"

    lastRow = wsRaw.Cells(wsRaw.Rows.Count, 1).End(xlUp).Row

    ' Collect data
    For i = 2 To lastRow
        dateVal = wsRaw.Cells(i, 3).Value
        category = wsRaw.Cells(i, 2).Value
        amount = wsRaw.Cells(i, 8).Value

        If IsDate(dateVal) And IsNumeric(amount) And category <> "" Then
            monthKey = Format(dateVal, "yyyy-mm")
            catKey = category

            If Not dict.exists(catKey) Then
                Set dict(catKey) = CreateObject("Scripting.Dictionary")
            End If
            If dict(catKey).exists(monthKey) Then
                dict(catKey)(monthKey) = dict(catKey)(monthKey) + amount
            Else
                dict(catKey).Add monthKey, amount
            End If

            If Not catList.exists(catKey) Then catList.Add catKey, True
            If Not monthList.exists(monthKey) Then monthList.Add monthKey, True
        End If
    Next i

    ' Sort months
    ReDim monthsSorted(0 To monthList.Count - 1)
    i = 0
    For Each month In monthList.Keys
        monthsSorted(i) = month
        i = i + 1
    Next month
    Call QuickSort(monthsSorted, LBound(monthsSorted), UBound(monthsSorted))

    ' Output headers
    wsMonthly.Cells(1, 1).Value = "Category"
    For i = 0 To UBound(monthsSorted)
        wsMonthly.Cells(1, i + 2).Value = Format(DateValue(monthsSorted(i) & "-01"), "mmmm yyyy")
    Next i
    wsMonthly.Cells(1, UBound(monthsSorted) + 2 + 1).Value = "Total"

    ' Output data
    rowOut = 2
    For Each cat In catList.Keys
        ' Skip revenue-type categories based on name or GL prefix
        Dim matchRow As Variant
        matchRow = Application.Match(cat, wsRaw.Columns(2), 0)
        If Not IsError(matchRow) Then
            Dim glCode As String
            glCode = Trim(wsRaw.Cells(matchRow + 1, 1).Text)
            If LCase(cat) Like "*revenue*" Or Left(glCode, 1) = "4" Then GoTo SkipCat
        End If

        If LCase(Trim(cat)) = "total revenue" Then GoTo SkipCat
        wsMonthly.Cells(rowOut, 1).Value = cat
        Dim totalAmt As Double: totalAmt = 0
        For i = 0 To UBound(monthsSorted)
            amount = 0
            If dict.exists(cat) Then
                If dict(cat).exists(monthsSorted(i)) Then
                    amount = dict(cat)(monthsSorted(i))
                End If
            End If

            wsMonthly.Cells(rowOut, i + 2).Value = amount
            totalAmt = totalAmt + amount
        Next i
        wsMonthly.Cells(rowOut, UBound(monthsSorted) + 2 + 1).Value = totalAmt
        rowOut = rowOut + 1
SkipCat:
    Next cat

    With wsMonthly.Range("B2:Z" & rowOut)
        .NumberFormat = "$#,##0.00"
        For Each totalCell In .Cells
            If IsNumeric(totalCell.Value) And totalCell.Value < 0 Then
                totalCell.Font.Color = vbRed
            Else
                totalCell.Font.Color = vbBlack
            End If
        Next totalCell
    End With

    wsMonthly.Range("A1").CurrentRegion.Columns.AutoFit
    Dim totalRow As Long
    totalRow = wsMonthly.Cells(wsMonthly.Rows.Count, 1).End(xlUp).Row + 1
    wsMonthly.Cells(totalRow, 1).Value = "Total"
    For i = 2 To UBound(monthsSorted) + 3
        wsMonthly.Cells(totalRow, i).Formula = "=SUM(" & wsMonthly.Cells(2, i).Address & ":" & wsMonthly.Cells(totalRow - 1, i).Address & ")"
        wsMonthly.Cells(totalRow, i).NumberFormat = "$#,##0.00"
        If wsMonthly.Cells(totalRow, i).Value < 0 Then
            wsMonthly.Cells(totalRow, i).Font.Color = vbRed
        End If
    Next i

    wsMonthly.ListObjects.Add(xlSrcRange, wsMonthly.Range("A1").CurrentRegion, , xlYes).Name = "MonthlySpendingTable"

    MsgBox "Monthly Spending sheet created."
End Sub
