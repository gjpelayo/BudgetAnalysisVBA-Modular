Attribute VB_Name = "modUtilities"
Function WorksheetExists(wsName As String) As Boolean
    Dim ws As Worksheet
    WorksheetExists = False
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = wsName Then
            WorksheetExists = True
            Exit Function
        End If
    Next ws
End Function

Sub QuickSort(arr() As String, first As Long, last As Long)
    Dim low As Long, high As Long
    Dim pivot As String, temp As String
    low = first
    high = last
    pivot = arr((first + last) \ 2)
    Do While low <= high
        Do While arr(low) < pivot
            low = low + 1
        Loop
        Do While arr(high) > pivot
            high = high - 1
        Loop
        If low <= high Then
            temp = arr(low)
            arr(low) = arr(high)
            arr(high) = temp
            low = low + 1
            high = high - 1
        End If
    Loop
    If first < high Then QuickSort arr, first, high
    If low < last Then QuickSort arr, low, last
End Sub
