Attribute VB_Name = "RefreshSalesData"
'================================================================
' SALES DATA REFRESH MODULE
' ------------------------------------------------
' Pulls raw sales data from DailySalesTransform.xlsx,
' filters by the date range in Date_Selector,
' and writes it into the Sales_Data sheet.
'
' The UNIQUE formulas in Saas_PO and PO_Qty_Calc
' handle deduplication and supplier filtering.
'
' Shortcut:  Ctrl+Shift+R  (assigned in Auto_Open)
'================================================================

Option Explicit

' ============================================================
' CONFIGURATION — Update these if your file/sheet names change
' ============================================================
Private Const SOURCE_PATH As String = "C:\Users\Peter\OneDrive - petergerard.com.au\Documents\Purchase_Order_Automation\DailySalesTransform.xlsx"
Private Const SOURCE_SHEET As String = "Transformed"

' Column positions in the SOURCE data (1-based)
Private Const COL_SUPPLIER   As Long = 1   ' Column A
Private Const COL_DATE       As Long = 2   ' Column B
Private Const COL_ITEM       As Long = 3   ' Column C
Private Const COL_DESC       As Long = 4   ' Column D
Private Const COL_QTY        As Long = 5   ' Column E

'================================================================
' PUBLIC: Refresh Sales_Data from external workbook
'================================================================
Public Sub RefreshSalesData()

    ' Read date range from Date_Selector
    Dim wsDateSel As Worksheet
    On Error Resume Next
    Set wsDateSel = ThisWorkbook.Sheets("Date_Selector")
    On Error GoTo 0

    If wsDateSel Is Nothing Then
        MsgBox "Date_Selector sheet not found.", vbExclamation
        Exit Sub
    End If

    Dim startDt As Date, endDt As Date

    ' C2 = start date, D2 = end date
    If Not IsDate(wsDateSel.Range("C2").Value) Then
        MsgBox "Start Date (Date_Selector C2) is not a valid date." & vbCrLf & _
               "Value found: " & CStr(wsDateSel.Range("C2").Value), vbExclamation
        Exit Sub
    End If
    If Not IsDate(wsDateSel.Range("D2").Value) Then
        MsgBox "End Date (Date_Selector D2) is not a valid date." & vbCrLf & _
               "Value found: " & CStr(wsDateSel.Range("D2").Value), vbExclamation
        Exit Sub
    End If

    startDt = CDate(wsDateSel.Range("C2").Value)
    endDt = CDate(wsDateSel.Range("D2").Value)

    If startDt > endDt Then
        MsgBox "Start Date (" & Format(startDt, "DD/MM/YYYY") & ") is after End Date (" & Format(endDt, "DD/MM/YYYY") & ").", vbExclamation
        Exit Sub
    End If

    ' Check if source file exists
    If Dir(SOURCE_PATH) = "" Then
        MsgBox "Source file not found:" & vbCrLf & vbCrLf & SOURCE_PATH & vbCrLf & vbCrLf & _
               "Check the SOURCE_PATH constant in the VBA module." & vbCrLf & _
               "(Alt+F11 to open VBA Editor, then open RefreshSalesData module)", _
               vbExclamation, "File Not Found"
        Exit Sub
    End If

    ' Open source workbook (read-only)
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = "Opening " & SOURCE_PATH & "..."

    Dim wbSource As Workbook
    On Error Resume Next
    Set wbSource = Workbooks.Open(Filename:=SOURCE_PATH, ReadOnly:=True, UpdateLinks:=0)
    On Error GoTo 0

    If wbSource Is Nothing Then
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Application.StatusBar = False
        MsgBox "Could not open the source workbook.", vbExclamation
        Exit Sub
    End If

    ' Find the source sheet
    Dim wsSource As Worksheet
    On Error Resume Next
    Set wsSource = wbSource.Sheets(SOURCE_SHEET)
    On Error GoTo 0

    If wsSource Is Nothing Then
        wbSource.Close SaveChanges:=False
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Application.StatusBar = False
        MsgBox "Sheet '" & SOURCE_SHEET & "' not found in source workbook." & vbCrLf & _
               "Available sheets: " & JoinSheetNames(wbSource), vbExclamation
        Exit Sub
    End If

    ' Read source data into array
    Dim lastRow As Long
    lastRow = wsSource.Cells(wsSource.Rows.Count, COL_ITEM).End(xlUp).Row

    If lastRow < 2 Then
        wbSource.Close SaveChanges:=False
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Application.StatusBar = False
        MsgBox "No data found in the source sheet.", vbInformation
        Exit Sub
    End If

    Application.StatusBar = "Reading " & (lastRow - 1) & " rows..."

    Dim srcData As Variant
    srcData = wsSource.Range(wsSource.Cells(2, 1), wsSource.Cells(lastRow, COL_QTY)).Value

    ' Close source workbook
    wbSource.Close SaveChanges:=False

    ' Get Sales_Data sheet
    Dim wsSalesData As Worksheet
    On Error Resume Next
    Set wsSalesData = ThisWorkbook.Sheets("Sales_Data")
    On Error GoTo 0

    If wsSalesData Is Nothing Then
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Application.StatusBar = False
        MsgBox "Sales_Data sheet not found in this workbook.", vbExclamation
        Exit Sub
    End If

    ' Clear old data (keep header row 1)
    Dim lastClear As Long
    lastClear = wsSalesData.Cells(wsSalesData.Rows.Count, COL_ITEM).End(xlUp).Row
    If lastClear >= 2 Then
        wsSalesData.Range(wsSalesData.Cells(2, 1), wsSalesData.Cells(lastClear, COL_QTY)).Clear
    End If

    ' === PASS 1: Filter by date range into output array ===
    Dim rowCount As Long, filteredCount As Long
    Dim i As Long
    Dim rowDate As Date

    rowCount = UBound(srcData, 1)
    filteredCount = 0

    ' Size output array to max possible
    Dim outData() As Variant
    ReDim outData(1 To rowCount, 1 To 5)

    Application.StatusBar = "Filtering " & rowCount & " rows..."

    For i = 1 To rowCount
        ' Skip rows with empty Item or Date
        If Not IsEmpty(srcData(i, COL_ITEM)) And Not IsEmpty(srcData(i, COL_DATE)) Then
            If IsDate(srcData(i, COL_DATE)) Then
                rowDate = CDate(srcData(i, COL_DATE))

                ' Date filter only — let formulas handle supplier filtering
                If rowDate >= startDt And rowDate <= endDt Then
                    filteredCount = filteredCount + 1

                    outData(filteredCount, 1) = srcData(i, COL_SUPPLIER)
                    outData(filteredCount, 2) = srcData(i, COL_DATE)
                    outData(filteredCount, 3) = srcData(i, COL_ITEM)
                    outData(filteredCount, 4) = srcData(i, COL_DESC)
                    outData(filteredCount, 5) = srcData(i, COL_QTY)
                End If
            End If
        End If
    Next i

    ' Handle zero results
    If filteredCount = 0 Then
        Dim sampleDate As String
        If rowCount > 0 And IsDate(srcData(1, COL_DATE)) Then
            sampleDate = Format(CDate(srcData(1, COL_DATE)), "DD/MM/YYYY")
        ElseIf rowCount > 0 Then
            sampleDate = CStr(srcData(1, COL_DATE)) & " (not a date)"
        End If

        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Application.StatusBar = False
        MsgBox "No rows matched the date range." & vbCrLf & vbCrLf & _
               "Filter: " & Format(startDt, "DD/MM/YYYY") & " to " & Format(endDt, "DD/MM/YYYY") & vbCrLf & _
               "Source rows: " & rowCount & vbCrLf & _
               "First row date: " & sampleDate, _
               vbExclamation, "No Results"
        Exit Sub
    End If

    ' === PASS 2: Bulk write to sheet in one shot ===
    Application.StatusBar = "Writing " & filteredCount & " rows..."

    ' Trim array to actual size
    Dim finalData() As Variant
    ReDim finalData(1 To filteredCount, 1 To 5)
    Dim j As Long
    For i = 1 To filteredCount
        For j = 1 To 5
            finalData(i, j) = outData(i, j)
        Next j
    Next i

    ' Bulk write — single operation
    wsSalesData.Range(wsSalesData.Cells(2, 1), wsSalesData.Cells(filteredCount + 1, 5)).Value = finalData

    ' Format date column
    wsSalesData.Range(wsSalesData.Cells(2, COL_DATE), wsSalesData.Cells(filteredCount + 1, COL_DATE)).NumberFormat = "D/MM/YYYY"

    ' Resize the Sales_Data table if it exists
    Dim lo As ListObject
    On Error Resume Next
    Set lo = wsSalesData.ListObjects("Sales_Data")
    On Error GoTo 0
    If Not lo Is Nothing Then
        lo.Resize wsSalesData.Range("A1:E" & (filteredCount + 1))
    End If

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False

    MsgBox "Sales Data refreshed!" & vbCrLf & vbCrLf & _
           "Date range: " & Format(startDt, "DD/MM/YYYY") & " to " & Format(endDt, "DD/MM/YYYY") & vbCrLf & _
           "Source rows scanned: " & rowCount & vbCrLf & _
           "Rows imported: " & filteredCount, _
           vbInformation, "Refresh Complete"

End Sub

'================================================================
' HELPER: Join sheet names for error message
'================================================================
Private Function JoinSheetNames(wb As Workbook) As String
    Dim s As String, ws As Worksheet
    For Each ws In wb.Sheets
        If Len(s) > 0 Then s = s & ", "
        s = s & "'" & ws.Name & "'"
    Next ws
    JoinSheetNames = s
End Function

'================================================================
' AUTO: Assign keyboard shortcuts on workbook open
'================================================================
Public Sub Auto_Open()
    Application.OnKey "+^r", "RefreshSalesData"   ' Ctrl+Shift+R = Sales Data
    Application.OnKey "+^d", "RefreshStockData"    ' Ctrl+Shift+D = Stock Data
End Sub

Public Sub Auto_Close()
    Application.OnKey "+^r"   ' Reset sales shortcut
    Application.OnKey "+^d"   ' Reset stock shortcut
End Sub
