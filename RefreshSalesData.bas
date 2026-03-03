Attribute VB_Name = "RefreshSalesData"
'================================================================
' SALES DATA REFRESH MODULE
' ------------------------------------------------
' Reads the raw QuickBooks "Sales Report" export directly,
' parses the hierarchical structure (Supplier > Item > Transactions),
' filters by the date range in Date_Selector,
' and writes flat data into the Sales_Data sheet.
'
' No Power Query or DailySalesTransform file needed!
'
' The UNIQUE formulas in Saas_PO and PO_Qty_Calc
' handle deduplication and supplier filtering.
'
' Shortcut:  Ctrl+Shift+R  (assigned in Auto_Open)
'================================================================

Option Explicit

' ============================================================
' CONFIGURATION — Update these if your file/path changes
' ============================================================
Private Const SOURCE_PATH As String = "C:\Users\Peter\OneDrive - petergerard.com.au\Documents\Purchase_Order_Automation\Sales Report Last Month To COB Yesterday.xlsx"
Private Const SOURCE_SHEET As String = "Sheet1"

' Row where data starts in the QB export (after title, date range, blank, headers)
Private Const QB_DATA_START_ROW As Long = 5

'================================================================
' PUBLIC: Refresh Sales_Data from raw QuickBooks sales export
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
        MsgBox "QuickBooks sales export not found:" & vbCrLf & vbCrLf & SOURCE_PATH & vbCrLf & vbCrLf & _
               "Check the SOURCE_PATH constant in the VBA module." & vbCrLf & _
               "(Alt+F11 to open VBA Editor, then open RefreshSalesData module)", _
               vbExclamation, "File Not Found"
        Exit Sub
    End If

    ' Disable screen updates and calculations for speed
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = "Opening QuickBooks sales export..."

    ' Open source workbook
    Dim wbSource As Workbook
    On Error Resume Next
    Set wbSource = Workbooks.Open(Filename:=SOURCE_PATH, ReadOnly:=True, UpdateLinks:=0)
    On Error GoTo 0

    If wbSource Is Nothing Then
        GoTo CleanupAndExit
    End If

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
        MsgBox "Sheet '" & SOURCE_SHEET & "' not found.", vbExclamation
        Exit Sub
    End If

    ' Read all source data into array
    Dim lastRow As Long
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row

    If lastRow < QB_DATA_START_ROW Then
        wbSource.Close SaveChanges:=False
        GoTo CleanupAndExit
    End If

    Application.StatusBar = "Reading " & lastRow & " rows..."
    Dim srcData As Variant
    srcData = wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(lastRow, 4)).Value

    wbSource.Close SaveChanges:=False

    ' === PARSE THE HIERARCHICAL QB STRUCTURE ===
    ' Structure:
    '   Supplier header row:  A = supplier name, B/C/D empty
    '   Item code row:        A = item code, B/C/D empty
    '   Transaction rows:     A empty, B = date, C = qty, D = description
    '   Total rows:           A = "Total for ..." or "TOTAL"

    Dim rowCount As Long
    rowCount = UBound(srcData, 1)

    ' Output array (max possible size)
    Dim outData() As Variant
    ReDim outData(1 To rowCount, 1 To 5)
    Dim outCount As Long
    outCount = 0

    Dim currentSupplier As String
    Dim currentItem As String
    Dim aVal As String, bVal As Variant, cVal As Variant, dVal As Variant
    Dim txDate As Date
    Dim i As Long

    currentSupplier = ""
    currentItem = ""

    Application.StatusBar = "Parsing QuickBooks data..."

    For i = QB_DATA_START_ROW To rowCount
        ' Read row values
        If IsEmpty(srcData(i, 1)) Then
            aVal = ""
        Else
            aVal = Trim(CStr(srcData(i, 1)))
        End If
        bVal = srcData(i, 2)
        cVal = srcData(i, 3)
        dVal = srcData(i, 4)

        ' Skip TOTAL rows and footer
        If UCase(aVal) = "TOTAL" Then GoTo NextQBRow
        If Left(aVal, 9) = "Total for" Then GoTo NextQBRow
        If Left(aVal, 1) = " " Then GoTo NextQBRow  ' Timestamp row
        If Left(aVal, 7) = "Accrual" Then GoTo NextQBRow

        ' Determine row type
        Dim hasDate As Boolean
        hasDate = False
        If Not IsEmpty(bVal) Then
            If IsDate(bVal) Then hasDate = True
        End If

        If Len(aVal) > 0 And Not hasDate And IsEmpty(cVal) And IsEmpty(dVal) Then
            ' This is either a SUPPLIER HEADER or an ITEM CODE row
            ' Supplier headers come before item codes
            ' We distinguish by checking if the next non-empty A-value row
            ' also has no B/C/D (then this is a supplier, next is an item)
            ' OR we can check if the next row has a date (then this is an item code)

            Dim nextHasDate As Boolean
            Dim nextAVal As String
            nextHasDate = False
            nextAVal = ""

            If i + 1 <= rowCount Then
                If Not IsEmpty(srcData(i + 1, 2)) Then
                    If IsDate(srcData(i + 1, 2)) Then nextHasDate = True
                End If
                If Not IsEmpty(srcData(i + 1, 1)) Then
                    nextAVal = Trim(CStr(srcData(i + 1, 1)))
                End If
            End If

            If nextHasDate Then
                ' Next row has a date = this row is an ITEM CODE
                currentItem = aVal
            ElseIf Len(nextAVal) > 0 And Not nextHasDate Then
                ' Next row has text in A but no date = this is a SUPPLIER HEADER
                ' (next row will be an item code)
                currentSupplier = aVal
                currentItem = ""
            Else
                ' Default: treat as item code
                currentItem = aVal
            End If

        ElseIf hasDate And Len(aVal) = 0 Then
            ' TRANSACTION ROW: has date in B, no value in A
            ' Only include if we have a valid supplier and item
            If Len(currentSupplier) > 0 And Len(currentItem) > 0 Then
                ' Check quantity exists
                If Not IsEmpty(cVal) And cVal <> "" Then
                    txDate = CDate(bVal)

                    ' Date filter
                    If txDate >= startDt And txDate <= endDt Then
                        outCount = outCount + 1

                        outData(outCount, 1) = currentSupplier         ' Supplier
                        outData(outCount, 2) = bVal                     ' Date
                        outData(outCount, 3) = currentItem              ' Item
                        If IsEmpty(dVal) Then
                            outData(outCount, 4) = ""
                        Else
                            outData(outCount, 4) = Trim(CStr(dVal))     ' Description
                        End If
                        outData(outCount, 5) = CDbl(cVal)               ' Qty
                    End If
                End If
            End If
        End If

NextQBRow:
    Next i

    ' === WRITE TO SALES_DATA ===
    Dim wsSalesData As Worksheet
    On Error Resume Next
    Set wsSalesData = ThisWorkbook.Sheets("Sales_Data")
    On Error GoTo 0

    If wsSalesData Is Nothing Then
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Application.StatusBar = False
        MsgBox "Sales_Data sheet not found.", vbExclamation
        Exit Sub
    End If

    ' Clear old data
    Dim lastClear As Long
    lastClear = wsSalesData.Cells(wsSalesData.Rows.Count, 3).End(xlUp).Row
    If lastClear >= 2 Then
        wsSalesData.Range(wsSalesData.Cells(2, 1), wsSalesData.Cells(lastClear, 5)).Clear
    End If

    If outCount = 0 Then
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Application.StatusBar = False
        MsgBox "No rows matched the date range." & vbCrLf & vbCrLf & _
               "Filter: " & Format(startDt, "DD/MM/YYYY") & " to " & Format(endDt, "DD/MM/YYYY") & vbCrLf & _
               "Source rows scanned: " & rowCount, _
               vbExclamation, "No Results"
        Exit Sub
    End If

    ' Trim output array and bulk write
    Application.StatusBar = "Writing " & outCount & " rows..."

    Dim finalData() As Variant
    ReDim finalData(1 To outCount, 1 To 5)
    Dim j As Long
    For i = 1 To outCount
        For j = 1 To 5
            finalData(i, j) = outData(i, j)
        Next j
    Next i

    wsSalesData.Range(wsSalesData.Cells(2, 1), wsSalesData.Cells(outCount + 1, 5)).Value = finalData

    ' Format date column
    wsSalesData.Range(wsSalesData.Cells(2, 2), wsSalesData.Cells(outCount + 1, 2)).NumberFormat = "D/MM/YYYY"

    ' Resize table if exists
    Dim lo As ListObject
    On Error Resume Next
    Set lo = wsSalesData.ListObjects("Sales_Data")
    On Error GoTo 0
    If Not lo Is Nothing Then
        lo.Resize wsSalesData.Range("A1:E" & (outCount + 1))
    End If

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False

    MsgBox "Sales Data refreshed!" & vbCrLf & vbCrLf & _
           "Date range: " & Format(startDt, "DD/MM/YYYY") & " to " & Format(endDt, "DD/MM/YYYY") & vbCrLf & _
           "Source rows scanned: " & rowCount & vbCrLf & _
           "Rows imported: " & outCount & vbCrLf & _
           "Suppliers found: " & currentSupplier & " (last)", _
           vbInformation, "Refresh Complete"

    Exit Sub

CleanupAndExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
    MsgBox "Could not open/read the source file.", vbExclamation

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
    Dim wb As String
    wb = "'" & ThisWorkbook.Name & "'!"

    Application.OnKey "+^r", wb & "RefreshSalesData.RefreshSalesData"    ' Ctrl+Shift+R
    Application.OnKey "+^d", wb & "RefreshStockData.RefreshStockData"    ' Ctrl+Shift+D
    Application.OnKey "+^e", wb & "POWorkflow.ExportPO"                  ' Ctrl+Shift+E
    Application.OnKey "+^n", wb & "POWorkflow.DetectNewItems"            ' Ctrl+Shift+N
    Application.OnKey "+^g", wb & "POWorkflow.CheckNegativeStock"        ' Ctrl+Shift+G
    Application.OnKey "+^a", wb & "POWorkflow.RunFullCycle"              ' Ctrl+Shift+A
End Sub

Public Sub Auto_Close()
    Application.OnKey "+^r"   ' Reset sales shortcut
    Application.OnKey "+^d"   ' Reset stock shortcut
    Application.OnKey "+^e"   ' Reset export shortcut
    Application.OnKey "+^n"   ' Reset new items shortcut
    Application.OnKey "+^g"   ' Reset negative stock shortcut
    Application.OnKey "+^a"   ' Reset full cycle shortcut
End Sub
