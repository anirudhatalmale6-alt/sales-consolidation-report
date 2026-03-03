Attribute VB_Name = "RefreshStockData"
'================================================================
' DAILY STOCK DATA REFRESH MODULE
' ------------------------------------------------
' Reads the QuickBooks "Product/Service List" export,
' transforms it into the Daily_Stock_Data format
' (Item, Description, Qty_On_Hand, Tax_Code),
' and writes it to the Daily_Stock_Data sheet.
'
' Handles:
'   - Skips QB header rows (title, business name, column headers)
'   - Skips TOTAL row and footer
'   - Extracts item code after colon for "Category:Code" items
'   - Treats blank quantities as 0
'
' Shortcut:  Ctrl+Shift+D  (assigned in Auto_Open)
'================================================================

Option Explicit

' ============================================================
' CONFIGURATION — Update these if your file/sheet names change
' ============================================================
Private Const STOCK_SOURCE_PATH As String = "C:\Users\Peter\OneDrive - petergerard.com.au\Documents\Purchase_Order_Automation\Product_Service List _Daily_6AM.xlsx"
Private Const STOCK_SOURCE_SHEET As String = "Sheet1"

' Row where actual data starts in the QB export (after title, name, blank, headers)
Private Const QB_DATA_START_ROW As Long = 5

' Column positions in the QB export (1-based)
Private Const QB_COL_ITEM     As Long = 1   ' Product/Service full name
Private Const QB_COL_DESC     As Long = 2   ' Memo/Description
Private Const QB_COL_QTY      As Long = 3   ' Quantity on hand
Private Const QB_COL_TAX      As Long = 4   ' GST Code

'================================================================
' PUBLIC: Refresh Daily_Stock_Data from QuickBooks export
'================================================================
Public Sub RefreshStockData()

    ' Check if source file exists
    If Dir(STOCK_SOURCE_PATH) = "" Then
        MsgBox "QuickBooks stock export not found:" & vbCrLf & vbCrLf & STOCK_SOURCE_PATH & vbCrLf & vbCrLf & _
               "Check the STOCK_SOURCE_PATH constant in the VBA module." & vbCrLf & _
               "(Alt+F11 to open VBA Editor, then open RefreshStockData module)", _
               vbExclamation, "File Not Found"
        Exit Sub
    End If

    ' Open source workbook (read-only)
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = "Opening QuickBooks stock export..."

    Dim wbSource As Workbook
    On Error Resume Next
    Set wbSource = Workbooks.Open(Filename:=STOCK_SOURCE_PATH, ReadOnly:=True, UpdateLinks:=0)
    On Error GoTo 0

    If wbSource Is Nothing Then
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Application.StatusBar = False
        MsgBox "Could not open the QuickBooks stock export.", vbExclamation
        Exit Sub
    End If

    ' Find the source sheet
    Dim wsSource As Worksheet
    On Error Resume Next
    Set wsSource = wbSource.Sheets(STOCK_SOURCE_SHEET)
    On Error GoTo 0

    If wsSource Is Nothing Then
        wbSource.Close SaveChanges:=False
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Application.StatusBar = False
        MsgBox "Sheet '" & STOCK_SOURCE_SHEET & "' not found in the export file.", vbExclamation
        Exit Sub
    End If

    ' Find last data row
    Dim lastRow As Long
    lastRow = wsSource.Cells(wsSource.Rows.Count, QB_COL_ITEM).End(xlUp).Row

    If lastRow < QB_DATA_START_ROW Then
        wbSource.Close SaveChanges:=False
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Application.StatusBar = False
        MsgBox "No data found in the export file.", vbInformation
        Exit Sub
    End If

    ' Read all data into array for speed
    Application.StatusBar = "Reading stock data..."
    Dim srcData As Variant
    srcData = wsSource.Range(wsSource.Cells(QB_DATA_START_ROW, 1), wsSource.Cells(lastRow, QB_COL_TAX)).Value

    ' Close source workbook
    wbSource.Close SaveChanges:=False

    ' Get Daily_Stock_Data sheet
    Dim wsStock As Worksheet
    On Error Resume Next
    Set wsStock = ThisWorkbook.Sheets("Daily_Stock_Data")
    On Error GoTo 0

    If wsStock Is Nothing Then
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Application.StatusBar = False
        MsgBox "Daily_Stock_Data sheet not found in this workbook.", vbExclamation
        Exit Sub
    End If

    ' Clear old data (keep header row 1)
    Dim lastClear As Long
    lastClear = wsStock.Cells(wsStock.Rows.Count, 1).End(xlUp).Row
    If lastClear >= 2 Then
        wsStock.Range(wsStock.Cells(2, 1), wsStock.Cells(lastClear, 4)).Clear
    End If

    ' === PASS 1: Process source data into output array ===
    Dim rowCount As Long, importCount As Long
    Dim i As Long
    Dim itemVal As String, colonPos As Long

    rowCount = UBound(srcData, 1)

    ' Size output array to max possible (will trim later)
    Dim outData() As Variant
    ReDim outData(1 To rowCount, 1 To 4)
    importCount = 0

    Application.StatusBar = "Processing " & rowCount & " rows..."

    For i = 1 To rowCount
        ' Get item value
        If IsEmpty(srcData(i, QB_COL_ITEM)) Then GoTo NextRow
        itemVal = Trim(CStr(srcData(i, QB_COL_ITEM)))

        ' Skip empty rows, TOTAL row, and timestamp rows
        If Len(itemVal) = 0 Then GoTo NextRow
        If UCase(itemVal) = "TOTAL" Then GoTo NextRow
        If Left(itemVal, 1) = " " Then GoTo NextRow

        ' Extract code after colon if present
        colonPos = InStrRev(itemVal, ":")
        If colonPos > 0 Then
            itemVal = Trim(Mid(itemVal, colonPos + 1))
        End If

        ' Skip if item is empty after extraction
        If Len(itemVal) = 0 Then GoTo NextRow

        importCount = importCount + 1

        ' Column 1: Item
        outData(importCount, 1) = itemVal

        ' Column 2: Description
        If IsEmpty(srcData(i, QB_COL_DESC)) Then
            outData(importCount, 2) = ""
        Else
            outData(importCount, 2) = Trim(CStr(srcData(i, QB_COL_DESC)))
        End If

        ' Column 3: Qty — treat blank as 0
        If IsEmpty(srcData(i, QB_COL_QTY)) Or srcData(i, QB_COL_QTY) = "" Then
            outData(importCount, 3) = 0
        ElseIf IsNumeric(srcData(i, QB_COL_QTY)) Then
            outData(importCount, 3) = CDbl(srcData(i, QB_COL_QTY))
        Else
            outData(importCount, 3) = 0
        End If

        ' Column 4: Tax code
        If IsEmpty(srcData(i, QB_COL_TAX)) Then
            outData(importCount, 4) = ""
        Else
            outData(importCount, 4) = Trim(CStr(srcData(i, QB_COL_TAX)))
        End If

NextRow:
    Next i

    ' === PASS 2: Write output array to sheet in one shot ===
    If importCount > 0 Then
        Application.StatusBar = "Writing " & importCount & " items..."

        ' Trim array to actual size
        Dim finalData() As Variant
        ReDim finalData(1 To importCount, 1 To 4)
        Dim j As Long
        For i = 1 To importCount
            For j = 1 To 4
                finalData(i, j) = outData(i, j)
            Next j
        Next i

        ' Bulk write — single operation, MUCH faster than cell-by-cell
        wsStock.Range(wsStock.Cells(2, 1), wsStock.Cells(importCount + 1, 4)).Value = finalData

        ' Format quantity column
        wsStock.Range(wsStock.Cells(2, 3), wsStock.Cells(importCount + 1, 3)).NumberFormat = "#,##0"
    End If

    ' Resize the table if it exists
    Dim lo As ListObject
    On Error Resume Next
    Set lo = wsStock.ListObjects(1)
    On Error GoTo 0
    If Not lo Is Nothing Then
        If importCount > 0 Then
            lo.Resize wsStock.Range("A1:D" & (importCount + 1))
        End If
    End If

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False

    MsgBox "Stock Data refreshed!" & vbCrLf & vbCrLf & _
           "Source rows scanned: " & rowCount & vbCrLf & _
           "Items imported: " & importCount, _
           vbInformation, "Stock Refresh Complete"

End Sub
