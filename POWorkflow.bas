Attribute VB_Name = "POWorkflow"
'================================================================
' PO WORKFLOW MODULE
' ------------------------------------------------
' Contains:
'   1. ExportPO      — Exports Saas_PO as .xlsx + .pdf
'   2. DetectNewItems — Flags items not in Master_Stock_List
'   3. MoveToMaster   — Moves reviewed New_Items into Master
'   4. CheckNegativeStock — Reports negative stock for floor check
'   5. RunFullCycle   — One-click: Refresh + Check + Export
'
' Shortcuts (assigned in Auto_Open):
'   Ctrl+Shift+E  = ExportPO
'   Ctrl+Shift+N  = DetectNewItems
'   Ctrl+Shift+G  = CheckNegativeStock
'   Ctrl+Shift+A  = RunFullCycle
'================================================================

Option Explicit

' ============================================================
' CONFIGURATION
' ============================================================
Private Const EXPORT_FOLDER As String = "C:\Users\Peter\OneDrive - petergerard.com.au\Documents\Daily Saasant Uploads\"

'================================================================
' 1. EXPORT PO — Save Saas_PO as .xlsx and .pdf
'    Filename: SupplierName_PO_YYYY-MM-DD
'    Shortcut: Ctrl+Shift+E
'================================================================
Public Sub ExportPO()

    ' Get supplier name from Date_Selector
    Dim wsDateSel As Worksheet
    On Error Resume Next
    Set wsDateSel = ThisWorkbook.Sheets("Date_Selector")
    On Error GoTo 0

    If wsDateSel Is Nothing Then
        MsgBox "Date_Selector sheet not found.", vbExclamation
        Exit Sub
    End If

    Dim supplierName As String
    supplierName = Trim(CStr(wsDateSel.Range("A2").Value))
    If Len(supplierName) = 0 Then
        MsgBox "No supplier selected in Date_Selector A2.", vbExclamation
        Exit Sub
    End If

    ' Get Saas_PO sheet
    Dim wsPO As Worksheet
    On Error Resume Next
    Set wsPO = ThisWorkbook.Sheets("Saas_PO")
    On Error GoTo 0

    If wsPO Is Nothing Then
        MsgBox "Saas_PO sheet not found.", vbExclamation
        Exit Sub
    End If

    ' Check Saas_PO has data
    Dim lastRow As Long
    lastRow = wsPO.Cells(wsPO.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "Saas_PO has no data to export.", vbExclamation
        Exit Sub
    End If

    ' Build filename
    Dim dateStr As String
    dateStr = Format(Date, "YYYY-MM-DD")

    ' Clean supplier name for filename (remove invalid chars)
    Dim cleanSupplier As String
    cleanSupplier = CleanFileName(supplierName)

    Dim baseName As String
    baseName = cleanSupplier & "_PO_" & dateStr

    ' Check export folder exists, create if not
    If Dir(EXPORT_FOLDER, vbDirectory) = "" Then
        MkDir EXPORT_FOLDER
    End If

    Dim xlsxPath As String
    Dim pdfPath As String
    xlsxPath = EXPORT_FOLDER & baseName & ".xlsx"
    pdfPath = EXPORT_FOLDER & baseName & ".pdf"

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' === Export as XLSX ===
    ' Copy Saas_PO to a new workbook (values only to strip formulas)
    wsPO.Copy
    Dim wbExport As Workbook
    Set wbExport = ActiveWorkbook

    ' Convert formulas to values
    Dim wsExport As Worksheet
    Set wsExport = wbExport.Sheets(1)
    Dim usedRange As Range
    Set usedRange = wsExport.UsedRange
    usedRange.Copy
    usedRange.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    ' Auto-fit columns for readability
    usedRange.Columns.AutoFit

    ' Save as xlsx
    wbExport.SaveAs Filename:=xlsxPath, FileFormat:=xlOpenXMLWorkbook
    wbExport.Close SaveChanges:=False

    ' === Export as PDF ===
    wsPO.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "PO exported successfully!" & vbCrLf & vbCrLf & _
           "Supplier: " & supplierName & vbCrLf & _
           "XLSX: " & xlsxPath & vbCrLf & _
           "PDF: " & pdfPath, _
           vbInformation, "Export Complete"

End Sub

'================================================================
' 2. DETECT NEW ITEMS — Find items in Sales_Data or Saas_PO
'    that don't exist in Master_Stock_List
'    Shortcut: Ctrl+Shift+N
'================================================================
Public Sub DetectNewItems()

    ' Get sheets
    Dim wsSalesData As Worksheet, wsMaster As Worksheet
    On Error Resume Next
    Set wsSalesData = ThisWorkbook.Sheets("Sales_Data")
    Set wsMaster = ThisWorkbook.Sheets("Master_Stock_List")
    On Error GoTo 0

    If wsSalesData Is Nothing Then
        MsgBox "Sales_Data sheet not found.", vbExclamation
        Exit Sub
    End If
    If wsMaster Is Nothing Then
        MsgBox "Master_Stock_List sheet not found.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = "Scanning for new items..."

    ' Read all Master_Stock_List item codes into a dictionary
    Dim dictMaster As Object
    Set dictMaster = CreateObject("Scripting.Dictionary")
    dictMaster.CompareMode = vbTextCompare  ' Case-insensitive

    Dim masterLastRow As Long
    masterLastRow = wsMaster.Cells(wsMaster.Rows.Count, 2).End(xlUp).Row

    If masterLastRow >= 2 Then
        Dim masterData As Variant
        masterData = wsMaster.Range("B2:B" & masterLastRow).Value

        Dim i As Long
        If IsArray(masterData) Then
            For i = 1 To UBound(masterData, 1)
                If Not IsEmpty(masterData(i, 1)) Then
                    Dim masterCode As String
                    masterCode = Trim(CStr(masterData(i, 1)))
                    If Len(masterCode) > 0 Then
                        If Not dictMaster.Exists(masterCode) Then
                            dictMaster.Add masterCode, True
                        End If
                    End If
                End If
            Next i
        End If
    End If

    ' Read Sales_Data items (column C = Item code, A = Supplier, D = Description)
    Dim salesLastRow As Long
    salesLastRow = wsSalesData.Cells(wsSalesData.Rows.Count, 3).End(xlUp).Row

    If salesLastRow < 2 Then
        GoTo CleanupNewItems
    End If

    Dim salesData As Variant
    salesData = wsSalesData.Range("A2:E" & salesLastRow).Value

    ' Find items not in master
    Dim dictNew As Object
    Set dictNew = CreateObject("Scripting.Dictionary")
    dictNew.CompareMode = vbTextCompare

    ' Store new item info: key = item code, value = Array(supplier, code, description)
    Dim itemCode As String, itemSupplier As String, itemDesc As String

    For i = 1 To UBound(salesData, 1)
        If Not IsEmpty(salesData(i, 3)) Then
            itemCode = Trim(CStr(salesData(i, 3)))
            If Len(itemCode) > 0 Then
                If Not dictMaster.Exists(itemCode) And Not dictNew.Exists(itemCode) Then
                    ' New item found
                    itemSupplier = ""
                    itemDesc = ""
                    If Not IsEmpty(salesData(i, 1)) Then itemSupplier = Trim(CStr(salesData(i, 1)))
                    If Not IsEmpty(salesData(i, 4)) Then itemDesc = Trim(CStr(salesData(i, 4)))
                    dictNew.Add itemCode, Array(itemSupplier, itemCode, itemDesc)
                End If
            End If
        End If
    Next i

    ' Get or create New_Items sheet
    Dim wsNew As Worksheet
    On Error Resume Next
    Set wsNew = ThisWorkbook.Sheets("New_Items")
    On Error GoTo 0

    If wsNew Is Nothing Then
        ' Create the sheet
        Set wsNew = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsNew.Name = "New_Items"

        ' Add headers
        wsNew.Range("A1").Value = "Supplier (QB Category)"
        wsNew.Range("B1").Value = "Code"
        wsNew.Range("C1").Value = "Description"
        wsNew.Range("D1").Value = "Max Shelf Qty"
        wsNew.Range("E1").Value = "Supplier Break Qty"
        wsNew.Range("F1").Value = "Location"
        wsNew.Range("G1").Value = "Date Detected"

        ' Format header row
        With wsNew.Range("A1:G1")
            .Font.Bold = True
            .Interior.Color = RGB(0, 102, 153)  ' Dark teal
            .Font.Color = RGB(255, 255, 255)     ' White text
        End With
    End If

    ' Read existing New_Items codes to avoid duplicates
    Dim dictExistingNew As Object
    Set dictExistingNew = CreateObject("Scripting.Dictionary")
    dictExistingNew.CompareMode = vbTextCompare

    Dim newLastRow As Long
    newLastRow = wsNew.Cells(wsNew.Rows.Count, 2).End(xlUp).Row

    If newLastRow >= 2 Then
        Dim existingData As Variant
        existingData = wsNew.Range("B2:B" & newLastRow).Value
        If IsArray(existingData) Then
            For i = 1 To UBound(existingData, 1)
                If Not IsEmpty(existingData(i, 1)) Then
                    Dim existCode As String
                    existCode = Trim(CStr(existingData(i, 1)))
                    If Len(existCode) > 0 And Not dictExistingNew.Exists(existCode) Then
                        dictExistingNew.Add existCode, True
                    End If
                End If
            Next i
        End If
    End If

    ' Write new items (skip any already on the New_Items sheet)
    Dim writeRow As Long
    writeRow = newLastRow + 1
    Dim addedCount As Long
    addedCount = 0

    Dim key As Variant
    For Each key In dictNew.Keys
        If Not dictExistingNew.Exists(CStr(key)) Then
            Dim itemInfo As Variant
            itemInfo = dictNew(key)

            wsNew.Cells(writeRow, 1).Value = itemInfo(0)  ' Supplier
            wsNew.Cells(writeRow, 2).Value = itemInfo(1)  ' Code
            wsNew.Cells(writeRow, 3).Value = itemInfo(2)  ' Description
            ' D, E, F left blank for user to fill in
            wsNew.Cells(writeRow, 7).Value = Date          ' Date Detected
            wsNew.Cells(writeRow, 7).NumberFormat = "DD/MM/YYYY"

            ' Highlight blank cells that need attention
            wsNew.Range(wsNew.Cells(writeRow, 4), wsNew.Cells(writeRow, 6)).Interior.Color = RGB(255, 255, 200)  ' Light yellow

            writeRow = writeRow + 1
            addedCount = addedCount + 1
        End If
    Next key

    ' Auto-fit columns
    wsNew.UsedRange.Columns.AutoFit

CleanupNewItems:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False

    If dictNew Is Nothing Then
        MsgBox "No sales data to check.", vbInformation
    ElseIf dictNew.Count = 0 Then
        MsgBox "All items are already in the Master Stock List. No new items found.", _
               vbInformation, "New Items Check"
    ElseIf addedCount = 0 Then
        MsgBox dictNew.Count & " new item(s) were found but already on the New_Items sheet." & vbCrLf & _
               "Fill in Max Shelf Qty, Break Qty, and Location then run 'Move to Master'.", _
               vbInformation, "New Items Check"
    Else
        MsgBox addedCount & " new item(s) added to New_Items sheet!" & vbCrLf & vbCrLf & _
               "Please fill in:" & vbCrLf & _
               "  - Max Shelf Qty (column D)" & vbCrLf & _
               "  - Supplier Break Qty (column E)" & vbCrLf & _
               "  - Location (column F)" & vbCrLf & vbCrLf & _
               "Yellow cells need your input." & vbCrLf & _
               "When done, run 'Move to Master' to transfer them.", _
               vbInformation, "New Items Detected"
        wsNew.Activate
    End If

End Sub

'================================================================
' 3. MOVE TO MASTER — Transfer completed New_Items rows
'    into Master_Stock_List (only rows where D+E are filled)
'================================================================
Public Sub MoveToMaster()

    Dim wsNew As Worksheet, wsMaster As Worksheet
    On Error Resume Next
    Set wsNew = ThisWorkbook.Sheets("New_Items")
    Set wsMaster = ThisWorkbook.Sheets("Master_Stock_List")
    On Error GoTo 0

    If wsNew Is Nothing Then
        MsgBox "New_Items sheet not found. Run 'Detect New Items' first.", vbExclamation
        Exit Sub
    End If
    If wsMaster Is Nothing Then
        MsgBox "Master_Stock_List sheet not found.", vbExclamation
        Exit Sub
    End If

    Dim newLastRow As Long
    newLastRow = wsNew.Cells(wsNew.Rows.Count, 2).End(xlUp).Row

    If newLastRow < 2 Then
        MsgBox "No items on the New_Items sheet to move.", vbInformation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' Find last row in Master_Stock_List
    Dim masterLastRow As Long
    masterLastRow = wsMaster.Cells(wsMaster.Rows.Count, 2).End(xlUp).Row
    Dim masterWriteRow As Long
    masterWriteRow = masterLastRow + 1

    ' Process New_Items rows from bottom to top (so we can delete moved rows)
    Dim movedCount As Long, skippedCount As Long
    movedCount = 0
    skippedCount = 0

    Dim i As Long
    For i = newLastRow To 2 Step -1
        Dim maxQty As Variant, breakQty As Variant
        maxQty = wsNew.Cells(i, 4).Value
        breakQty = wsNew.Cells(i, 5).Value

        ' Only move if Max Shelf Qty (D) and Break Qty (E) are filled in
        If Not IsEmpty(maxQty) And IsNumeric(maxQty) And _
           Not IsEmpty(breakQty) And IsNumeric(breakQty) Then

            ' Copy to Master_Stock_List
            wsMaster.Cells(masterWriteRow, 1).Value = wsNew.Cells(i, 1).Value  ' Supplier
            wsMaster.Cells(masterWriteRow, 2).Value = wsNew.Cells(i, 2).Value  ' Code
            wsMaster.Cells(masterWriteRow, 3).Value = wsNew.Cells(i, 3).Value  ' Description
            wsMaster.Cells(masterWriteRow, 4).Value = CDbl(maxQty)             ' Max Shelf Qty
            wsMaster.Cells(masterWriteRow, 5).Value = CDbl(breakQty)           ' Break Qty
            wsMaster.Cells(masterWriteRow, 6).Value = wsNew.Cells(i, 6).Value  ' Location

            masterWriteRow = masterWriteRow + 1

            ' Delete the row from New_Items
            wsNew.Rows(i).Delete
            movedCount = movedCount + 1
        Else
            skippedCount = skippedCount + 1
        End If
    Next i

    ' Resize Master table if it exists
    Dim lo As ListObject
    On Error Resume Next
    Set lo = wsMaster.ListObjects(1)
    On Error GoTo 0
    If Not lo Is Nothing Then
        Dim newMasterLast As Long
        newMasterLast = wsMaster.Cells(wsMaster.Rows.Count, 2).End(xlUp).Row
        If newMasterLast >= 2 Then
            lo.Resize wsMaster.Range("A1:F" & newMasterLast)
        End If
    End If

    Application.ScreenUpdating = True

    Dim msg As String
    msg = movedCount & " item(s) moved to Master Stock List."
    If skippedCount > 0 Then
        msg = msg & vbCrLf & skippedCount & " item(s) skipped (still need Max Shelf Qty and/or Break Qty)."
    End If

    MsgBox msg, vbInformation, "Move to Master"

End Sub

'================================================================
' 4. CHECK NEGATIVE STOCK — Flag items with negative qty
'    for physical floor check
'    Shortcut: Ctrl+Shift+G
'================================================================
Public Sub CheckNegativeStock()

    ' Get sheets
    Dim wsStock As Worksheet, wsNeg As Worksheet
    On Error Resume Next
    Set wsStock = ThisWorkbook.Sheets("Daily_Stock_Data")
    On Error GoTo 0

    If wsStock Is Nothing Then
        MsgBox "Daily_Stock_Data sheet not found.", vbExclamation
        Exit Sub
    End If

    ' Get or use existing Negative_Stock_Check sheet
    On Error Resume Next
    Set wsNeg = ThisWorkbook.Sheets("Negative_Stock_Check")
    On Error GoTo 0

    If wsNeg Is Nothing Then
        ' Create the sheet
        Set wsNeg = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsNeg.Name = "Negative_Stock_Check"

        ' Add headers
        wsNeg.Range("A1").Value = "Item"
        wsNeg.Range("B1").Value = "Description"
        wsNeg.Range("C1").Value = "Qty_On_Hand"
        wsNeg.Range("D1").Value = "Check Date"
        wsNeg.Range("E1").Value = "Floor Count"
        wsNeg.Range("F1").Value = "Notes"

        With wsNeg.Range("A1:F1")
            .Font.Bold = True
            .Interior.Color = RGB(192, 0, 0)    ' Dark red
            .Font.Color = RGB(255, 255, 255)     ' White
        End With
    End If

    Application.ScreenUpdating = False
    Application.StatusBar = "Checking for negative stock..."

    ' Clear old data (keep headers)
    Dim negLastRow As Long
    negLastRow = wsNeg.Cells(wsNeg.Rows.Count, 1).End(xlUp).Row
    If negLastRow >= 2 Then
        wsNeg.Range("A2:F" & negLastRow).Clear
    End If

    ' Read stock data
    Dim stockLastRow As Long
    stockLastRow = wsStock.Cells(wsStock.Rows.Count, 1).End(xlUp).Row

    If stockLastRow < 2 Then
        Application.ScreenUpdating = True
        Application.StatusBar = False
        MsgBox "No data in Daily_Stock_Data.", vbInformation
        Exit Sub
    End If

    Dim stockData As Variant
    stockData = wsStock.Range("A2:D" & stockLastRow).Value

    ' Find negative quantities
    Dim negCount As Long
    negCount = 0
    Dim i As Long

    ' Size output array
    Dim outData() As Variant
    ReDim outData(1 To UBound(stockData, 1), 1 To 4)

    For i = 1 To UBound(stockData, 1)
        If Not IsEmpty(stockData(i, 3)) And IsNumeric(stockData(i, 3)) Then
            If CDbl(stockData(i, 3)) < 0 Then
                negCount = negCount + 1
                outData(negCount, 1) = stockData(i, 1)  ' Item
                outData(negCount, 2) = stockData(i, 2)  ' Description
                outData(negCount, 3) = stockData(i, 3)  ' Qty
                outData(negCount, 4) = Date               ' Check date
            End If
        End If
    Next i

    ' Write results
    If negCount > 0 Then
        ' Trim and write
        Dim finalData() As Variant
        ReDim finalData(1 To negCount, 1 To 4)
        Dim j As Long
        For i = 1 To negCount
            For j = 1 To 4
                finalData(i, j) = outData(i, j)
            Next j
        Next i

        wsNeg.Range("A2:D" & (negCount + 1)).Value = finalData
        wsNeg.Range("D2:D" & (negCount + 1)).NumberFormat = "DD/MM/YYYY"

        ' Highlight the qty column red
        With wsNeg.Range("C2:C" & (negCount + 1))
            .Font.Color = RGB(192, 0, 0)
            .Font.Bold = True
        End With

        ' Leave E (Floor Count) and F (Notes) blank for the user
        wsNeg.Range("E2:F" & (negCount + 1)).Interior.Color = RGB(255, 255, 200)

        wsNeg.UsedRange.Columns.AutoFit
    End If

    Application.ScreenUpdating = True
    Application.StatusBar = False

    If negCount = 0 Then
        MsgBox "No negative stock found. All quantities are 0 or above.", _
               vbInformation, "Negative Stock Check"
    Else
        MsgBox negCount & " item(s) with NEGATIVE stock found!" & vbCrLf & vbCrLf & _
               "These need a physical floor check." & vbCrLf & _
               "Use column E to record your floor count." & vbCrLf & _
               "Then update QuickBooks and re-run stock refresh.", _
               vbExclamation, "Negative Stock Check"
        wsNeg.Activate

        ' Also offer to export as PDF for floor walk
        Dim resp As VbMsgBoxResult
        resp = MsgBox("Export negative stock list as PDF for the floor walk?", _
                       vbYesNo + vbQuestion, "Export Checklist?")
        If resp = vbYes Then
            Dim pdfPath As String
            pdfPath = EXPORT_FOLDER & "Negative_Stock_Check_" & Format(Date, "YYYY-MM-DD") & ".pdf"

            ' Check export folder exists
            If Dir(EXPORT_FOLDER, vbDirectory) = "" Then
                MkDir EXPORT_FOLDER
            End If

            wsNeg.ExportAsFixedFormat Type:=xlTypePDF, _
                Filename:=pdfPath, _
                Quality:=xlQualityStandard, _
                OpenAfterPublish:=False

            MsgBox "PDF saved:" & vbCrLf & pdfPath, vbInformation, "Checklist Exported"
        End If
    End If

End Sub

'================================================================
' 5. RUN FULL CYCLE — One button to do everything
'    Refresh Sales + Stock → Check Negative → Detect New Items
'    Shortcut: Ctrl+Shift+A
'================================================================
Public Sub RunFullCycle()

    Dim resp As VbMsgBoxResult
    resp = MsgBox("This will run the full PO cycle:" & vbCrLf & vbCrLf & _
                  "1. Refresh Stock Data" & vbCrLf & _
                  "2. Check Negative Stock" & vbCrLf & _
                  "3. Refresh Sales Data" & vbCrLf & _
                  "4. Detect New Items" & vbCrLf & vbCrLf & _
                  "After reviewing, run Export PO (Ctrl+Shift+E)." & vbCrLf & vbCrLf & _
                  "Continue?", _
                  vbYesNo + vbQuestion, "Run Full PO Cycle")

    If resp = vbNo Then Exit Sub

    ' Step 1: Refresh Stock
    Application.StatusBar = "Step 1/4: Refreshing stock data..."
    RefreshStockData

    ' Step 2: Check Negative Stock
    Application.StatusBar = "Step 2/4: Checking negative stock..."
    CheckNegativeStock

    ' Step 3: Refresh Sales
    Application.StatusBar = "Step 3/4: Refreshing sales data..."
    RefreshSalesData

    ' Step 4: Detect New Items
    Application.StatusBar = "Step 4/4: Detecting new items..."
    DetectNewItems

    Application.StatusBar = False

    MsgBox "Full cycle complete!" & vbCrLf & vbCrLf & _
           "Review the Saas_PO sheet, then press Ctrl+Shift+E to export.", _
           vbInformation, "Cycle Complete"

End Sub

'================================================================
' HELPER: Clean filename (remove invalid characters)
'================================================================
Private Function CleanFileName(ByVal s As String) As String
    Dim result As String
    Dim i As Long
    Dim c As String

    result = ""
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        Select Case c
            Case "\", "/", ":", "*", "?", """", "<", ">", "|"
                result = result & "_"
            Case Else
                result = result & c
        End Select
    Next i

    CleanFileName = Trim(result)
End Function

'================================================================
' HELPER: Get the Nth Monday of a given month/year
'================================================================
Private Function NthMonday(ByVal yr As Long, ByVal mo As Long, ByVal n As Long) As Date
    Dim firstDay As Date
    firstDay = DateSerial(yr, mo, 1)

    ' Find first Monday
    Dim firstMonday As Date
    firstMonday = firstDay + ((9 - Weekday(firstDay, vbSunday)) Mod 7)

    NthMonday = firstMonday + (n - 1) * 7
End Function

'================================================================
' HELPER: Detect which fortnight we're in
' Returns "First" or "Second"
'================================================================
Public Function GetCurrentFortnight() As String
    Dim today As Date
    today = Date

    Dim thirdMonday As Date
    thirdMonday = NthMonday(Year(today), Month(today), 3)

    If today < thirdMonday Then
        GetCurrentFortnight = "First"
    Else
        GetCurrentFortnight = "Second"
    End If
End Function

'================================================================
' HELPER: Calculate fortnightly date range
' For First fortnight: 3rd Monday last month → last day last month
' For Second fortnight: 1st Monday this month → 3rd Monday this month
'================================================================
Public Sub GetFortnightDates(ByRef startDt As Date, ByRef endDt As Date)
    Dim today As Date
    today = Date
    Dim fortnight As String
    fortnight = GetCurrentFortnight()

    If fortnight = "First" Then
        ' 3rd Monday of last month to end of last month
        Dim lastMo As Long, lastYr As Long
        If Month(today) = 1 Then
            lastMo = 12
            lastYr = Year(today) - 1
        Else
            lastMo = Month(today) - 1
            lastYr = Year(today)
        End If
        startDt = NthMonday(lastYr, lastMo, 3)
        endDt = DateSerial(Year(today), Month(today), 0)  ' Last day of prev month
    Else
        ' 1st Monday of current month to 3rd Monday of current month
        startDt = NthMonday(Year(today), Month(today), 1)
        endDt = NthMonday(Year(today), Month(today), 3)
    End If
End Sub
