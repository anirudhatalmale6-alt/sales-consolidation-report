Attribute VB_Name = "RefreshModule"
'================================================================
' SALES CONSOLIDATION - REFRESH MODULE
' ------------------------------------------------
' RefreshData:         Reads external OneDrive workbook,
'                      filters by date range, groups by Item,
'                      sums Qty, writes to Summary sheet.
'
' TestWithSampleData:  Same logic but reads from the built-in
'                      SampleData sheet (no external file needed).
'
' Shortcut:  Ctrl+Shift+R  (assigned in Auto_Open)
'================================================================

Option Explicit

' ---- Column positions in the SOURCE data (1-based) ----
Private Const COL_SUPPLIER   As Long = 1   ' Column A
Private Const COL_DATE       As Long = 2   ' Column B
Private Const COL_ITEM       As Long = 3   ' Column C
Private Const COL_DESC       As Long = 4   ' Column D
Private Const COL_QTY        As Long = 5   ' Column E

' ---- Output starts at this row on Summary sheet ----
Private Const OUT_START_ROW  As Long = 7

'================================================================
' PUBLIC: Refresh from external workbook
'================================================================
Public Sub RefreshData()

    Dim wsSettings As Worksheet
    Set wsSettings = ThisWorkbook.Sheets("Settings")

    Dim filePath As String
    filePath = Trim(CStr(wsSettings.Range("B5").Value))

    Dim sheetName As String
    sheetName = Trim(CStr(wsSettings.Range("B6").Value))

    Dim startDt As Date, endDt As Date

    ' Validate dates
    If Not IsDate(wsSettings.Range("B10").Value) Then
        MsgBox "Start Date (Settings B10) is not a valid date.", vbExclamation
        Exit Sub
    End If
    If Not IsDate(wsSettings.Range("B11").Value) Then
        MsgBox "End Date (Settings B11) is not a valid date.", vbExclamation
        Exit Sub
    End If

    startDt = CDate(wsSettings.Range("B10").Value)
    endDt = CDate(wsSettings.Range("B11").Value)

    If startDt > endDt Then
        MsgBox "Start Date cannot be after End Date.", vbExclamation
        Exit Sub
    End If

    ' Validate file path
    If Len(filePath) = 0 Then
        MsgBox "Please enter the source file path on the Settings sheet (B5).", vbExclamation
        Exit Sub
    End If

    ' Check if file exists
    If Dir(filePath) = "" Then
        MsgBox "Source file not found:" & vbCrLf & filePath & vbCrLf & vbCrLf & _
               "Please check the path on the Settings sheet (B5)." & vbCrLf & _
               "Tip: Right-click the file in File Explorer > Properties > Location", _
               vbExclamation, "File Not Found"
        Exit Sub
    End If

    ' Open source workbook (read-only)
    Application.ScreenUpdating = False
    Application.StatusBar = "Opening source workbook..."

    Dim wbSource As Workbook
    On Error Resume Next
    Set wbSource = Workbooks.Open(Filename:=filePath, ReadOnly:=True, UpdateLinks:=0)
    On Error GoTo 0

    If wbSource Is Nothing Then
        Application.ScreenUpdating = True
        Application.StatusBar = False
        MsgBox "Could not open the source workbook." & vbCrLf & filePath, vbExclamation
        Exit Sub
    End If

    ' Find the source sheet
    Dim wsSource As Worksheet
    On Error Resume Next
    Set wsSource = wbSource.Sheets(sheetName)
    On Error GoTo 0

    If wsSource Is Nothing Then
        wbSource.Close SaveChanges:=False
        Application.ScreenUpdating = True
        Application.StatusBar = False
        MsgBox "Sheet '" & sheetName & "' not found in the source workbook." & vbCrLf & _
               "Available sheets: " & JoinSheetNames(wbSource), vbExclamation
        Exit Sub
    End If

    ' Read source data into array
    Dim lastRow As Long
    lastRow = wsSource.Cells(wsSource.Rows.Count, COL_ITEM).End(xlUp).Row

    If lastRow < 2 Then
        wbSource.Close SaveChanges:=False
        Application.ScreenUpdating = True
        Application.StatusBar = False
        MsgBox "No data found in the source sheet.", vbInformation
        Exit Sub
    End If

    Dim srcData As Variant
    srcData = wsSource.Range(wsSource.Cells(2, 1), wsSource.Cells(lastRow, COL_QTY)).Value

    ' Close source workbook
    wbSource.Close SaveChanges:=False

    ' Process and write results
    ProcessData srcData, startDt, endDt

    Application.ScreenUpdating = True
    Application.StatusBar = False

End Sub

'================================================================
' PUBLIC: Test with built-in SampleData sheet
'================================================================
Public Sub TestWithSampleData()

    Dim wsSettings As Worksheet
    Set wsSettings = ThisWorkbook.Sheets("Settings")

    Dim startDt As Date, endDt As Date

    If Not IsDate(wsSettings.Range("B10").Value) Then
        MsgBox "Start Date (Settings B10) is not a valid date.", vbExclamation
        Exit Sub
    End If
    If Not IsDate(wsSettings.Range("B11").Value) Then
        MsgBox "End Date (Settings B11) is not a valid date.", vbExclamation
        Exit Sub
    End If

    startDt = CDate(wsSettings.Range("B10").Value)
    endDt = CDate(wsSettings.Range("B11").Value)

    Dim wsSample As Worksheet
    On Error Resume Next
    Set wsSample = ThisWorkbook.Sheets("SampleData")
    On Error GoTo 0

    If wsSample Is Nothing Then
        MsgBox "SampleData sheet not found in this workbook.", vbExclamation
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = wsSample.Cells(wsSample.Rows.Count, COL_ITEM).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "No data found in SampleData sheet.", vbInformation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim srcData As Variant
    srcData = wsSample.Range(wsSample.Cells(2, 1), wsSample.Cells(lastRow, COL_QTY)).Value

    ProcessData srcData, startDt, endDt

    Application.ScreenUpdating = True

End Sub

'================================================================
' PRIVATE: Filter by date, group by Item, sum Qty, write output
'================================================================
Private Sub ProcessData(srcData As Variant, startDt As Date, endDt As Date)

    Dim wsSummary As Worksheet
    Set wsSummary = ThisWorkbook.Sheets("Summary")

    ' Clear old data (from OUT_START_ROW down)
    Dim lastClear As Long
    lastClear = wsSummary.Cells(wsSummary.Rows.Count, 1).End(xlUp).Row
    If lastClear >= OUT_START_ROW Then
        wsSummary.Range(wsSummary.Cells(OUT_START_ROW, 1), _
                        wsSummary.Cells(lastClear, 4)).Clear
    End If

    ' Use dictionaries for grouping
    Dim dictQty As Object, dictSupplier As Object, dictDesc As Object, dictOrder As Object
    Set dictQty = CreateObject("Scripting.Dictionary")
    Set dictSupplier = CreateObject("Scripting.Dictionary")
    Set dictDesc = CreateObject("Scripting.Dictionary")
    Set dictOrder = CreateObject("Scripting.Dictionary")

    Dim i As Long, rowCount As Long, filteredCount As Long
    Dim itemKey As String, rowDate As Date
    Dim orderIdx As Long

    rowCount = UBound(srcData, 1)
    orderIdx = 0
    filteredCount = 0

    Application.StatusBar = "Processing " & rowCount & " rows..."

    For i = 1 To rowCount
        ' Skip rows with empty Item or Date
        If Not IsEmpty(srcData(i, COL_ITEM)) And Not IsEmpty(srcData(i, COL_DATE)) Then
            If IsDate(srcData(i, COL_DATE)) Then
                rowDate = CDate(srcData(i, COL_DATE))

                ' Date filter
                If rowDate >= startDt And rowDate <= endDt Then
                    filteredCount = filteredCount + 1
                    itemKey = CStr(srcData(i, COL_ITEM))

                    If dictQty.Exists(itemKey) Then
                        ' Add to existing quantity
                        dictQty(itemKey) = dictQty(itemKey) + CDbl(srcData(i, COL_QTY))
                    Else
                        ' New item
                        dictQty.Add itemKey, CDbl(srcData(i, COL_QTY))
                        dictSupplier.Add itemKey, CStr(srcData(i, COL_SUPPLIER))
                        dictDesc.Add itemKey, CStr(srcData(i, COL_DESC))
                        orderIdx = orderIdx + 1
                        dictOrder.Add itemKey, orderIdx
                    End If
                End If
            End If
        End If
    Next i

    ' Check if any rows matched
    If dictQty.Count = 0 Then
        ' Show diagnostic info to help troubleshoot
        Dim sampleDate As String
        If rowCount > 0 Then
            If IsDate(srcData(1, COL_DATE)) Then
                sampleDate = Format(CDate(srcData(1, COL_DATE)), "DD/MM/YYYY")
            Else
                sampleDate = CStr(srcData(1, COL_DATE)) & " (not a date)"
            End If
        End If

        wsSummary.Range("B4").Value = "Last refreshed: " & Format(Now, "DD/MM/YYYY HH:MM:SS") & " — No results"
        wsSummary.Range("B4").Font.Italic = True
        wsSummary.Range("B4").Font.Color = RGB(200, 0, 0)

        MsgBox "No rows matched the date range." & vbCrLf & vbCrLf & _
               "Date filter: " & Format(startDt, "DD/MM/YYYY") & " to " & Format(endDt, "DD/MM/YYYY") & vbCrLf & _
               "Source rows scanned: " & rowCount & vbCrLf & _
               "First row date value: " & sampleDate & vbCrLf & vbCrLf & _
               "Check that your date range overlaps with the source data.", _
               vbExclamation, "No Results"

        wsSummary.Activate
        Application.StatusBar = False
        Exit Sub
    End If

    ' Write results to Summary sheet
    Dim keys As Variant
    keys = dictQty.keys

    Dim outRow As Long
    Dim thinBorder As Variant

    ' Sort keys by original order
    Dim sortedKeys() As String
    ReDim sortedKeys(0 To dictQty.Count - 1)
    Dim k As Variant
    For Each k In keys
        sortedKeys(dictOrder(k) - 1) = CStr(k)
    Next k

    ' Formatting
    Dim altFill As Long

    For i = 0 To UBound(sortedKeys)
        outRow = OUT_START_ROW + i
        itemKey = sortedKeys(i)

        wsSummary.Cells(outRow, 1).Value = dictSupplier(itemKey)
        wsSummary.Cells(outRow, 2).Value = itemKey
        wsSummary.Cells(outRow, 3).Value = dictDesc(itemKey)
        wsSummary.Cells(outRow, 4).Value = dictQty(itemKey)
        wsSummary.Cells(outRow, 4).NumberFormat = "#,##0"

        ' Borders
        Dim c As Long
        For c = 1 To 4
            With wsSummary.Cells(outRow, c).Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = RGB(200, 200, 200)
            End With
        Next c

        ' Alternating row color
        If i Mod 2 = 1 Then
            wsSummary.Range(wsSummary.Cells(outRow, 1), wsSummary.Cells(outRow, 4)).Interior.Color = RGB(232, 245, 233)
        Else
            wsSummary.Range(wsSummary.Cells(outRow, 1), wsSummary.Cells(outRow, 4)).Interior.ColorIndex = xlNone
        End If
    Next i

    ' Update timestamp
    wsSummary.Range("B4").Value = "Last refreshed: " & Format(Now, "DD/MM/YYYY HH:MM:SS")
    wsSummary.Range("B4").Font.Italic = True
    wsSummary.Range("B4").Font.Color = RGB(0, 100, 0)

    ' Summary message
    Dim uniqueCount As Long
    uniqueCount = dictQty.Count

    MsgBox "Refresh complete!" & vbCrLf & vbCrLf & _
           "Source rows scanned: " & rowCount & vbCrLf & _
           "Rows matching date range: " & filteredCount & vbCrLf & _
           "Unique products (SKUs): " & uniqueCount, _
           vbInformation, "Sales Consolidation"

    ' Activate Summary sheet
    wsSummary.Activate
    wsSummary.Range("A" & OUT_START_ROW).Select

    Application.StatusBar = False

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
' AUTO: Assign keyboard shortcut on workbook open
'================================================================
Public Sub Auto_Open()
    Application.OnKey "+^r", "RefreshData"  ' Ctrl+Shift+R
End Sub

Public Sub Auto_Close()
    Application.OnKey "+^r"  ' Reset shortcut
End Sub
