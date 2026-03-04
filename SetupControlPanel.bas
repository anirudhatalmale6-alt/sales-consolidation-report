Attribute VB_Name = "SetupControlPanel"
'================================================================
' SETUP CONTROL PANEL
' ------------------------------------------------
' Run this ONCE to create the Control Panel sheet
' with clickable buttons, supplier dropdown, and
' "Today's Orders" auto-detection.
'
' After running, you can delete this module — the buttons
' and formulas stay on the sheet permanently.
'================================================================

Option Explicit

Public Sub CreateControlPanel()

    Application.ScreenUpdating = False

    ' Delete existing Control_Panel sheet if it exists
    Dim wsOld As Worksheet
    On Error Resume Next
    Set wsOld = ThisWorkbook.Sheets("Control_Panel")
    On Error GoTo 0
    If Not wsOld Is Nothing Then
        Application.DisplayAlerts = False
        wsOld.Delete
        Application.DisplayAlerts = True
    End If

    ' Create the sheet at the beginning
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
    ws.Name = "Control_Panel"

    ' Remove gridlines
    ActiveWindow.DisplayGridlines = False

    ' Set column widths
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 22
    ws.Columns("C").ColumnWidth = 22
    ws.Columns("D").ColumnWidth = 22
    ws.Columns("E").ColumnWidth = 22
    ws.Columns("F").ColumnWidth = 3

    ' === TITLE ===
    With ws.Range("B1:E1")
        .Merge
        .Value = "PURCHASE ORDER CONTROL PANEL"
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 51, 102)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 45
    End With

    ' === SUBTITLE ===
    With ws.Range("B2:E2")
        .Merge
        .Value = "Select a supplier, then click buttons to run each step"
        .Font.Size = 10
        .Font.Color = RGB(100, 100, 100)
        .HorizontalAlignment = xlCenter
        .RowHeight = 25
    End With

    ' ============================================================
    ' === SECTION: TODAY'S ORDERS ===
    ' ============================================================
    Dim row As Long
    row = 4
    Call AddSectionHeader(ws, row, "B", "E", "TODAY'S ORDERS", RGB(204, 102, 0))

    row = row + 1
    ws.Rows(row).RowHeight = 22
    ws.Cells(row, 2).Value = "Today:"
    ws.Cells(row, 2).Font.Bold = True
    ws.Cells(row, 2).Font.Size = 11

    ' Formula that checks which suppliers have their order day today
    ' Compares today's weekday/date against Order_Cycle_Calc Order_Day
    ' and cross-references with Supplier_Details
    With ws.Cells(row, 3)
        .Formula = "=TEXT(TODAY(),""dddd"") & "" "" & TEXT(TODAY(),""DD/MM/YYYY"")"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = RGB(204, 102, 0)
    End With

    row = row + 1
    ws.Rows(row).RowHeight = 30

    With ws.Range("B" & row & ":E" & row)
        .Merge
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = RGB(0, 128, 0)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    ' This formula builds a list of suppliers whose order day matches today
    ' It checks each supplier's cycle against Order_Cycle_Calc to find today's day
    ws.Cells(row, 2).Formula = _
        "=LET(suppliers,Supplier_Details!A2:A100," & _
        "cycles,Supplier_Details!B2:B100," & _
        "days,Order_Cycle_Calc!B2:B20," & _
        "cycleNames,Order_Cycle_Calc!A2:A20," & _
        "today,TEXT(TODAY(),""dddd"")," & _
        "todayDOM,DAY(TODAY())," & _
        "result,TEXTJOIN("", "", TRUE," & _
        "IF((XLOOKUP(cycles,cycleNames,days,"""")=today)" & _
        "+((XLOOKUP(cycles,cycleNames,days,"""")=""1st Monday"")*((todayDOM<=7)*(today=""Monday"")))" & _
        "+((XLOOKUP(cycles,cycleNames,days,"""")=""3rd Monday"")*((todayDOM>=15)*(todayDOM<=21)*(today=""Monday"")))," & _
        "suppliers,""""))," & _
        "IF(result="""",""No orders due today"",""Orders due: "" & result))"

    ' ============================================================
    ' === SECTION: SELECT SUPPLIER ===
    ' ============================================================
    row = row + 2
    Call AddSectionHeader(ws, row, "B", "E", "SELECT SUPPLIER", RGB(0, 51, 102))

    row = row + 1
    ws.Rows(row).RowHeight = 35

    ws.Cells(row, 2).Value = "Supplier:"
    ws.Cells(row, 2).Font.Bold = True
    ws.Cells(row, 2).Font.Size = 12
    ws.Cells(row, 2).VerticalAlignment = xlCenter

    ' Supplier dropdown cell — shows current value from Date_Selector!A2
    ' (No formula — data validation dropdown needs a value cell, not a formula cell)
    Dim dropCell As Range
    Set dropCell = ws.Cells(row, 3)

    ' Read current supplier from Date_Selector
    Dim currentSupplier As String
    currentSupplier = ""
    On Error Resume Next
    currentSupplier = Trim(CStr(ThisWorkbook.Sheets("Date_Selector").Range("A2").Value))
    On Error GoTo 0

    With dropCell
        .Value = currentSupplier
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = RGB(0, 51, 102)
        .Interior.Color = RGB(230, 240, 250)
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
    End With

    ' Merge C and D for the dropdown
    ws.Range("C" & row & ":D" & row).Merge

    ' Add data validation dropdown (list from Supplier_Details)
    With dropCell.Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="=Supplier_Details!$A$2:$A$100"
        .IgnoreBlank = True
        .ShowInput = True
        .ShowError = True
        .InputTitle = "Select Supplier"
        .InputMessage = "Choose from the list, then click SET SUPPLIER"
    End With

    ' "Set Supplier" button in column E
    Dim setBtnLeft As Double, setBtnTop As Double
    setBtnLeft = ws.Range("E" & row).Left + 2
    setBtnTop = ws.Range("E" & row).Top + 3

    Dim setBtn As Object
    Set setBtn = ws.Buttons.Add(setBtnLeft, setBtnTop, ws.Range("E" & row).Width - 4, 28)
    With setBtn
        .Caption = "SET SUPPLIER"
        .OnAction = "POWorkflow.SyncSupplier"
        .Font.Size = 9
        .Font.Bold = True
    End With

    ' Info labels
    Dim infoRow As Long
    infoRow = row + 1
    ws.Rows(infoRow).RowHeight = 20
    ws.Cells(infoRow, 2).Value = "Cycle:"
    ws.Cells(infoRow, 2).Font.Bold = True
    With ws.Cells(infoRow, 3)
        .Formula = "=Date_Selector!B2"
        .Font.Size = 10
    End With

    infoRow = infoRow + 1
    ws.Rows(infoRow).RowHeight = 20
    ws.Cells(infoRow, 2).Value = "Date Range:"
    ws.Cells(infoRow, 2).Font.Bold = True
    With ws.Cells(infoRow, 3)
        .Formula = "=TEXT(Date_Selector!C2,""DD/MM/YYYY"") & "" to "" & TEXT(Date_Selector!D2,""DD/MM/YYYY"")"
        .Font.Size = 10
    End With

    infoRow = infoRow + 1
    ws.Rows(infoRow).RowHeight = 20
    ws.Cells(infoRow, 2).Value = "Optional Cycle:"
    ws.Cells(infoRow, 2).Font.Bold = True
    With ws.Cells(infoRow, 3)
        .Formula = "=IF(Date_Selector!E2="""",""None"",Date_Selector!E2 & "": "" & TEXT(Date_Selector!F2,""DD/MM/YYYY"") & "" to "" & TEXT(Date_Selector!G2,""DD/MM/YYYY""))"
        .Font.Size = 10
        .Font.Color = RGB(128, 0, 128)
    End With

    ' ============================================================
    ' === SECTION: WORKFLOW BUTTONS ===
    ' ============================================================
    row = infoRow + 2
    Call AddSectionHeader(ws, row, "B", "E", "STEP 1: REFRESH DATA", RGB(0, 102, 153))

    row = row + 1
    Call AddButton(ws, row, "B", "Refresh Stock Data", "RefreshStockData.RefreshStockData", _
                   RGB(0, 128, 180), "Reads QB Product/Service List export")

    Call AddButton(ws, row, "D", "Refresh Sales Data", "RefreshSalesData.RefreshSalesData", _
                   RGB(0, 128, 180), "Reads QB Sales Report export")

    row = row + 3
    Call AddSectionHeader(ws, row, "B", "E", "STEP 2: QUALITY CHECKS", RGB(192, 0, 0))

    row = row + 1
    Call AddButton(ws, row, "B", "Check Negative Stock", "POWorkflow.CheckNegativeStock", _
                   RGB(200, 50, 50), "Flags items needing a floor count")

    Call AddButton(ws, row, "D", "Detect New Items", "POWorkflow.DetectNewItems", _
                   RGB(200, 50, 50), "Finds items not in Master Stock List")

    row = row + 3
    Call AddSectionHeader(ws, row, "B", "E", "STEP 3: MASTER STOCK", RGB(0, 128, 0))

    row = row + 1
    Call AddButton(ws, row, "B", "Move to Master", "POWorkflow.MoveToMaster", _
                   RGB(50, 150, 50), "Transfers completed New_Items to Master")

    row = row + 3
    Call AddSectionHeader(ws, row, "B", "E", "STEP 4: EXPORT", RGB(128, 0, 128))

    row = row + 1
    Call AddButton(ws, row, "B", "Export PO", "POWorkflow.ExportPO", _
                   RGB(150, 50, 150), "Saves .xlsx + .pdf, merges ad-hoc items")

    ' ============================================================
    ' === BIG BUTTON: FULL CYCLE ===
    ' ============================================================
    row = row + 3
    Call AddSectionHeader(ws, row, "B", "E", "OR: ONE-CLICK FULL CYCLE", RGB(0, 51, 102))

    row = row + 1
    ws.Rows(row).RowHeight = 50

    Dim bigBtnLeft As Double, bigBtnTop As Double
    bigBtnLeft = ws.Range("B" & row).Left + 5
    bigBtnTop = ws.Range("B" & row).Top + 5

    Dim bigBtn As Object
    Set bigBtn = ws.Buttons.Add(bigBtnLeft, bigBtnTop, _
                                 ws.Range("B" & row).Width + ws.Range("C" & row).Width + _
                                 ws.Range("D" & row).Width + ws.Range("E" & row).Width - 10, 40)
    With bigBtn
        .Caption = "RUN FULL PO CYCLE"
        .OnAction = "POWorkflow.RunFullCycle"
        .Font.Size = 14
        .Font.Bold = True
    End With

    row = row + 1
    ws.Rows(row).RowHeight = 18
    With ws.Range("B" & row & ":E" & row)
        .Merge
        .Value = "Runs all steps in order: Stock > Neg Check > Sales > New Items. Then review Saas_PO and click Export."
        .Font.Size = 9
        .Font.Italic = True
        .Font.Color = RGB(100, 100, 100)
        .HorizontalAlignment = xlCenter
    End With

    ' ============================================================
    ' === PROTECT SHEET (but allow dropdown) ===
    ' ============================================================
    ' Unlock the supplier dropdown cell so it can be changed
    ws.Cells.Locked = True
    dropCell.Locked = False

    ws.Protect Password:="", UserInterfaceOnly:=True, _
        AllowFormattingCells:=False, AllowFormattingColumns:=False, _
        AllowFormattingRows:=False

    ws.Activate
    ws.Range("A1").Select

    Application.ScreenUpdating = True

    MsgBox "Control Panel created!" & vbCrLf & vbCrLf & _
           "Features:" & vbCrLf & _
           "  - Today's Orders: shows which suppliers are due today" & vbCrLf & _
           "  - Supplier Dropdown: select supplier right here" & vbCrLf & _
           "  - Step-by-step buttons: idiot-proof workflow" & vbCrLf & _
           "  - Full Cycle button: one-click does everything" & vbCrLf & vbCrLf & _
           "You can delete the SetupControlPanel module now.", _
           vbInformation, "Setup Complete"

End Sub

'================================================================
' HELPER: Add a section header row
'================================================================
Private Sub AddSectionHeader(ws As Worksheet, row As Long, colStart As String, colEnd As String, _
                              title As String, bgColor As Long)
    With ws.Range(colStart & row & ":" & colEnd & row)
        .Merge
        .Value = title
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = bgColor
        .HorizontalAlignment = xlLeft
        .RowHeight = 26
        .IndentLevel = 1
    End With
End Sub

'================================================================
' HELPER: Add a styled button with description
'================================================================
Private Sub AddButton(ws As Worksheet, row As Long, col As String, _
                       caption As String, macroName As String, _
                       btnColor As Long, description As String)

    Dim btnLeft As Double, btnTop As Double
    Dim btnWidth As Double, btnHeight As Double

    ' Use 2 columns wide for each button
    Dim colNum As Long
    colNum = ws.Range(col & "1").Column

    btnLeft = ws.Cells(row, colNum).Left + 5
    btnTop = ws.Cells(row, colNum).Top + 3
    btnWidth = ws.Cells(row, colNum).Width + ws.Cells(row, colNum + 1).Width - 10
    btnHeight = 34

    ws.Rows(row).RowHeight = 40

    Dim btn As Object
    Set btn = ws.Buttons.Add(btnLeft, btnTop, btnWidth, btnHeight)

    With btn
        .Caption = caption
        .OnAction = macroName
        .Font.Size = 11
        .Font.Bold = True
    End With

    ' Description text in the next row
    ws.Rows(row + 1).RowHeight = 16
    Dim descCell As Range
    Set descCell = ws.Cells(row + 1, colNum)
    descCell.Value = description
    descCell.Font.Size = 8
    descCell.Font.Italic = True
    descCell.Font.Color = RGB(100, 100, 100)

End Sub

'================================================================
' SUPPLIER CHANGE HANDLER
' ------------------------------------------------
' When the supplier dropdown on Control_Panel is changed,
' this updates Date_Selector!A2 to match.
' Must be placed in the ThisWorkbook module's Worksheet_Change
' event — see instructions below.
'================================================================
' INSTRUCTIONS: Copy this code into ThisWorkbook (not a module):
'
' Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
'     If Sh.Name = "Control_Panel" Then
'         If Not Intersect(Target, Sh.Range("C11")) Is Nothing Then
'             ' Update Date_Selector when supplier is changed on Control Panel
'             Application.EnableEvents = False
'             ThisWorkbook.Sheets("Date_Selector").Range("A2").Value = Target.Value
'             Application.EnableEvents = True
'         End If
'     End If
' End Sub
