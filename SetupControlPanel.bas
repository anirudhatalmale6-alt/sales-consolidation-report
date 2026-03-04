Attribute VB_Name = "SetupControlPanel"
'================================================================
' SETUP CONTROL PANEL
' ------------------------------------------------
' Run this ONCE to create the Control Panel sheet
' with clickable buttons for all PO workflow functions.
'
' After running, you can delete this module — the buttons
' stay on the sheet permanently.
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
    ws.Columns("B").ColumnWidth = 40
    ws.Columns("C").ColumnWidth = 5
    ws.Columns("D").ColumnWidth = 40
    ws.Columns("E").ColumnWidth = 3

    ' === TITLE ===
    With ws.Range("B1:D1")
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
    With ws.Range("B2:D2")
        .Merge
        .Value = "Click a button to run the corresponding function"
        .Font.Size = 10
        .Font.Color = RGB(100, 100, 100)
        .HorizontalAlignment = xlCenter
        .RowHeight = 25
    End With

    ' === SECTION: DATA REFRESH ===
    Dim row As Long
    row = 4
    Call AddSectionHeader(ws, row, "B", "D", "DATA REFRESH", RGB(0, 102, 153))

    row = row + 1
    Call AddButton(ws, row, "B", "Refresh Stock Data", "RefreshStockData.RefreshStockData", _
                   RGB(0, 128, 180), "Reads QB Product/Service List, updates Daily_Stock_Data")

    row = row + 2
    Call AddButton(ws, row, "D", "Refresh Sales Data", "RefreshSalesData.RefreshSalesData", _
                   RGB(0, 128, 180), "Reads QB Sales Report, updates Sales_Data")

    ' === SECTION: CHECKS ===
    row = row + 3
    Call AddSectionHeader(ws, row, "B", "D", "QUALITY CHECKS", RGB(192, 0, 0))

    row = row + 1
    Call AddButton(ws, row, "B", "Check Negative Stock", "POWorkflow.CheckNegativeStock", _
                   RGB(200, 50, 50), "Flags negative qty items for physical floor check")

    row = row + 2
    Call AddButton(ws, row, "D", "Detect New Items", "POWorkflow.DetectNewItems", _
                   RGB(200, 50, 50), "Finds items not yet in Master Stock List")

    ' === SECTION: MASTER STOCK ===
    row = row + 3
    Call AddSectionHeader(ws, row, "B", "D", "MASTER STOCK LIST", RGB(0, 128, 0))

    row = row + 1
    Call AddButton(ws, row, "B", "Move New Items to Master", "POWorkflow.MoveToMaster", _
                   RGB(50, 150, 50), "Transfers completed New_Items into Master Stock List")

    ' === SECTION: EXPORT ===
    row = row + 3
    Call AddSectionHeader(ws, row, "B", "D", "EXPORT", RGB(128, 0, 128))

    row = row + 1
    Call AddButton(ws, row, "B", "Export PO", "POWorkflow.ExportPO", _
                   RGB(150, 50, 150), "Saves Saas_PO as .xlsx + .pdf (merges ad-hoc items)")

    ' === SECTION: FULL CYCLE ===
    row = row + 3
    Call AddSectionHeader(ws, row, "B", "D", "FULL CYCLE (ALL-IN-ONE)", RGB(0, 51, 102))

    row = row + 1
    Call AddButton(ws, row, "B", "Run Full PO Cycle", "POWorkflow.RunFullCycle", _
                   RGB(0, 51, 102), "Stock Refresh > Neg Check > Sales Refresh > New Items")

    row = row + 2
    With ws.Range("B" & row & ":D" & row)
        .Merge
        .Value = "Full Cycle runs steps 1-4 in order. Review Saas_PO, then click Export PO."
        .Font.Size = 9
        .Font.Italic = True
        .Font.Color = RGB(100, 100, 100)
        .HorizontalAlignment = xlCenter
    End With

    ' === CURRENT SETTINGS DISPLAY ===
    row = row + 3
    Call AddSectionHeader(ws, row, "B", "D", "CURRENT SETTINGS", RGB(80, 80, 80))

    row = row + 1
    ws.Cells(row, 2).Value = "Supplier:"
    ws.Cells(row, 2).Font.Bold = True
    ws.Cells(row, 3).Value = ""
    With ws.Cells(row, 4)
        .Formula = "=Date_Selector!A2"
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = RGB(0, 51, 102)
    End With

    row = row + 1
    ws.Cells(row, 2).Value = "Cycle:"
    ws.Cells(row, 2).Font.Bold = True
    With ws.Cells(row, 4)
        .Formula = "=Date_Selector!B2"
        .Font.Size = 11
    End With

    row = row + 1
    ws.Cells(row, 2).Value = "Date Range:"
    ws.Cells(row, 2).Font.Bold = True
    With ws.Cells(row, 4)
        .Formula = "=TEXT(Date_Selector!C2,""DD/MM/YYYY"") & "" to "" & TEXT(Date_Selector!D2,""DD/MM/YYYY"")"
        .Font.Size = 11
    End With

    ' Protect sheet (allow button clicks but prevent accidental edits)
    ws.Protect Password:="", UserInterfaceOnly:=True, _
        AllowFormattingCells:=False, AllowFormattingColumns:=False, _
        AllowFormattingRows:=False

    ws.Activate
    ws.Range("A1").Select

    Application.ScreenUpdating = True

    MsgBox "Control Panel created!" & vbCrLf & vbCrLf & _
           "The Control_Panel sheet is now your main dashboard." & vbCrLf & _
           "Click any button to run the corresponding function." & vbCrLf & vbCrLf & _
           "You can delete the SetupControlPanel module now — " & vbCrLf & _
           "the buttons are permanent.", _
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
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = bgColor
        .HorizontalAlignment = xlLeft
        .RowHeight = 28
        .IndentLevel = 1
    End With
End Sub

'================================================================
' HELPER: Add a styled button with description
'================================================================
Private Sub AddButton(ws As Worksheet, row As Long, col As String, _
                       caption As String, macroName As String, _
                       btnColor As Long, description As String)

    ' Button dimensions
    Dim btnLeft As Double, btnTop As Double
    Dim btnWidth As Double, btnHeight As Double

    btnLeft = ws.Range(col & row).Left + 5
    btnTop = ws.Range(col & row).Top + 3
    btnWidth = ws.Range(col & row).Width - 10
    btnHeight = 36

    ' Set row height
    ws.Rows(row).RowHeight = 42

    ' Create button
    Dim btn As Object
    Set btn = ws.Buttons.Add(btnLeft, btnTop, btnWidth, btnHeight)

    With btn
        .Caption = caption
        .OnAction = macroName
        .Font.Size = 12
        .Font.Bold = True
    End With

    ' Description text in the next row
    ws.Rows(row + 1).RowHeight = 18
    Dim descCell As Range
    Set descCell = ws.Cells(row + 1, ws.Range(col & "1").Column)
    descCell.Value = description
    descCell.Font.Size = 9
    descCell.Font.Italic = True
    descCell.Font.Color = RGB(100, 100, 100)

End Sub
