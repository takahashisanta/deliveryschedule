Attribute VB_Name = "DriverScheduleMod"
'================================================================================
' PSD Driver Schedule – VBA Module
' Sheet: DRIVER_SCHEDULE  (output / printable)
' Sheet: DRIVER_ENGINE    (hidden calculation table)
' Sheet: DRIVER_RULES     (configuration tables)
' Source: MASTER DATA     (existing sheet in this workbook)
'
' HOW TO IMPORT:
'   1. Open PSD_Driver_Schedule.xlsx in Excel
'   2. Alt + F11  →  VBA Editor
'   3. File → Import File → select DriverSchedule.bas
'   4. Close VBA Editor
'   5. Add a button on DRIVER_SCHEDULE and assign macro "GenerateSchedule"
'================================================================================
Option Explicit

'================================================================================
' CONFIGURATION – adjust these if column positions in MASTER DATA change
'================================================================================
Private Const MD_SHEET          As String = "MASTER DATA"

' MASTER DATA columns (1-based)
Private Const MD_COL_REQ_ID     As Long = 1   ' A  – Request ID
Private Const MD_COL_SHIP_DATE  As Long = 2   ' B  – Ship Date
Private Const MD_COL_SOLD_TO    As Long = 3   ' C  – Sold To (customer)
Private Const MD_COL_SHIP_TO    As Long = 4   ' D  – Ship To
Private Const MD_COL_AREA       As Long = 5   ' E  – Area
Private Const MD_COL_SHIP_MTH   As Long = 6   ' F  – Shipping Method
Private Const MD_COL_PRODUCT    As Long = 7   ' G  – Product
Private Const MD_COL_QTY        As Long = 8   ' H  – Quantity / Pallets
Private Const MD_COL_WEIGHT     As Long = 9   ' I  – Weight
Private Const MD_COL_STORAGE    As Long = 10  ' J  – Storage Location
Private Const MD_COL_WIN_START  As Long = 11  ' K  – Window Start
Private Const MD_COL_WIN_END    As Long = 12  ' L  – Window End
Private Const MD_COL_VERSION    As Long = 20  ' T  – Version / Status

' Shipping method values to INCLUDE (case-insensitive match)
Private Const SHIP_INCLUDE      As String = "Delivery"

' Vehicle capacity (pallets)
Private Const TRUCK_CAPACITY    As Long = 8
Private Const VAN_CAPACITY      As Long = 3

' Pallet threshold for Flexible customers (>= this → Truck)
Private Const FLEX_TRUCK_PALLETS As Long = 4

' DRIVER_ENGINE columns (1-based)
Private Const ENG_COL_REQ_ID    As Long = 1
Private Const ENG_COL_SHIP_DATE As Long = 2
Private Const ENG_COL_SOLD_TO   As Long = 3
Private Const ENG_COL_SHIP_TO   As Long = 4
Private Const ENG_COL_AREA      As Long = 5
Private Const ENG_COL_PRODUCT   As Long = 6
Private Const ENG_COL_QTY       As Long = 7
Private Const ENG_COL_WEIGHT    As Long = 8
Private Const ENG_COL_STORAGE   As Long = 9
Private Const ENG_COL_WIN_START As Long = 10
Private Const ENG_COL_WIN_END   As Long = 11
Private Const ENG_COL_WIN_NOTE  As Long = 12
Private Const ENG_COL_VEH_RAW   As Long = 13
Private Const ENG_COL_VEH_ASGN  As Long = 14
Private Const ENG_COL_TIME_PRI  As Long = 15
Private Const ENG_COL_AREA_RANK As Long = 16
Private Const ENG_COL_SORT_KEY  As Long = 17

'================================================================================
' MAIN ENTRY POINT
'================================================================================
Public Sub GenerateSchedule()
    Application.ScreenUpdating = False
    Application.Calculation   = xlCalculationManual

    On Error GoTo ErrHandler

    Dim wsSchedule  As Worksheet
    Dim wsEngine    As Worksheet
    Dim wsRules     As Worksheet
    Dim wsMaster    As Worksheet

    Set wsSchedule = ThisWorkbook.Sheets("DRIVER_SCHEDULE")
    Set wsEngine   = ThisWorkbook.Sheets("DRIVER_ENGINE")
    Set wsRules    = ThisWorkbook.Sheets("DRIVER_RULES")

    ' ── Validate / locate MASTER DATA ──────────────────────────────────────────
    On Error Resume Next
    Set wsMaster = ThisWorkbook.Sheets(MD_SHEET)
    On Error GoTo ErrHandler
    If wsMaster Is Nothing Then
        MsgBox "Sheet '" & MD_SHEET & "' not found in this workbook." & vbLf & _
               "Please ensure MASTER DATA is in the same workbook.", vbCritical, "Missing Sheet"
        GoTo CleanUp
    End If

    ' ── Read selected date ─────────────────────────────────────────────────────
    Dim selectedDate As Date
    If IsEmpty(wsSchedule.Range("B2").Value) Or Not IsDate(wsSchedule.Range("B2").Value) Then
        MsgBox "Please enter a valid delivery date in cell B2.", vbExclamation, "Date Required"
        GoTo CleanUp
    End If
    selectedDate = CDate(wsSchedule.Range("B2").Value)

    ' ── Get valid areas for this day of week ───────────────────────────────────
    Dim dayAreas() As String
    dayAreas = GetAreasForDay(wsRules, selectedDate)
    If UBound(dayAreas) < 0 Then
        MsgBox "No area schedule defined for " & Format(selectedDate, "dddd") & "." & vbLf & _
               "Check TABLE 1 in DRIVER_RULES.", vbExclamation, "No Areas"
        GoTo CleanUp
    End If

    ' ── Clear the engine sheet (keep header rows 1-2) ──────────────────────────
    Dim engLastRow As Long
    engLastRow = wsEngine.Cells(wsEngine.Rows.Count, 1).End(xlUp).Row
    If engLastRow > 2 Then
        wsEngine.Rows("3:" & engLastRow).ClearContents
    End If

    ' ── Pull matching rows from MASTER DATA into DRIVER_ENGINE ─────────────────
    Dim mdLastRow As Long
    mdLastRow = wsMaster.Cells(wsMaster.Rows.Count, MD_COL_REQ_ID).End(xlUp).Row

    Dim engineRow As Long
    engineRow = 3

    ' First pass: collect latest versions per Request ID
    Dim reqVersions As Object
    Set reqVersions = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 2 To mdLastRow
        Dim reqID   As String
        Dim verVal  As Variant
        reqID  = Trim(CStr(wsMaster.Cells(i, MD_COL_REQ_ID).Value))
        verVal = wsMaster.Cells(i, MD_COL_VERSION).Value

        Dim verNum As Double
        If IsNumeric(verVal) Then
            verNum = CDbl(verVal)
        Else
            verNum = 0
        End If

        If reqVersions.Exists(reqID) Then
            If verNum > reqVersions(reqID) Then reqVersions(reqID) = verNum
        Else
            reqVersions(reqID) = verNum
        End If
    Next i

    ' Second pass: filter and populate engine
    For i = 2 To mdLastRow
        ' Check ship date
        Dim shipDateVal As Variant
        shipDateVal = wsMaster.Cells(i, MD_COL_SHIP_DATE).Value
        If IsEmpty(shipDateVal) Then GoTo NextRow
        If Not IsDate(shipDateVal) Then GoTo NextRow
        If CDate(shipDateVal) <> selectedDate Then GoTo NextRow

        ' Check shipping method = "Delivery"
        Dim shipMethod As String
        shipMethod = Trim(wsMaster.Cells(i, MD_COL_SHIP_MTH).Value)
        If InStr(1, shipMethod, SHIP_INCLUDE, vbTextCompare) = 0 Then GoTo NextRow

        ' Check area matches day schedule
        Dim areaVal As String
        areaVal = Trim(wsMaster.Cells(i, MD_COL_AREA).Value)
        If Not AreaInList(areaVal, dayAreas) Then GoTo NextRow

        ' Check this is the latest version (not superseded)
        reqID  = Trim(CStr(wsMaster.Cells(i, MD_COL_REQ_ID).Value))
        verVal = wsMaster.Cells(i, MD_COL_VERSION).Value
        If IsNumeric(verVal) Then
            verNum = CDbl(verVal)
        Else
            verNum = 0
        End If
        If reqVersions.Exists(reqID) Then
            If verNum < reqVersions(reqID) Then GoTo NextRow  ' superseded
        End If

        ' ── Write to engine ────────────────────────────────────────────────────
        Dim soldTo As String
        soldTo = Trim(wsMaster.Cells(i, MD_COL_SOLD_TO).Value)

        Dim qtyVal As Variant
        qtyVal = wsMaster.Cells(i, MD_COL_QTY).Value
        Dim pallets As Double
        pallets = IIf(IsNumeric(qtyVal), CDbl(qtyVal), 0)

        Dim winStart As Variant
        Dim winEnd   As Variant
        winStart = wsMaster.Cells(i, MD_COL_WIN_START).Value
        winEnd   = wsMaster.Cells(i, MD_COL_WIN_END).Value

        Dim timePriority As Integer
        timePriority = GetTimePriority(winEnd)

        Dim vehicleRaw  As String
        Dim vehicleAsgn As String
        vehicleRaw  = GetVehiclePreferenceFromRules(wsRules, soldTo)
        vehicleAsgn = AssignVehicle(vehicleRaw, pallets)

        Dim areaRank As Integer
        areaRank = GetAreaRank(areaVal)

        Dim sortKey As String
        sortKey = IIf(vehicleAsgn = "Truck", "1", "2") & _
                  Format(timePriority, "0") & _
                  Format(areaRank, "00") & _
                  Left(soldTo & "          ", 10)

        ' Window display string
        Dim winNote As String
        winNote = FormatWindow(winStart, winEnd)

        wsEngine.Cells(engineRow, ENG_COL_REQ_ID).Value    = wsMaster.Cells(i, MD_COL_REQ_ID).Value
        wsEngine.Cells(engineRow, ENG_COL_SHIP_DATE).Value = wsMaster.Cells(i, MD_COL_SHIP_DATE).Value
        wsEngine.Cells(engineRow, ENG_COL_SOLD_TO).Value   = soldTo
        wsEngine.Cells(engineRow, ENG_COL_SHIP_TO).Value   = wsMaster.Cells(i, MD_COL_SHIP_TO).Value
        wsEngine.Cells(engineRow, ENG_COL_AREA).Value      = areaVal
        wsEngine.Cells(engineRow, ENG_COL_PRODUCT).Value   = wsMaster.Cells(i, MD_COL_PRODUCT).Value
        wsEngine.Cells(engineRow, ENG_COL_QTY).Value       = qtyVal
        wsEngine.Cells(engineRow, ENG_COL_WEIGHT).Value    = wsMaster.Cells(i, MD_COL_WEIGHT).Value
        wsEngine.Cells(engineRow, ENG_COL_STORAGE).Value   = wsMaster.Cells(i, MD_COL_STORAGE).Value
        wsEngine.Cells(engineRow, ENG_COL_WIN_START).Value = winStart
        wsEngine.Cells(engineRow, ENG_COL_WIN_END).Value   = winEnd
        wsEngine.Cells(engineRow, ENG_COL_WIN_NOTE).Value  = winNote
        wsEngine.Cells(engineRow, ENG_COL_VEH_RAW).Value   = vehicleRaw
        wsEngine.Cells(engineRow, ENG_COL_VEH_ASGN).Value  = vehicleAsgn
        wsEngine.Cells(engineRow, ENG_COL_TIME_PRI).Value  = timePriority
        wsEngine.Cells(engineRow, ENG_COL_AREA_RANK).Value = areaRank
        wsEngine.Cells(engineRow, ENG_COL_SORT_KEY).Value  = sortKey

        engineRow = engineRow + 1
NextRow:
    Next i

    Dim totalEngineRows As Long
    totalEngineRows = engineRow - 3   ' rows written (excluding header rows 1-2)

    If totalEngineRows = 0 Then
        MsgBox "No matching deliveries found for " & Format(selectedDate, "MM/DD/YYYY") & "." & vbLf & _
               "Check date, area schedule, and shipping method.", vbInformation, "No Data"
        GoTo CleanUp
    End If

    ' ── Sort engine data by SortKey ────────────────────────────────────────────
    SortEngineData wsEngine, engineRow - 1

    ' ── Write formatted output to DRIVER_SCHEDULE ─────────────────────────────
    WriteScheduleOutput wsSchedule, wsEngine, engineRow - 1, selectedDate

    MsgBox "Schedule generated: " & totalEngineRows & " deliveries for " & _
           Format(selectedDate, "dddd, MM/DD/YYYY"), vbInformation, "Done"

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation   = xlCalculationAutomatic
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.Calculation   = xlCalculationAutomatic
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error"
End Sub


'================================================================================
' GetAreasForDay  –  reads TABLE 1 in DRIVER_RULES
' Returns string array of area names for the day of selectedDate.
' Day names in col A must match exactly (Mon, Tue, Wed, Thu, Fri).
'================================================================================
Private Function GetAreasForDay(wsRules As Worksheet, selectedDate As Date) As String()
    Dim dayName As String
    Select Case Weekday(selectedDate, vbMonday)   ' 1=Mon … 5=Fri
        Case 1: dayName = "Mon"
        Case 2: dayName = "Tue"
        Case 3: dayName = "Wed"
        Case 4: dayName = "Thu"
        Case 5: dayName = "Fri"
        Case Else
            GetAreasForDay = Split("", ",")  ' empty array for weekend
            Exit Function
    End Select

    ' Scan col A for the day name (TABLE 1 starts around row 4)
    Dim r As Long
    For r = 4 To 20
        If Trim(wsRules.Cells(r, 1).Value) = dayName Then
            ' Collect non-empty values from columns B-G
            Dim areas() As String
            ReDim areas(0 To 5)
            Dim count As Integer
            count = 0
            Dim c As Long
            For c = 2 To 7
                Dim areaName As String
                areaName = Trim(wsRules.Cells(r, c).Value)
                If areaName <> "" Then
                    areas(count) = areaName
                    count = count + 1
                End If
            Next c
            ReDim Preserve areas(0 To count - 1)
            GetAreasForDay = areas
            Exit Function
        End If
    Next r

    ' Not found – return empty
    GetAreasForDay = Split("", ",")
End Function


'================================================================================
' AreaInList  –  case-insensitive check if areaVal is in dayAreas()
'================================================================================
Private Function AreaInList(areaVal As String, dayAreas() As String) As Boolean
    Dim a As Variant
    For Each a In dayAreas
        If StrComp(Trim(areaVal), Trim(CStr(a)), vbTextCompare) = 0 Then
            AreaInList = True
            Exit Function
        End If
    Next a
    AreaInList = False
End Function


'================================================================================
' GetVehiclePreferenceFromRules  –  reads TABLE 2 in DRIVER_RULES
' Returns "Truck", "Van", or "Flexible"
'================================================================================
Private Function GetVehiclePreferenceFromRules(wsRules As Worksheet, _
                                               soldTo As String) As String
    ' TABLE 2 header is around row 13; data starts row 14.
    ' Find the header row to be safe.
    Dim headerRow As Long
    headerRow = 0
    Dim r As Long
    For r = 10 To 50
        If InStr(1, wsRules.Cells(r, 1).Value, "Customer Name", vbTextCompare) > 0 Then
            headerRow = r
            Exit For
        End If
    Next r
    If headerRow = 0 Then
        GetVehiclePreferenceFromRules = "Flexible"
        Exit Function
    End If

    ' Scan customer names (col A) for a partial match
    For r = headerRow + 1 To headerRow + 50
        Dim custName As String
        custName = Trim(wsRules.Cells(r, 1).Value)
        If custName = "" Then Exit For

        ' Match: soldTo contains custName OR custName contains soldTo
        If InStr(1, soldTo, custName, vbTextCompare) > 0 Or _
           InStr(1, custName, soldTo, vbTextCompare) > 0 Then
            Dim pref As String
            pref = Trim(wsRules.Cells(r, 2).Value)
            If pref <> "" Then
                GetVehiclePreferenceFromRules = pref
                Exit Function
            End If
        End If
    Next r

    ' Special rule: "Mitsuwa" in name → Van
    If InStr(1, soldTo, "Mitsuwa", vbTextCompare) > 0 Then
        GetVehiclePreferenceFromRules = "Van"
        Exit Function
    End If

    GetVehiclePreferenceFromRules = "Flexible"
End Function


'================================================================================
' AssignVehicle  –  converts raw preference + pallets into Truck / Van
'================================================================================
Private Function AssignVehicle(vehicleRaw As String, pallets As Double) As String
    Select Case LCase(Trim(vehicleRaw))
        Case "truck"
            AssignVehicle = "Truck"
        Case "van"
            AssignVehicle = "Van"
        Case Else   ' Flexible
            AssignVehicle = IIf(pallets >= FLEX_TRUCK_PALLETS, "Truck", "Van")
    End Select
End Function


'================================================================================
' GetTimePriority  –  maps delivery window end time to priority 1-9
'================================================================================
Private Function GetTimePriority(winEnd As Variant) As Integer
    If IsEmpty(winEnd) Or Not IsDate(winEnd) Then
        GetTimePriority = 9
        Exit Function
    End If

    ' Extract time portion (works whether winEnd is a time or datetime)
    Dim t As Date
    t = TimeValue(CDate(winEnd))

    Dim h As Integer
    h = Hour(t)
    Dim m As Integer
    m = Minute(t)

    ' By 8:00 AM
    If h < 8 Or (h = 8 And m = 0) Then
        GetTimePriority = 1
    ' 8:01 – 12:00
    ElseIf h < 12 Or (h = 12 And m = 0) Then
        GetTimePriority = 2
    ' 12:01 – 15:00
    ElseIf h < 15 Or (h = 15 And m = 0) Then
        GetTimePriority = 3
    Else
        GetTimePriority = 9
    End If
End Function


'================================================================================
' GetAreaRank  –  numeric rank for area sorting within a vehicle group
'================================================================================
Private Function GetAreaRank(areaVal As String) As Integer
    ' Define preferred delivery order within a vehicle
    Dim areaOrder As Variant
    areaOrder = Array("JFC/IDC", "EAST LA", "CENTRAL LA", "SAN GABRIEL", _
                      "SOUTH BAY", "SOUTH LA", "WEST LA", "LAX", "NAX", _
                      "NORTH OC")
    Dim i As Integer
    For i = 0 To UBound(areaOrder)
        If StrComp(Trim(areaVal), CStr(areaOrder(i)), vbTextCompare) = 0 Then
            GetAreaRank = i + 1
            Exit Function
        End If
    Next i
    GetAreaRank = 99  ' unknown area → sort last
End Function


'================================================================================
' FormatWindow  –  builds readable time window string "HH:MM – HH:MM"
'================================================================================
Private Function FormatWindow(winStart As Variant, winEnd As Variant) As String
    If IsEmpty(winStart) And IsEmpty(winEnd) Then
        FormatWindow = "Open"
        Exit Function
    End If
    Dim s As String
    Dim e As String
    s = IIf(IsDate(winStart), Format(TimeValue(CDate(winStart)), "HH:MM"), "")
    e = IIf(IsDate(winEnd),   Format(TimeValue(CDate(winEnd)),   "HH:MM"), "")
    If s <> "" And e <> "" Then
        FormatWindow = s & " – " & e
    ElseIf e <> "" Then
        FormatWindow = "By " & e
    ElseIf s <> "" Then
        FormatWindow = "After " & s
    Else
        FormatWindow = "Open"
    End If
End Function


'================================================================================
' SortEngineData  –  sorts rows 3:lastRow by Sort Key (col Q)
'================================================================================
Private Sub SortEngineData(wsEngine As Worksheet, lastRow As Long)
    If lastRow < 4 Then Exit Sub  ' nothing to sort (row 3 is first data row)

    With wsEngine.Sort
        .SortFields.Clear
        .SortFields.Add Key:=wsEngine.Range( _
            wsEngine.Cells(3, ENG_COL_SORT_KEY), _
            wsEngine.Cells(lastRow, ENG_COL_SORT_KEY)), _
            SortOn:=xlSortOnValues, Order:=xlAscending
        .SetRange wsEngine.Range( _
            wsEngine.Cells(3, 1), _
            wsEngine.Cells(lastRow, ENG_COL_SORT_KEY))
        .Header  = xlNo
        .Apply
    End With
End Sub


'================================================================================
' WriteScheduleOutput  –  writes Truck + Van sections to DRIVER_SCHEDULE
'================================================================================
Private Sub WriteScheduleOutput(wsSchedule As Worksheet, wsEngine As Worksheet, _
                                 lastEngineRow As Long, selectedDate As Date)

    ' ── Wipe previous output (rows 6 onwards) ──────────────────────────────────
    wsSchedule.Rows("6:200").ClearContents
    wsSchedule.Rows("6:200").Interior.ColorIndex = xlNone
    wsSchedule.Rows("6:200").Borders.LineStyle   = xlNone

    ' ── Read all engine rows into arrays for speed ────────────────────────────
    Dim truckRows() As Long
    Dim vanRows()   As Long
    ReDim truckRows(0 To lastEngineRow)
    ReDim vanRows(0 To lastEngineRow)
    Dim tCount As Long : tCount = 0
    Dim vCount As Long : vCount = 0

    Dim r As Long
    For r = 3 To lastEngineRow
        Dim veh As String
        veh = Trim(wsEngine.Cells(r, ENG_COL_VEH_ASGN).Value)
        If veh = "Truck" Then
            truckRows(tCount) = r
            tCount = tCount + 1
        Else
            vanRows(vCount) = r
            vCount = vCount + 1
        End If
    Next r

    ' ── Write TRUCK section ────────────────────────────────────────────────────
    Dim writeRow As Long
    writeRow = 4    ' rows 4-5 are section header + column headers (pre-set by template)

    ' Update section header with date
    wsSchedule.Range("A4").Value = "🚚  TRUCK ROUTE  –  " & Format(selectedDate, "ddd MM/DD/YYYY")

    writeRow = 6   ' data starts row 6
    Dim stopNum As Long : stopNum = 1

    Dim t As Long
    For t = 0 To tCount - 1
        Dim eRow As Long : eRow = truckRows(t)
        WriteOutputRow wsSchedule, wsEngine, writeRow, eRow, stopNum, "Truck"
        stopNum  = stopNum + 1
        writeRow = writeRow + 1
    Next t

    If tCount = 0 Then
        wsSchedule.Cells(writeRow, 1).Value = "(No truck deliveries)"
        wsSchedule.Cells(writeRow, 1).Font.Italic = True
        wsSchedule.Cells(writeRow, 1).Font.Color  = RGB(127, 127, 127)
        wsSchedule.Merge_cells writeRow   ' minor style only
        writeRow = writeRow + 1
    End If

    ' ── Gap row + VAN header ───────────────────────────────────────────────────
    writeRow = writeRow + 1   ' blank gap

    Dim vanHeaderRow As Long : vanHeaderRow = writeRow
    wsSchedule.Range("A" & vanHeaderRow & ":H" & vanHeaderRow).Merge
    With wsSchedule.Cells(vanHeaderRow, 1)
        .Value = "🚐  VAN ROUTE  –  " & Format(selectedDate, "ddd MM/DD/YYYY")
        .Font  = Nothing
        .Font.Name  = "Calibri"
        .Font.Bold  = True
        .Font.Size  = 13
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(55, 86, 35)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment   = xlCenter
        .IndentLevel = 1
    End With
    wsSchedule.Rows(vanHeaderRow).RowHeight = 24
    writeRow = writeRow + 1

    ' Van column headers
    Dim vanHdrCols As Variant
    vanHdrCols = Array("Stop #", "Customer", "Ship To / Area", "Product", _
                       "Qty/Pallets", "Window", "Storage", "Notes")
    Dim colIdx As Long
    For colIdx = 0 To 7
        With wsSchedule.Cells(writeRow, colIdx + 1)
            .Value = vanHdrCols(colIdx)
            .Font.Name  = "Calibri"
            .Font.Bold  = True
            .Font.Size  = 10
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(83, 129, 53)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment   = xlCenter
            .Borders.LineStyle   = xlContinuous
            .Borders.Weight      = xlThin
        End With
    Next colIdx
    wsSchedule.Rows(writeRow).RowHeight = 20
    writeRow = writeRow + 1

    ' Van data rows
    stopNum = 1
    Dim v As Long
    For v = 0 To vCount - 1
        eRow = vanRows(v)
        WriteOutputRow wsSchedule, wsEngine, writeRow, eRow, stopNum, "Van"
        stopNum  = stopNum + 1
        writeRow = writeRow + 1
    Next v

    If vCount = 0 Then
        wsSchedule.Cells(writeRow, 1).Value = "(No van deliveries)"
        wsSchedule.Cells(writeRow, 1).Font.Italic = True
        wsSchedule.Cells(writeRow, 1).Font.Color  = RGB(127, 127, 127)
        writeRow = writeRow + 1
    End If

    ' ── Footer ─────────────────────────────────────────────────────────────────
    writeRow = writeRow + 1
    wsSchedule.Cells(writeRow, 1).Value = "Generated: " & Now()
    wsSchedule.Cells(writeRow, 1).Font.Italic = True
    wsSchedule.Cells(writeRow, 1).Font.Size   = 9
    wsSchedule.Cells(writeRow, 1).Font.Color  = RGB(127, 127, 127)
    wsSchedule.Merge_cells writeRow

    ' ── Update print area ──────────────────────────────────────────────────────
    wsSchedule.PageSetup.PrintArea = "A1:H" & writeRow

End Sub


'================================================================================
' WriteOutputRow  –  writes one delivery row to the output sheet
'================================================================================
Private Sub WriteOutputRow(wsSchedule As Worksheet, wsEngine As Worksheet, _
                            outRow As Long, engRow As Long, _
                            stopNum As Long, vehicle As String)

    Dim isTruck As Boolean : isTruck = (vehicle = "Truck")
    Dim bgColor As Long
    bgColor = IIf(outRow Mod 2 = 0, _
                  IIf(isTruck, RGB(214, 228, 240), RGB(226, 239, 218)), _
                  RGB(255, 255, 255))

    Dim timePri As Integer
    timePri = CInt(wsEngine.Cells(engRow, ENG_COL_TIME_PRI).Value)

    ' Urgency flag
    Dim urgFlag As String
    urgFlag = IIf(timePri = 1, "⚡ URGENT – ", "")

    Dim soldTo  As String : soldTo  = wsEngine.Cells(engRow, ENG_COL_SOLD_TO).Value
    Dim shipTo  As String : shipTo  = wsEngine.Cells(engRow, ENG_COL_SHIP_TO).Value
    Dim area    As String : area    = wsEngine.Cells(engRow, ENG_COL_AREA).Value
    Dim product As String : product = wsEngine.Cells(engRow, ENG_COL_PRODUCT).Value
    Dim qty     As Variant : qty    = wsEngine.Cells(engRow, ENG_COL_QTY).Value
    Dim storage As String : storage = wsEngine.Cells(engRow, ENG_COL_STORAGE).Value
    Dim winNote As String : winNote = wsEngine.Cells(engRow, ENG_COL_WIN_NOTE).Value

    Dim shipToArea As String
    shipToArea = IIf(shipTo <> "", shipTo & " (" & area & ")", area)

    Dim vals As Variant
    vals = Array(stopNum, urgFlag & soldTo, shipToArea, product, qty, winNote, storage, "")

    Dim c As Long
    For c = 0 To 7
        With wsSchedule.Cells(outRow, c + 1)
            .Value = vals(c)
            .Font.Name  = "Calibri"
            .Font.Size  = 10
            .Font.Bold  = (c = 0)
            .Interior.Color        = bgColor
            .HorizontalAlignment   = IIf(c = 0 Or c = 4, xlCenter, xlLeft)
            .VerticalAlignment     = xlCenter
            .Borders.LineStyle     = xlContinuous
            .Borders.Weight        = xlThin
        End With
    Next c

    ' Highlight urgent rows (time priority 1) in orange
    If timePri = 1 Then
        For c = 1 To 8
            wsSchedule.Cells(outRow, c).Interior.Color = RGB(255, 199, 100)
        Next c
    End If

    wsSchedule.Rows(outRow).RowHeight = 18
End Sub


'================================================================================
' Merge_cells helper  –  merges A:H for a given row (avoids range string concat)
'================================================================================
Private Sub Merge_cells(ws As Worksheet, rowNum As Long)
    ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 8)).Merge
End Sub


'================================================================================
' ClearSchedule  –  optional: button to wipe output without regenerating
'================================================================================
Public Sub ClearSchedule()
    Dim wsSchedule As Worksheet
    Set wsSchedule = ThisWorkbook.Sheets("DRIVER_SCHEDULE")
    wsSchedule.Rows("6:200").ClearContents
    wsSchedule.Rows("6:200").Interior.ColorIndex = xlNone
    wsSchedule.Rows("6:200").Borders.LineStyle   = xlNone
    MsgBox "Schedule cleared.", vbInformation, "Cleared"
End Sub
