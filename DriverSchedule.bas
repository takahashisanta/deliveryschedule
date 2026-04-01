Attribute VB_Name = "DriverScheduleMod"
'================================================================================
' PSD Driver Schedule – VBA Module
' Reads MASTER DATA → writes to "Driver Schedule Sheet"
'
' HOW TO IMPORT:
'   1. Open PSD Order Process Sheet.xlsm in Excel
'   2. Alt + F11  →  VBA Editor
'   3. File → Import File → select DriverSchedule.bas
'   4. Close VBA Editor
'   5. On "Driver Schedule Sheet", draw a Form Control Button and
'      assign macro: DriverScheduleMod.GenerateSchedule
'================================================================================
Option Explicit

'================================================================================
' ── CONFIGURATION ──────────────────────────────────────────────────────────────
' Adjust constants here if your sheet/column positions ever change.
'================================================================================

' Sheet names
Private Const MD_SHEET          As String = "MASTER DATA"
Private Const SCHED_SHEET       As String = "Driver Schedule Sheet"
Private Const ENGINE_SHEET      As String = "DRIVER_ENGINE"
Private Const RULES_SHEET       As String = "DRIVER_RULES"

' ── MASTER DATA column numbers (1 = col A) ────────────────────────────────────
Private Const MD_COL_REQ_ID     As Long = 1   ' A  – Request ID
Private Const MD_COL_SHIP_DATE  As Long = 2   ' B  – Date (Ship Date)
' C (3)  = Category   – not used
' D (4)  = Rep        – not used
Private Const MD_COL_SHIP_MTH   As Long = 5   ' E  – Shipping Method
' F (6)  = Delivery Window label – not used (we use G/H for actual times)
Private Const MD_COL_WIN_START  As Long = 7   ' G  – Start (window start time)
Private Const MD_COL_WIN_END    As Long = 8   ' H  – Finish (window end time)
' I (9)  = PO #       – not used
Private Const MD_COL_SOLD_TO    As Long = 10  ' J  – Sold To Customer  (used for vehicle preference lookup)
' K (11) = Sold To Address – not used for schedule
Private Const MD_COL_SHIP_TO    As Long = 12  ' L  – Ship To Customer  (shown as "Delivery to" on schedule)
Private Const MD_COL_ADDRESS    As Long = 13  ' M  – Ship To Address   (shown as dropoff address)
Private Const MD_COL_AREA       As Long = 14  ' N  – Area
Private Const MD_COL_PRODUCT    As Long = 15  ' O  – Product / Item Description
Private Const MD_COL_QTY        As Long = 16  ' P  – Quantity / Cases
' Q (17) = Weight     – not used
' R (18) = Price/Unit – not used
Private Const MD_COL_STORAGE    As Long = 19  ' S  – Storage Location
Private Const MD_COL_VERSION    As Long = 20  ' T  – Version (highest number = current)
' U (21) = Color Sort – not used
Private Const MD_COL_PALLETS    As Long = 22  ' V  – Pallets  ← ADD col V to MASTER DATA
'
' ⚠  PALLETS (col V = 22):
'    MASTER DATA currently ends at col U. Add a "Pallets" header at col V and
'    enter the pallet count per order row. The macro sums pallets per stop
'    to decide Truck vs Van for flexible customers.
'    Rule: total pallets for a stop >= FLEX_TRUCK_PALLETS (4) → Truck, else Van.

' ── Driver Schedule Sheet – output row anchors ────────────────────────────────
Private Const TRUCK_DATA_START  As Long = 5   ' first data row, section 1 (Truck)
Private Const TRUCK_DATA_CLEAR  As Long = 68  ' clear rows 5 → 68 before writing
Private Const VAN_DATA_START    As Long = 74  ' first data row, section 2 (Van)
Private Const VAN_DATA_CLEAR    As Long = 150 ' clear rows 74 → 150 before writing

' ── Driver Schedule Sheet – output columns (1-based) ──────────────────────────
'   A  B             C           D       E        F           G(empty)  H
'   #  Delivery/PU   qualifier   time1   time2    Customer              Address
'   I              J            K        L              M
'   Item Desc      Qty/Cases    Pallets  Storage        Info for Driver (blank)
Private Const OUT_A_STOP        As Long = 1   ' Stop number
Private Const OUT_B_TYPE        As Long = 2   ' "Delivery" / "Pick-up"
Private Const OUT_C_QUAL        As Long = 3   ' "By" / "Between" / "Anytime"
Private Const OUT_D_TIME1       As Long = 4   ' time value (By or Between start)
Private Const OUT_E_TIME2       As Long = 5   ' time value (Between end only)
Private Const OUT_F_CUST        As Long = 6   ' Customer (Sold To)
' col G (7) = always empty
Private Const OUT_H_ADDR        As Long = 8   ' Full delivery address
Private Const OUT_I_ITEM        As Long = 9   ' Item Description
Private Const OUT_J_QTY         As Long = 10  ' Quantity / Cases
Private Const OUT_K_PALLETS     As Long = 11  ' Pallets
Private Const OUT_L_STORAGE     As Long = 12  ' Item Location / Storage
' col M (13) = Info for Driver – left blank for manual entry

' ── Vehicle capacity & assignment ─────────────────────────────────────────────
Private Const TRUCK_CAPACITY    As Long = 8   ' max pallets for Truck
Private Const VAN_CAPACITY      As Long = 3   ' max pallets for Van
Private Const FLEX_TRUCK_PALLETS As Long = 4  ' flexible customer: >= this → Truck

' ── Row colours ───────────────────────────────────────────────────────────────
Private Const CLR_PICKUP        As Long = 15624315  ' salmon/pink  (RGB 237,187,187)
Private Const CLR_WHITE         As Long = 16777215  ' white


'================================================================================
' MAIN ENTRY POINT  –  called by the Generate Schedule button
'================================================================================
Public Sub GenerateSchedule()
    Application.ScreenUpdating = False
    Application.Calculation   = xlCalculationManual

    On Error GoTo ErrHandler

    ' ── Locate sheets ──────────────────────────────────────────────────────────
    Dim wsSched   As Worksheet
    Dim wsEngine  As Worksheet
    Dim wsRules   As Worksheet
    Dim wsMaster  As Worksheet

    Set wsSched = GetSheet(SCHED_SHEET)
    If wsSched Is Nothing Then
        MsgBox "Sheet '" & SCHED_SHEET & "' not found.", vbCritical: GoTo CleanUp
    End If

    Set wsMaster = GetSheet(MD_SHEET)
    If wsMaster Is Nothing Then
        MsgBox "Sheet '" & MD_SHEET & "' not found.", vbCritical: GoTo CleanUp
    End If

    Set wsEngine = GetSheet(ENGINE_SHEET)   ' optional debug sheet – OK if absent
    Set wsRules  = GetSheet(RULES_SHEET)

    ' ── Read selected date ─────────────────────────────────────────────────────
    ' Date input is in cell B1 of Driver Schedule Sheet
    If IsEmpty(wsSched.Range("B1").Value) Or Not IsDate(wsSched.Range("B1").Value) Then
        MsgBox "Please enter a valid delivery date in cell B1 of '" & SCHED_SHEET & "'.", _
               vbExclamation: GoTo CleanUp
    End If
    Dim selDate As Date
    selDate = CDate(wsSched.Range("B1").Value)

    ' ── Get valid areas for the day of week ────────────────────────────────────
    Dim dayAreas() As String
    If Not wsRules Is Nothing Then
        dayAreas = GetAreasForDay(wsRules, selDate)
    Else
        ' Fallback: no DRIVER_RULES sheet – accept all areas
        ReDim dayAreas(0): dayAreas(0) = "*"
    End If

    If UBound(dayAreas) < 0 Then
        MsgBox "No area schedule defined for " & Format(selDate, "dddd") & "." & vbLf & _
               "Check TABLE 1 in DRIVER_RULES.", vbExclamation: GoTo CleanUp
    End If

    ' ── Collect qualifying rows from MASTER DATA ───────────────────────────────
    Dim rows() As DelivRow
    Dim rowCount As Long
    rowCount = 0
    ReDim rows(0 To 500)

    ' Build latest-version lookup
    Dim latestVer As Object
    Set latestVer = CreateObject("Scripting.Dictionary")

    Dim mdLast As Long
    mdLast = wsMaster.Cells(wsMaster.Rows.Count, MD_COL_REQ_ID).End(xlUp).Row

    Dim i As Long
    For i = 2 To mdLast
        Dim rid As String: rid = Trim(CStr(wsMaster.Cells(i, MD_COL_REQ_ID).Value))
        If rid = "" Then GoTo NextVer
        Dim ver As Double
        ver = IIf(IsNumeric(wsMaster.Cells(i, MD_COL_VERSION).Value), _
                  CDbl(wsMaster.Cells(i, MD_COL_VERSION).Value), 0)
        If Not latestVer.Exists(rid) Then
            latestVer(rid) = ver
        ElseIf ver > latestVer(rid) Then
            latestVer(rid) = ver
        End If
NextVer:
    Next i

    ' Filter rows
    For i = 2 To mdLast
        ' Ship date
        Dim sdv As Variant: sdv = wsMaster.Cells(i, MD_COL_SHIP_DATE).Value
        If IsEmpty(sdv) Or Not IsDate(sdv) Then GoTo SkipRow
        If CDate(sdv) <> selDate Then GoTo SkipRow

        ' Shipping method: must be "Delivery" or "Pick-up"
        Dim mth As String
        mth = Trim(wsMaster.Cells(i, MD_COL_SHIP_MTH).Value)
        Dim isDelivery As Boolean: isDelivery = (InStr(1, mth, "Delivery", vbTextCompare) > 0)
        Dim isPickup   As Boolean: isPickup   = (InStr(1, mth, "Pick-up",  vbTextCompare) > 0 Or _
                                                  InStr(1, mth, "Pickup",   vbTextCompare) > 0)
        If Not isDelivery And Not isPickup Then GoTo SkipRow

        ' Area matches day schedule
        Dim areaVal As String: areaVal = Trim(wsMaster.Cells(i, MD_COL_AREA).Value)
        If Not AreaAllowed(areaVal, dayAreas) Then GoTo SkipRow

        ' Not superseded
        rid = Trim(CStr(wsMaster.Cells(i, MD_COL_REQ_ID).Value))
        ver = IIf(IsNumeric(wsMaster.Cells(i, MD_COL_VERSION).Value), _
                  CDbl(wsMaster.Cells(i, MD_COL_VERSION).Value), 0)
        If latestVer.Exists(rid) Then
            If ver < latestVer(rid) Then GoTo SkipRow
        End If

        ' ── Store row ─────────────────────────────────────────────────────────
        If rowCount > UBound(rows) Then ReDim Preserve rows(0 To UBound(rows) + 200)

        With rows(rowCount)
            .ReqID     = rid
            .SoldTo    = Trim(wsMaster.Cells(i, MD_COL_SOLD_TO).Value)
            .ShipTo    = Trim(wsMaster.Cells(i, MD_COL_SHIP_TO).Value)
            .Address   = Trim(wsMaster.Cells(i, MD_COL_ADDRESS).Value)
            .Area      = areaVal
            .StopType  = IIf(isPickup, "Pick-up", "Delivery")
            .Product   = Trim(wsMaster.Cells(i, MD_COL_PRODUCT).Value)
            .QtyCase   = wsMaster.Cells(i, MD_COL_QTY).Value
            .Pallets   = IIf(IsNumeric(wsMaster.Cells(i, MD_COL_PALLETS).Value), _
                              CDbl(wsMaster.Cells(i, MD_COL_PALLETS).Value), 0)
            .Storage   = Trim(wsMaster.Cells(i, MD_COL_STORAGE).Value)
            .WinStart  = wsMaster.Cells(i, MD_COL_WIN_START).Value
            .WinEnd    = wsMaster.Cells(i, MD_COL_WIN_END).Value
            ' StopKey groups same customer + same address as one stop
            .StopKey   = LCase(.SoldTo) & "||" & LCase(.Address)
        End With
        rowCount = rowCount + 1
SkipRow:
    Next i

    If rowCount = 0 Then
        MsgBox "No matching deliveries found for " & Format(selDate, "MM/DD/YYYY") & "." & vbLf & _
               "Check date, area schedule, and shipping method.", vbInformation: GoTo CleanUp
    End If

    ReDim Preserve rows(0 To rowCount - 1)

    ' ── Assign vehicle to each stop (sum pallets per StopKey) ─────────────────
    AssignVehicles rows, rowCount, wsRules

    ' ── Calculate sort key for each row ───────────────────────────────────────
    For i = 0 To rowCount - 1
        rows(i).TimePri  = GetTimePriority(rows(i).WinEnd)
        rows(i).AreaRank = GetAreaRank(rows(i).Area)
        rows(i).SortKey  = IIf(rows(i).Vehicle = "Truck", "1", "2") & _
                           Format(rows(i).TimePri, "0") & _
                           Format(rows(i).AreaRank, "00") & _
                           Left(rows(i).SoldTo & String(20, " "), 20) & _
                           Left(rows(i).Address & String(30, " "), 30)
    Next i

    ' ── Sort rows by SortKey ───────────────────────────────────────────────────
    SortRows rows, rowCount

    ' ── Optionally populate DRIVER_ENGINE for audit/debug ─────────────────────
    If Not wsEngine Is Nothing Then
        PopulateEngine wsEngine, rows, rowCount
    End If

    ' ── Clear existing data from output sheet ─────────────────────────────────
    ClearOutputArea wsSched

    ' ── Write Truck stops ─────────────────────────────────────────────────────
    Dim writeRow As Long
    writeRow = TRUCK_DATA_START
    Dim stopNum As Long: stopNum = 1
    Dim prevKey As String: prevKey = ""

    For i = 0 To rowCount - 1
        If rows(i).Vehicle <> "Truck" Then GoTo NextTruck

        Dim isNewStop As Boolean
        isNewStop = (rows(i).StopKey <> prevKey)

        If isNewStop And prevKey <> "" Then
            ' blank separator row between stops
            writeRow = writeRow + 1
        End If

        If writeRow > TRUCK_DATA_CLEAR Then
            MsgBox "Warning: Truck route has too many rows." & vbLf & _
                   "Some stops may overlap with Van section. Reduce or split.", vbExclamation
            Exit For
        End If

        WriteDelivRow wsSched, writeRow, rows(i), IIf(isNewStop, stopNum, 0)

        If isNewStop Then
            stopNum  = stopNum + 1
            prevKey  = rows(i).StopKey
        End If
        writeRow = writeRow + 1
NextTruck:
    Next i

    ' ── Write Van stops ───────────────────────────────────────────────────────
    writeRow = VAN_DATA_START
    stopNum  = 1
    prevKey  = ""

    For i = 0 To rowCount - 1
        If rows(i).Vehicle <> "Van" Then GoTo NextVan

        isNewStop = (rows(i).StopKey <> prevKey)

        If isNewStop And prevKey <> "" Then
            writeRow = writeRow + 1   ' blank separator
        End If

        If writeRow > VAN_DATA_CLEAR Then
            MsgBox "Warning: Van route has too many rows.", vbExclamation
            Exit For
        End If

        WriteDelivRow wsSched, writeRow, rows(i), IIf(isNewStop, stopNum, 0)

        If isNewStop Then
            stopNum = stopNum + 1
            prevKey = rows(i).StopKey
        End If
        writeRow = writeRow + 1
NextVan:
    Next i

    MsgBox "Schedule generated for " & Format(selDate, "dddd, MM/DD/YYYY") & "." & vbLf & _
           rowCount & " line items written.", vbInformation, "Done"

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation   = xlCalculationAutomatic
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.Calculation   = xlCalculationAutomatic
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error in GenerateSchedule"
End Sub


'================================================================================
' ClearSchedule  –  wipe output rows only (leaves driver info + headers intact)
'================================================================================
Public Sub ClearSchedule()
    Dim wsSched As Worksheet
    Set wsSched = GetSheet(SCHED_SHEET)
    If wsSched Is Nothing Then Exit Sub
    ClearOutputArea wsSched
    MsgBox "Schedule cleared.", vbInformation, "Cleared"
End Sub


'================================================================================
' USER-DEFINED TYPE  –  one MASTER DATA line item
'================================================================================
Private Type DelivRow
    ReqID     As String
    SoldTo    As String
    ShipTo    As String
    Address   As String
    Area      As String
    StopType  As String   ' "Delivery" or "Pick-up"
    Product   As String
    QtyCase   As Variant
    Pallets   As Double
    Storage   As String
    WinStart  As Variant
    WinEnd    As Variant
    Vehicle   As String   ' "Truck" or "Van"  (set by AssignVehicles)
    TimePri   As Integer
    AreaRank  As Integer
    SortKey   As String
    StopKey   As String   ' SoldTo||Address – groups items into stops
End Type


'================================================================================
' AssignVehicles
' Calculates total pallets per StopKey, then assigns Truck/Van to every row.
'================================================================================
Private Sub AssignVehicles(rows() As DelivRow, rowCount As Long, _
                            wsRules As Worksheet)
    ' Sum pallets per StopKey
    Dim stopPallets As Object
    Set stopPallets = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 0 To rowCount - 1
        Dim sk As String: sk = rows(i).StopKey
        If stopPallets.Exists(sk) Then
            stopPallets(sk) = stopPallets(sk) + rows(i).Pallets
        Else
            stopPallets(sk) = rows(i).Pallets
        End If
    Next i

    ' Assign vehicle to each row
    For i = 0 To rowCount - 1
        Dim totalPlt As Double
        totalPlt = stopPallets(rows(i).StopKey)

        Dim rawPref As String
        If Not wsRules Is Nothing Then
            rawPref = GetVehiclePreference(wsRules, rows(i).SoldTo)
        Else
            rawPref = "Flexible"
        End If

        Select Case LCase(Trim(rawPref))
            Case "truck"
                rows(i).Vehicle = "Truck"
            Case "van"
                rows(i).Vehicle = "Van"
            Case Else   ' Flexible
                rows(i).Vehicle = IIf(totalPlt >= FLEX_TRUCK_PALLETS, "Truck", "Van")
        End Select
    Next i
End Sub


'================================================================================
' WriteDelivRow
' Writes one product line to the output sheet.
' stopNumToWrite = 0  →  continuation row (same stop, additional product)
' stopNumToWrite > 0  →  first row of a new stop
'================================================================================
Private Sub WriteDelivRow(ws As Worksheet, outRow As Long, r As DelivRow, _
                           stopNumToWrite As Long)
    Dim isFirst As Boolean: isFirst = (stopNumToWrite > 0)
    Dim isPickup As Boolean: isPickup = (r.StopType = "Pick-up")

    ' ── Stop number & delivery type (first row of stop only) ──────────────────
    If isFirst Then
        ws.Cells(outRow, OUT_A_STOP).Value = stopNumToWrite
        ws.Cells(outRow, OUT_B_TYPE).Value = r.StopType

        ' ── Time columns ──────────────────────────────────────────────────────
        Dim hasStart As Boolean: hasStart = IsDate(r.WinStart) And Not IsEmpty(r.WinStart)
        Dim hasEnd   As Boolean: hasEnd   = IsDate(r.WinEnd)   And Not IsEmpty(r.WinEnd)

        If Not hasStart And Not hasEnd Then
            ' No time window
            ws.Cells(outRow, OUT_C_QUAL).Value  = "Anytime"
            ws.Cells(outRow, OUT_D_TIME1).Value = ""
            ws.Cells(outRow, OUT_E_TIME2).Value = ""
        ElseIf hasEnd And Not hasStart Then
            ' By <time>
            ws.Cells(outRow, OUT_C_QUAL).Value  = "By"
            ws.Cells(outRow, OUT_D_TIME1).Value = TimeValue(CDate(r.WinEnd))
            ws.Cells(outRow, OUT_E_TIME2).Value = ""
            ws.Cells(outRow, OUT_D_TIME1).NumberFormat = "h:MM AM/PM"
        ElseIf hasStart And hasEnd Then
            ' Between <start> and <end>
            ws.Cells(outRow, OUT_C_QUAL).Value  = "Between"
            ws.Cells(outRow, OUT_D_TIME1).Value = TimeValue(CDate(r.WinStart))
            ws.Cells(outRow, OUT_E_TIME2).Value = TimeValue(CDate(r.WinEnd))
            ws.Cells(outRow, OUT_D_TIME1).NumberFormat = "h:MM AM/PM"
            ws.Cells(outRow, OUT_E_TIME2).NumberFormat = "h:MM AM/PM"
        ElseIf hasStart And Not hasEnd Then
            ' After <start>
            ws.Cells(outRow, OUT_C_QUAL).Value  = "After"
            ws.Cells(outRow, OUT_D_TIME1).Value = TimeValue(CDate(r.WinStart))
            ws.Cells(outRow, OUT_E_TIME2).Value = ""
            ws.Cells(outRow, OUT_D_TIME1).NumberFormat = "h:MM AM/PM"
        End If
    End If

    ' ── Product line columns (every row) ──────────────────────────────────────
    ws.Cells(outRow, OUT_F_CUST).Value    = r.SoldTo
    ' col G (7) intentionally left empty
    ws.Cells(outRow, OUT_H_ADDR).Value    = r.Address
    ws.Cells(outRow, OUT_I_ITEM).Value    = r.Product
    ws.Cells(outRow, OUT_J_QTY).Value     = r.QtyCase
    If r.Pallets > 0 Then
        ws.Cells(outRow, OUT_K_PALLETS).Value = r.Pallets & " PLT"
    End If
    ws.Cells(outRow, OUT_L_STORAGE).Value = r.Storage
    ' col M (13) = Info for Driver – left blank

    ' ── Row background colour ─────────────────────────────────────────────────
    Dim bg As Long
    If isPickup Then
        bg = CLR_PICKUP   ' salmon pink for pick-up rows
    Else
        bg = CLR_WHITE
    End If

    Dim c As Long
    For c = 1 To 13
        With ws.Cells(outRow, c)
            .Interior.Color = bg
        End With
    Next c

    ' ── Time cells: red bold font (matching template style) ───────────────────
    If isFirst Then
        With ws.Cells(outRow, OUT_C_QUAL).Font
            If ws.Cells(outRow, OUT_C_QUAL).Value = "Anytime" Then
                .Color = RGB(255, 0, 0)
                .Bold  = True
            End If
        End With
        If ws.Cells(outRow, OUT_D_TIME1).Value <> "" Then
            ws.Cells(outRow, OUT_D_TIME1).Font.Color = RGB(255, 0, 0)
            ws.Cells(outRow, OUT_D_TIME1).Font.Bold  = True
        End If
        If ws.Cells(outRow, OUT_E_TIME2).Value <> "" Then
            ws.Cells(outRow, OUT_E_TIME2).Font.Color = RGB(255, 0, 0)
            ws.Cells(outRow, OUT_E_TIME2).Font.Bold  = True
        End If
    End If
End Sub


'================================================================================
' ClearOutputArea  –  wipes only data rows; preserves headers + driver info
'================================================================================
Private Sub ClearOutputArea(ws As Worksheet)
    ' Section 1: rows 5 to TRUCK_DATA_CLEAR
    With ws.Range(ws.Cells(TRUCK_DATA_START, 1), ws.Cells(TRUCK_DATA_CLEAR, 13))
        .ClearContents
        .Interior.Color   = CLR_WHITE
        .Font.Color        = RGB(0, 0, 0)
        .Font.Bold         = False
    End With

    ' Section 2: rows 74 to VAN_DATA_CLEAR
    With ws.Range(ws.Cells(VAN_DATA_START, 1), ws.Cells(VAN_DATA_CLEAR, 13))
        .ClearContents
        .Interior.Color   = CLR_WHITE
        .Font.Color        = RGB(0, 0, 0)
        .Font.Bold         = False
    End With
End Sub


'================================================================================
' GetAreasForDay  –  reads TABLE 1 in DRIVER_RULES (row with matching day name)
'================================================================================
Private Function GetAreasForDay(wsRules As Worksheet, selDate As Date) As String()
    Dim dayName As String
    Select Case Weekday(selDate, vbMonday)
        Case 1: dayName = "Mon"
        Case 2: dayName = "Tue"
        Case 3: dayName = "Wed"
        Case 4: dayName = "Thu"
        Case 5: dayName = "Fri"
        Case Else
            GetAreasForDay = Split("", ","): Exit Function
    End Select

    Dim r As Long
    For r = 4 To 25
        If Trim(wsRules.Cells(r, 1).Value) = dayName Then
            Dim areas() As String
            ReDim areas(0 To 5)
            Dim n As Integer: n = 0
            Dim c As Long
            For c = 2 To 7
                Dim a As String: a = Trim(wsRules.Cells(r, c).Value)
                If a <> "" Then: areas(n) = a: n = n + 1
            Next c
            ReDim Preserve areas(0 To n - 1)
            GetAreasForDay = areas
            Exit Function
        End If
    Next r
    GetAreasForDay = Split("", ",")
End Function


'================================================================================
' AreaAllowed  –  returns True if area is in the allowed list
'                 (wildcard "*" = accept all)
'================================================================================
Private Function AreaAllowed(areaVal As String, allowed() As String) As Boolean
    Dim a As Variant
    For Each a In allowed
        If CStr(a) = "*" Then AreaAllowed = True: Exit Function
        If StrComp(Trim(areaVal), Trim(CStr(a)), vbTextCompare) = 0 Then
            AreaAllowed = True: Exit Function
        End If
    Next a
    AreaAllowed = False
End Function


'================================================================================
' GetVehiclePreference  –  reads TABLE 2 in DRIVER_RULES
' Returns "Truck", "Van", or "Flexible"
'================================================================================
Private Function GetVehiclePreference(wsRules As Worksheet, soldTo As String) As String
    ' Find TABLE 2 header
    Dim hdr As Long: hdr = 0
    Dim r As Long
    For r = 10 To 60
        If InStr(1, wsRules.Cells(r, 1).Value, "Customer Name", vbTextCompare) > 0 Then
            hdr = r: Exit For
        End If
    Next r
    If hdr = 0 Then GetVehiclePreference = "Flexible": Exit Function

    For r = hdr + 1 To hdr + 60
        Dim custName As String: custName = Trim(wsRules.Cells(r, 1).Value)
        If custName = "" Then Exit For
        If InStr(1, soldTo, custName, vbTextCompare) > 0 Or _
           InStr(1, custName, soldTo, vbTextCompare) > 0 Then
            Dim pref As String: pref = Trim(wsRules.Cells(r, 2).Value)
            If pref <> "" Then GetVehiclePreference = pref: Exit Function
        End If
    Next r

    ' Built-in fallback: any "Mitsuwa" → Van
    If InStr(1, soldTo, "Mitsuwa", vbTextCompare) > 0 Then
        GetVehiclePreference = "Van": Exit Function
    End If

    GetVehiclePreference = "Flexible"
End Function


'================================================================================
' GetTimePriority  –  maps delivery window end to 1 / 2 / 3 / 9
'================================================================================
Private Function GetTimePriority(winEnd As Variant) As Integer
    If IsEmpty(winEnd) Or Not IsDate(winEnd) Then GetTimePriority = 9: Exit Function
    Dim t As Date: t = TimeValue(CDate(winEnd))
    Dim h As Integer: h = Hour(t)
    Dim m As Integer: m = Minute(t)
    If h < 8 Or (h = 8 And m = 0) Then
        GetTimePriority = 1
    ElseIf h < 12 Or (h = 12 And m = 0) Then
        GetTimePriority = 2
    ElseIf h < 15 Or (h = 15 And m = 0) Then
        GetTimePriority = 3
    Else
        GetTimePriority = 9
    End If
End Function


'================================================================================
' GetAreaRank  –  numeric order for sorting stops within a vehicle run
'================================================================================
Private Function GetAreaRank(areaVal As String) As Integer
    Dim order As Variant
    order = Array("JFC/IDC", "EAST LA", "CENTRAL LA", "SAN GABRIEL", _
                  "SOUTH BAY", "SOUTH LA", "WEST LA", "LAX", "NAX", "NORTH OC")
    Dim i As Integer
    For i = 0 To UBound(order)
        If StrComp(Trim(areaVal), CStr(order(i)), vbTextCompare) = 0 Then
            GetAreaRank = i + 1: Exit Function
        End If
    Next i
    GetAreaRank = 99
End Function


'================================================================================
' SortRows  –  insertion sort on rows() by SortKey (ascending)
'================================================================================
Private Sub SortRows(rows() As DelivRow, rowCount As Long)
    Dim i As Long, j As Long
    Dim tmp As DelivRow
    For i = 1 To rowCount - 1
        tmp = rows(i)
        j = i - 1
        Do While j >= 0 And rows(j).SortKey > tmp.SortKey
            rows(j + 1) = rows(j)
            j = j - 1
        Loop
        rows(j + 1) = tmp
    Next i
End Sub


'================================================================================
' PopulateEngine  –  writes sorted rows to DRIVER_ENGINE for audit/debug
'================================================================================
Private Sub PopulateEngine(wsEng As Worksheet, rows() As DelivRow, rowCount As Long)
    ' Clear old data (keep row 1 header, row 2 note)
    Dim lastE As Long
    lastE = wsEng.Cells(wsEng.Rows.Count, 1).End(xlUp).Row
    If lastE > 2 Then wsEng.Rows("3:" & lastE).ClearContents

    Dim i As Long
    For i = 0 To rowCount - 1
        Dim r As Long: r = i + 3
        With wsEng
            .Cells(r, 1).Value  = rows(i).ReqID
            .Cells(r, 2).Value  = rows(i).SoldTo
            .Cells(r, 3).Value  = rows(i).ShipTo
            .Cells(r, 4).Value  = rows(i).Address
            .Cells(r, 5).Value  = rows(i).Area
            .Cells(r, 6).Value  = rows(i).StopType
            .Cells(r, 7).Value  = rows(i).Product
            .Cells(r, 8).Value  = rows(i).QtyCase
            .Cells(r, 9).Value  = rows(i).Pallets
            .Cells(r, 10).Value = rows(i).Storage
            .Cells(r, 11).Value = rows(i).WinStart
            .Cells(r, 12).Value = rows(i).WinEnd
            .Cells(r, 13).Value = rows(i).Vehicle
            .Cells(r, 14).Value = rows(i).TimePri
            .Cells(r, 15).Value = rows(i).AreaRank
            .Cells(r, 16).Value = rows(i).StopKey
            .Cells(r, 17).Value = rows(i).SortKey
        End With
    Next i
End Sub


'================================================================================
' GetSheet  –  returns worksheet or Nothing (no error)
'================================================================================
Private Function GetSheet(sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
End Function
