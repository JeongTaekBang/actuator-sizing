Attribute VB_Name = "modDatasheet"
Option Explicit

' ============================================
' Noah Actuator Sizing Tool - Datasheet Export
' ============================================

' Header row offset (rows 1-5 reserved for logo/header)
Private Const DS_HEADER_ROWS As Long = 5

Public Sub ExportDatasheet()
    Dim wsValve As Worksheet
    Dim wsTemplate As Worksheet
    Dim wbNew As Workbook
    Dim wsNew As Worksheet
    Dim lastRow As Long
    Dim validLines As Long
    Dim i As Long
    Dim col As Long
    Dim savePath As String

    On Error GoTo ErrorHandler
    
    ' Check if required sheets exist
    If Not SheetExists(SH_VALVELIST) Then
        ShowError "ValveList sheet not found."
        Exit Sub
    End If
    
    If Not SheetExists(SH_TEMPLATE) Then
        ShowError "Template_Datasheet sheet not found."
        Exit Sub
    End If
    
    Set wsValve = ThisWorkbook.Worksheets(SH_VALVELIST)
    Set wsTemplate = ThisWorkbook.Worksheets(SH_TEMPLATE)

    ' Count valid lines (with sizing results)
    lastRow = GetLastRow(wsValve, COL_LINENO)
    validLines = 0

    For i = ROW_DATA_START To lastRow
        If Trim(wsValve.Cells(i, COL_MODEL).value) <> "" Then
            validLines = validLines + 1
        End If
    Next i

    If validLines = 0 Then
        ShowWarning "No sizing results to export. Please run sizing first."
        Exit Sub
    End If

    Application.ScreenUpdating = False
    
    ' Create new workbook
    Set wbNew = Workbooks.Add
    Set wsNew = wbNew.Worksheets(1)
    wsNew.Name = "Datasheet"

    ' Copy template structure
    wsTemplate.Cells.Copy wsNew.Cells
    Application.CutCopyMode = False

    ' Load settings
    Dim s As SizingSettings
    s = LoadSettings()

    ' Fill data for each line
    col = 3 ' Start from column C (Line 1)
    Dim currentLine As Long
    currentLine = 0

    For i = ROW_DATA_START To lastRow
        If Trim(wsValve.Cells(i, COL_MODEL).value) <> "" Then
            currentLine = currentLine + 1
            ShowProgress currentLine, validLines, "Exporting line"
            
            ' Update header row with actual line number (row 6 = after 5 header rows)
            wsNew.Cells(DS_HEADER_ROWS + 1, col).value = "Line " & currentLine
            
            FillDatasheetLine wsNew, wsValve, i, col, s
            col = col + 1
        End If
    Next i
    
    ' Clear unused columns (template has "Line 1", "Line 2" by default in C, D)
    ' col now points to the first unused column
    If col <= 4 Then
        ' Less than 2 lines - clear column D and beyond
        wsNew.Range(wsNew.Columns(col), wsNew.Columns(10)).Delete
    End If
    
    ' Apply borders to all data columns (columns 3 to last used column)
    ApplyDatasheetBorders wsNew, col - 1

    ' Auto-fit columns
    wsNew.Columns.AutoFit
    
    Application.ScreenUpdating = True
    ClearProgress

    ' Ask for save location
    savePath = Application.GetSaveAsFilename( _
        InitialFileName:="Noah_Datasheet_" & Format(Now, "yyyymmdd_hhmmss"), _
        FileFilter:="Excel Files (*.xlsx), *.xlsx", _
        Title:="Save Datasheet")

    If savePath <> "False" Then
        On Error Resume Next
        wbNew.SaveAs savePath, xlOpenXMLWorkbook
        If Err.Number <> 0 Then
            ShowError "Failed to save file: " & Err.Description
            Err.Clear
        Else
            ShowInfo "Datasheet exported to:" & vbCrLf & savePath
        End If
        On Error GoTo ErrorHandler
    Else
        ShowInfo "Datasheet created but not saved."
    End If
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    ClearProgress
    ShowError "Error exporting datasheet: " & Err.Description
End Sub

Private Sub FillDatasheetLine(wsNew As Worksheet, wsValve As Worksheet, _
    valveRow As Long, col As Long, s As SizingSettings)

    Dim lineNum As Long
    Dim torqueNm As Double
    Dim thrustKN As Double
    Dim calcTorque As Double
    Dim calcThrust As Double
    Dim actModel As String
    Dim R As Long  ' Row offset helper
    
    R = DS_HEADER_ROWS  ' Add this to all row numbers
    
    lineNum = GetCellInt(wsValve.Cells(valveRow, COL_LINENO))
    If lineNum <= 0 Then
        lineNum = valveRow - ROW_HEADER
    End If

    ' Line Number
    wsNew.Cells(R + 2, col).value = lineNum

    ' Tag Number
    wsNew.Cells(R + 3, col).value = wsValve.Cells(valveRow, COL_TAG).value

    ' Quantity
    wsNew.Cells(R + 4, col).value = 1

    ' === Valve Requirements ===
    ' Type
    wsNew.Cells(R + 7, col).value = wsValve.Cells(valveRow, COL_VALVETYPE).value

    ' Size
    wsNew.Cells(R + 8, col).value = wsValve.Cells(valveRow, COL_SIZE).value

    ' Class
    wsNew.Cells(R + 9, col).value = wsValve.Cells(valveRow, COL_CLASS).value

    ' Torque (convert to Nm for display)
    torqueNm = GetCellDouble(wsValve.Cells(valveRow, COL_TORQUE))
    torqueNm = ConvertTorqueToNm(torqueNm, s.TorqueUnit)
    wsNew.Cells(R + 10, col).value = Round(torqueNm, 2)

    ' Thrust (convert to kN for display)
    thrustKN = GetCellDouble(wsValve.Cells(valveRow, COL_THRUST))
    thrustKN = ConvertThrustToKN(thrustKN, s.ThrustUnit)
    wsNew.Cells(R + 11, col).value = Round(thrustKN, 2)

    ' Coupling Type
    wsNew.Cells(R + 12, col).value = wsValve.Cells(valveRow, COL_COUPLINGTYPE).value

    ' Coupling Dimension
    wsNew.Cells(R + 13, col).value = wsValve.Cells(valveRow, COL_COUPLINGDIM).value

    ' Turns (calculated from Lift / Pitch, or 0.25 for Part-turn)
    Dim reqLift As Double, reqPitch As Double, calcTurns As Double
    Dim valveType As String
    valveType = CStr(wsValve.Cells(valveRow, COL_VALVETYPE).value)

    reqLift = GetCellDouble(wsValve.Cells(valveRow, COL_LIFT))
    reqPitch = GetCellDouble(wsValve.Cells(valveRow, COL_PITCH))

    If reqPitch > 0 Then
        ' Multi-turn: calculate from Lift / Pitch
        calcTurns = reqLift / reqPitch
        wsNew.Cells(R + 14, col).value = Round(calcTurns, 2)
    ElseIf GetActuatorTypeFromValve(valveType) = "Part-turn" Then
        ' Part-turn: 90 degrees = 0.25 turns
        wsNew.Cells(R + 14, col).value = 0.25
    End If

    ' Operating Time (required)
    wsNew.Cells(R + 15, col).value = wsValve.Cells(valveRow, COL_OPTIME).value

    ' === Equipment Offered ===
    ' Actuator Model
    actModel = CStr(wsValve.Cells(valveRow, COL_MODEL).value)
    wsNew.Cells(R + 18, col).value = actModel
    calcThrust = GetActuatorThrustByModel(actModel)

    ' Actuator Speed (RPM)
    wsNew.Cells(R + 19, col).value = wsValve.Cells(valveRow, COL_RPM).value

    ' Motor Power (kW) - for MA series (Multi-turn large)
    Dim motorKW As Double
    motorKW = GetCellDouble(wsValve.Cells(valveRow, COL_KW))
    If motorKW > 0 Then
        wsNew.Cells(R + 20, col).value = motorKW
    End If

    ' Gearbox
    Dim gbModel As String
    Dim gbRatio As Double

    gbModel = CStr(wsValve.Cells(valveRow, COL_GEARBOX).value)
    wsNew.Cells(R + 21, col).value = gbModel

    ' Gearbox Ratio (from DB, not ValveList - Excel converts "8:1" to time serial)
    If gbModel <> "" Then
        gbRatio = GetGearboxRatioByModel(gbModel)
        If gbRatio > 0 Then
            wsNew.Cells(R + 22, col).NumberFormat = "@"  ' Text format
            wsNew.Cells(R + 22, col).value = gbRatio & ":1"
        End If
    End If

    ' Output Flange
    wsNew.Cells(R + 23, col).value = wsValve.Cells(valveRow, COL_OUTFLANGE).value

    ' === Weights ===
    Dim actWeight As Double
    Dim gbWeight As Double

    ' Pass motorKW for MA series (same model has different weights by kW)
    actWeight = GetActuatorWeightByModel(actModel, motorKW)
    gbWeight = GetGearboxWeightByModel(gbModel)

    ' Actuator Weight
    If actWeight > 0 Then
        wsNew.Cells(R + 24, col).value = actWeight
    End If

    ' Gearbox Weight
    If gbWeight > 0 Then
        wsNew.Cells(R + 25, col).value = gbWeight
    End If

    ' Combination Weight
    If actWeight > 0 Or gbWeight > 0 Then
        wsNew.Cells(R + 26, col).value = actWeight + gbWeight
    End If

    ' === Actuator Performance ===
    ' Calculated Torque
    wsNew.Cells(R + 29, col).value = wsValve.Cells(valveRow, COL_CALCTORQUE).value

    If calcThrust > 0 Then
        wsNew.Cells(R + 30, col).value = Round(calcThrust, 2)
    End If

    ' Output Speed (RPM) = Actuator RPM / Gearbox Ratio
    Dim actRPM As Double
    Dim outputRPM As Double

    actRPM = GetCellDouble(wsValve.Cells(valveRow, COL_RPM))

    If gbModel <> "" Then
        ' Get ratio from DB_Gearboxes
        gbRatio = GetGearboxRatioByModel(gbModel)
        If gbRatio > 0 Then
            outputRPM = actRPM / gbRatio
        Else
            outputRPM = actRPM
        End If
    Else
        ' No gearbox
        outputRPM = actRPM
    End If

    If outputRPM > 0 Then
        wsNew.Cells(R + 31, col).value = Round(outputRPM, 2)
    End If

    ' Calculated Operating Time
    wsNew.Cells(R + 32, col).value = wsValve.Cells(valveRow, COL_CALCOPTIME).value

    ' === Safety Factors ===
    ' Requested
    wsNew.Cells(R + 35, col).value = s.SafetyFactor
    wsNew.Cells(R + 36, col).value = s.SafetyFactor

    ' Calculated (actual safety factor achieved)
    calcTorque = GetCellDouble(wsValve.Cells(valveRow, COL_CALCTORQUE))
    If torqueNm > 0 Then
        wsNew.Cells(R + 37, col).value = Round(calcTorque / torqueNm, 2)
    End If

    If thrustKN > 0 And calcThrust > 0 Then
        wsNew.Cells(R + 38, col).value = Round(calcThrust / thrustKN, 2)
    End If

    ' === Electrical Data ===
    wsNew.Cells(R + 41, col).value = s.Voltage
    wsNew.Cells(R + 42, col).value = s.Phase
    wsNew.Cells(R + 43, col).value = s.Frequency & " Hz"

    ' Get electrical data from DB_ElectricalData (normalized)
    Dim elecData As Variant
    elecData = GetActuatorElectricalData(actModel, s)

    ' elecData array: (0)StartingCurrent, (1)StartingPF, (2)RatedCurrent,
    '                 (3)AvgCurrent, (4)AvgPF, (5)AvgPower, (6)MotorPoles
    If IsArray(elecData) Then
        wsNew.Cells(R + 44, col).value = elecData(0)  ' Starting current
        wsNew.Cells(R + 45, col).value = elecData(1)  ' Starting power factor
        wsNew.Cells(R + 46, col).value = elecData(2)  ' Rated load current
        wsNew.Cells(R + 47, col).value = elecData(3)  ' Current at average load
        wsNew.Cells(R + 48, col).value = elecData(4)  ' Power factor at average load
        wsNew.Cells(R + 49, col).value = elecData(5)  ' Motor power at average load
        wsNew.Cells(R + 50, col).value = elecData(6)  ' Number of poles
    End If

End Sub

Private Function GetActuatorElectricalData(actModel As String, s As SizingSettings) As Variant
    ' Returns array of electrical data from DB_ElectricalData (normalized)
    ' Lookup by Model + Voltage + Phase + Freq combination
    ' Columns: 5=StartingCurrent, 6=StartingPF, 7=RatedCurrent,
    '          8=AvgCurrent, 9=AvgPF, 10=AvgPower, 11=MotorPoles

    Dim wsElec As Worksheet
    Dim i As Long, lastRow As Long
    Dim result(0 To 6) As Variant

    ' Initialize with empty
    For i = 0 To 6
        result(i) = ""
    Next i

    If actModel = "" Then
        GetActuatorElectricalData = result
        Exit Function
    End If

    On Error Resume Next
    If Not SheetExists(SH_ELECTRICAL) Then
        GetActuatorElectricalData = result
        Exit Function
    End If

    Set wsElec = ThisWorkbook.Worksheets(SH_ELECTRICAL)
    lastRow = GetLastRow(wsElec, 1)

    For i = 2 To lastRow
        If CStr(wsElec.Cells(i, 1).value) = actModel Then
            ' Match voltage, phase, frequency
            If GetCellInt(wsElec.Cells(i, 2)) = s.Voltage And _
               GetCellInt(wsElec.Cells(i, 3)) = s.Phase And _
               GetCellInt(wsElec.Cells(i, 4)) = s.Frequency Then
                result(0) = wsElec.Cells(i, 5).value   ' StartingCurrent_A
                result(1) = wsElec.Cells(i, 6).value   ' StartingPF
                result(2) = wsElec.Cells(i, 7).value   ' RatedCurrent_A
                result(3) = wsElec.Cells(i, 8).value   ' AvgCurrent_A
                result(4) = wsElec.Cells(i, 9).value   ' AvgPF
                result(5) = wsElec.Cells(i, 10).value  ' AvgPower_kW
                result(6) = wsElec.Cells(i, 11).value  ' MotorPoles
                Exit For
            End If
        End If
    Next i

    GetActuatorElectricalData = result
End Function

' Note: GetCellDouble is now in modHelpers.bas as a Public function

' ============================================
' Apply Borders to Datasheet
' ============================================

Private Sub ApplyDatasheetBorders(ws As Worksheet, lastDataCol As Long)
    Dim rng As Range
    Dim lastDataRow As Long

    ' Data area: from row 6 (after header) to row 55 (last electrical data)
    ' Columns: A to lastDataCol
    ' Note: Row 55 = DS_HEADER_ROWS(5) + 50 (Number of poles row)
    lastDataRow = DS_HEADER_ROWS + 50

    ' Apply borders to entire data area
    Set rng = ws.Range(ws.Cells(DS_HEADER_ROWS + 1, 1), ws.Cells(lastDataRow, lastDataCol))

    With rng.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With

    ' Make section headers bold
    Dim sectionRows As Variant
    Dim r As Variant

    ' Section header rows (relative to DS_HEADER_ROWS)
    ' Row offsets: 6=Valve Requirements, 17=Equipment Offered, 28=Performance, 34=Safety, 40=Electrical
    sectionRows = Array(6, 17, 28, 34, 40)

    For Each r In sectionRows
        ws.Cells(DS_HEADER_ROWS + r, 1).Font.Bold = True
    Next r
End Sub
