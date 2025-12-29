Attribute VB_Name = "modSizing"
Option Explicit

' ============================================
' Noah Actuator Sizing Tool - Sizing Engine
' ============================================

' Sizing result structure (ActuatorRecord and GearboxRecord are in modHelpers)
Public Type SizingResult
    Success As Boolean
    ActuatorModel As String
    GearboxModel As String
    RPM As Double
    Ratio As Double
    OutputFlange As String
    CalcTorque As Double
    CalcThrust As Double      ' Multi-turn only (추력)
    CalcOpTime As Double
    ActualSF As Double        ' 실제 안전율 (CalcTorque / ReqTorque)
    MaxStemDim As Double      ' 최대 스템 직경 (mm)
    MotorPower_kW As Double   ' 모터 출력 (kW)
    TotalPrice As Double
    Status As String
    Alternatives As Collection
End Type

' ============================================
' Main Sizing Functions
' ============================================

Public Sub SizingAll()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim successCount As Long, failCount As Long
    Dim totalLines As Long

    On Error GoTo ErrorHandler
    
    ' Check if sheet exists
    If Not SheetExists(SH_VALVELIST) Then
        ShowError "ValveList sheet not found."
        Exit Sub
    End If
    
    Set ws = ThisWorkbook.Worksheets(SH_VALVELIST)
    lastRow = GetLastRow(ws, COL_LINENO)  ' Line No. 컬럼 기준으로 데이터 확인

    If lastRow < ROW_DATA_START Then
        ShowWarning "No valve data to size."
        Exit Sub
    End If

    ' Load settings
    Dim s As SizingSettings
    s = LoadSettings()

    If Not ValidateSettings(s) Then
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    successCount = 0
    failCount = 0
    totalLines = 0
    
    ' Count total lines first
    For i = ROW_DATA_START To lastRow
        If ws.Cells(i, COL_LINENO).value <> "" Then
            totalLines = totalLines + 1
        End If
    Next i

    Dim currentLine As Long
    currentLine = 0
    
    For i = ROW_DATA_START To lastRow
        If ws.Cells(i, COL_LINENO).value <> "" Then
            currentLine = currentLine + 1
            ShowProgress currentLine, totalLines, "Sizing line"
            
            If SizeLine(i, s) Then
                successCount = successCount + 1
            Else
                failCount = failCount + 1
            End If
        End If
    Next i

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ClearProgress

    ShowInfo "Sizing completed." & vbCrLf & _
        "Success: " & successCount & vbCrLf & _
        "Failed: " & failCount
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ClearProgress
    ShowError "Error during sizing: " & Err.Description
End Sub

Public Sub SizingSelected()
    Dim ws As Worksheet
    Dim selectedRow As Long
    Dim s As SizingSettings

    On Error GoTo ErrorHandler
    
    ' Check if sheet exists
    If Not SheetExists(SH_VALVELIST) Then
        ShowError "ValveList sheet not found."
        Exit Sub
    End If
    
    Set ws = ThisWorkbook.Worksheets(SH_VALVELIST)

    ' Get selected row
    If TypeName(Selection) = "Range" Then
        If Not Selection.Parent Is ws Then
            ShowWarning "Please select a row in ValveList."
            Exit Sub
        End If
        selectedRow = Selection.Row
    Else
        ShowWarning "Please select a row in ValveList."
        Exit Sub
    End If

    If selectedRow < ROW_DATA_START Then
        ShowWarning "Please select a data row (not header)."
        Exit Sub
    End If

    If ws.Cells(selectedRow, COL_LINENO).value = "" Then
        ShowWarning "Selected row has no data."
        Exit Sub
    End If

    ' Load settings
    s = LoadSettings()
    If Not ValidateSettings(s) Then
        Exit Sub
    End If

    If SizeLine(selectedRow, s) Then
        ShowInfo "Sizing completed for line " & (selectedRow - ROW_HEADER)
    Else
        ShowWarning "Sizing failed for line " & (selectedRow - ROW_HEADER) & vbCrLf & _
            ws.Cells(selectedRow, COL_STATUS).value
    End If
    Exit Sub
    
ErrorHandler:
    ShowError "Error during sizing: " & Err.Description
End Sub

' ============================================
' Single Line Sizing
' ============================================

Public Function SizeLine(rowNum As Long, s As SizingSettings) As Boolean
    Dim ws As Worksheet
    Dim result As SizingResult

    Set ws = ThisWorkbook.Worksheets(SH_VALVELIST)

    ' Override ActuatorType based on ValveType in this row
    Dim valveType As String
    valveType = GetCellString(ws.Cells(rowNum, COL_VALVETYPE))
    If valveType <> "" Then
        Dim derivedActType As String
        derivedActType = GetActuatorTypeFromValve(valveType)
        If derivedActType <> "" Then
            s.ActuatorType = derivedActType
        End If
    End If

    ' Read valve requirements
    Dim reqTorque As Double, reqThrust As Double
    Dim reqOpTime As Double, reqTurns As Double
    Dim reqLift As Double, reqPitch As Double

    reqTorque = GetCellDouble(ws.Cells(rowNum, COL_TORQUE))
    reqThrust = GetCellDouble(ws.Cells(rowNum, COL_THRUST))
    reqOpTime = GetCellDouble(ws.Cells(rowNum, COL_OPTIME))
    reqLift = GetCellDouble(ws.Cells(rowNum, COL_LIFT))
    reqPitch = GetCellDouble(ws.Cells(rowNum, COL_PITCH))

    ' Calculate Turns from Lift and Pitch (Multi-turn only)
    If reqPitch > 0 Then
        reqTurns = reqLift / reqPitch
    Else
        reqTurns = 0
    End If

    Dim couplingType As String
    Dim couplingDim As Double
    Dim minDim As Double, maxDim As Double

    couplingType = GetCellString(ws.Cells(rowNum, COL_COUPLINGTYPE))
    couplingDim = GetCellDouble(ws.Cells(rowNum, COL_COUPLINGDIM))

    If couplingType <> "" Then
        If Not GetCouplingLimits(couplingType, minDim, maxDim) Then
            If minDim < 0 Then
                WriteResult ws, rowNum, result, "DB_Couplings sheet not found."
            Else
                WriteResult ws, rowNum, result, "Unknown coupling type: " & couplingType
            End If
            SizeLine = False
            Exit Function
        End If

        If minDim > 0 Or maxDim > 0 Then
            If couplingDim <= 0 Then
                WriteResult ws, rowNum, result, "Coupling dimension required for " & couplingType
                SizeLine = False
                Exit Function
            End If

            If couplingDim < minDim Or couplingDim > maxDim Then
                WriteResult ws, rowNum, result, "Coupling dimension out of range (" & _
                    minDim & "-" & maxDim & " mm)"
                SizeLine = False
                Exit Function
            End If
        End If
    End If

    ' Convert units to Nm/kN
    reqTorque = ConvertTorqueToNm(reqTorque, s.TorqueUnit)
    reqThrust = ConvertThrustToKN(reqThrust, s.ThrustUnit)

    ' Apply safety factor
    reqTorque = reqTorque * s.SafetyFactor
    reqThrust = reqThrust * s.SafetyFactor

    ' Validate inputs based on actuator type
    If s.ActuatorType = "Linear" Then
        ' Linear uses Thrust instead of Torque
        If reqThrust <= 0 Then
            WriteResult ws, rowNum, result, "No thrust specified for Linear actuator"
            SizeLine = False
            Exit Function
        End If
    Else
        ' Multi-turn and Part-turn use Torque
        If reqTorque <= 0 Then
            WriteResult ws, rowNum, result, "No torque specified"
            SizeLine = False
            Exit Function
        End If
    End If

    ' Perform sizing (pass couplingDim for MaxStemDim check)
    result = FindBestActuator(reqTorque, reqThrust, reqOpTime, reqTurns, couplingDim, s)

    ' Write result to sheet (pass reqTorque and SafetyFactor for ActualSF calculation)
    WriteResult ws, rowNum, result, "", reqTorque, s.SafetyFactor

    SizeLine = result.Success
End Function

' ============================================
' Find Best Actuator
' ============================================

Private Function FindBestActuator(reqTorque As Double, reqThrust As Double, _
    reqOpTime As Double, reqTurns As Double, reqStemDim As Double, _
    s As SizingSettings) As SizingResult

    Dim wsModels As Worksheet
    Dim result As SizingResult
    Dim gearboxResult As SizingResult
    Dim i As Long, lastRow As Long
    Dim act As ActuatorRecord
    Dim bestAct As ActuatorRecord
    Dim foundDirect As Boolean
    Dim minDirectPrice As Double
    Dim minTorqueMargin As Double
    Dim calcOpTime As Double
    Dim bestCalcOpTime As Double
    Dim torqueMargin As Double
    ' Counters for direct actuator filtering (for error messages)
    Dim countDirectTorque As Long
    Dim countDirectOpTime As Long

    If Not SheetExists(SH_MODELS) Then
        result.Success = False
        result.Status = "DB_Models sheet not found."
        FindBestActuator = result
        Exit Function
    End If

    Set wsModels = ThisWorkbook.Worksheets(SH_MODELS)

    lastRow = GetLastRow(wsModels, 1)
    If lastRow < 2 Then
        result.Success = False
        result.Status = "DB_Models is empty."
        FindBestActuator = result
        Exit Function
    End If

    foundDirect = False
    minDirectPrice = MAX_PRICE
    minTorqueMargin = MAX_PRICE

    ' Phase 1: Find direct match (no gearbox)
    For i = 2 To lastRow
        ' Use common helper function
        If Not TryResolveActuator(wsModels, i, s, reqThrust, act) Then GoTo NextActuator

        ' Filter by torque (not for Linear - Linear uses Thrust)
        If s.ActuatorType <> "Linear" Then
            If act.Torque < reqTorque Then GoTo NextActuator
        End If
        countDirectTorque = countDirectTorque + 1

        ' Filter by stem dimension (direct actuator)
        If reqStemDim > 0 And act.MaxStemDim > 0 Then
            If reqStemDim > act.MaxStemDim Then GoTo NextActuator
        End If

        ' Calculate operating time (direct actuator, no gearbox: gbRatio=1)
        ' Part-turn: pass act.OpTime from DB
        ' Linear: pass act.Speed and act.Stroke from DB
        calcOpTime = CalculateOpTime(act.RPM, reqTurns, s.ActuatorType, 1, act.OpTime, act.Speed, act.Stroke)

        ' Check operating time range
        If reqOpTime > 0 Then
            If Not CheckOpTimeRange(calcOpTime, reqOpTime, s.OpTimeMinPct, s.OpTimeMaxPct) Then
                GoTo NextActuator
            End If
        End If
        countDirectOpTime = countDirectOpTime + 1

        ' Found a match - keep the lowest price (tie-breaker: smallest torque margin)
        torqueMargin = act.Torque - reqTorque

        If act.Price < minDirectPrice Or (act.Price = minDirectPrice And torqueMargin < minTorqueMargin) Then
            minDirectPrice = act.Price
            minTorqueMargin = torqueMargin
            bestAct = act
            foundDirect = True
            bestCalcOpTime = calcOpTime
        End If

NextActuator:
    Next i

    If foundDirect Then
        result.Success = True
        result.ActuatorModel = bestAct.Model
        result.GearboxModel = ""
        result.RPM = bestAct.RPM
        result.Ratio = 0  ' No gearbox
        result.OutputFlange = bestAct.OutputFlange
        result.CalcTorque = bestAct.Torque
        result.CalcThrust = bestAct.Thrust  ' Multi-turn and Linear only (Part-turn은 0)
        result.CalcOpTime = bestCalcOpTime
        result.MaxStemDim = bestAct.MaxStemDim  ' Direct: actuator's MaxStemDim
        result.MotorPower_kW = bestAct.MotorPower_kW
        result.TotalPrice = bestAct.Price
        result.Status = "OK"
    End If

    ' Phase 2: Find actuator + gearbox combination
    ' Note: Linear actuators don't use gearboxes (direct connection only)
    If s.ActuatorType = "Linear" Then
        ' Skip gearbox phase for Linear
        If foundDirect Then
            FindBestActuator = result
        Else
            result.Success = False
            result.Status = "No suitable Linear actuator found."
            FindBestActuator = result
        End If
        Exit Function
    End If

    gearboxResult = FindActuatorWithGearbox(reqTorque, reqThrust, reqOpTime, reqTurns, reqStemDim, s, _
        countDirectTorque, countDirectOpTime)

    ' Compare direct vs gearbox (choose lower total price when both are valid)
    If foundDirect And gearboxResult.Success Then
        If result.TotalPrice <= gearboxResult.TotalPrice Then
            FindBestActuator = result
        Else
            FindBestActuator = gearboxResult
        End If
    ElseIf foundDirect Then
        FindBestActuator = result
    Else
        FindBestActuator = gearboxResult
    End If
End Function

' ============================================
' Find Actuator with Gearbox
' ============================================

Private Function FindActuatorWithGearbox(reqTorque As Double, reqThrust As Double, _
    reqOpTime As Double, reqTurns As Double, reqStemDim As Double, _
    s As SizingSettings, countDirectTorque As Long, countDirectOpTime As Long) As SizingResult

    Dim wsModels As Worksheet, wsGb As Worksheet
    Dim result As SizingResult
    Dim i As Long, j As Long
    Dim modelsLastRow As Long, gbLastRow As Long
    Dim m As ModelRecord
    Dim act As ActuatorRecord
    Dim gb As GearboxRecord
    Dim bestAct As ActuatorRecord
    Dim bestGb As GearboxRecord
    Dim found As Boolean
    Dim minPrice As Double
    Dim outputTorque As Double
    Dim calcOpTime As Double
    Dim totalPrice As Double
    Dim bestCalcTorque As Double
    Dim bestCalcOpTime As Double
    ' Counters for error tracking
    Dim totalModels As Long
    Dim countType As Long, countSeries As Long
    Dim countPower As Long, countEnclosure As Long, countThrust As Long
    Dim countGbFlange As Long, countGbInputTorque As Long
    Dim countGbOutputTorque As Long, countGbOpTime As Long
    Dim hasGearboxData As Boolean
    Dim dummyPrice As Double
    Dim tempOutput As Double

    If Not SheetExists(SH_MODELS) Then
        result.Success = False
        result.Status = "DB_Models sheet not found."
        FindActuatorWithGearbox = result
        Exit Function
    End If

    If Not SheetExists(SH_GEARBOXES) Then
        result.Success = False
        result.Status = "DB_Gearboxes sheet not found."
        FindActuatorWithGearbox = result
        Exit Function
    End If

    Set wsModels = ThisWorkbook.Worksheets(SH_MODELS)
    Set wsGb = ThisWorkbook.Worksheets(SH_GEARBOXES)

    modelsLastRow = GetLastRow(wsModels, 1)
    gbLastRow = GetLastRow(wsGb, 1)
    hasGearboxData = (gbLastRow >= 2)

    found = False
    minPrice = MAX_PRICE

    For i = 2 To modelsLastRow
        m = ReadModelRecord(wsModels, i)

        If Trim$(m.Model) = "" Then GoTo NextAct2
        totalModels = totalModels + 1

        ' Filter by base specs (type, series, thrust)
        If m.ActType <> s.ActuatorType Then GoTo NextAct2
        countType = countType + 1

        If Not MatchModelRange(m.Series, s.ModelRange) Then GoTo NextAct2
        countSeries = countSeries + 1

        If s.ActuatorType = "Multi-turn" And reqThrust > 0 Then
            If m.Thrust < reqThrust Then GoTo NextAct2
        End If
        countThrust = countThrust + 1

        ' Try to resolve power and enclosure options
        If Not ResolveActuator(m, s, act) Then
            ' Track which option failed for better error messages
            If HasPowerOption(m.Model, s.Voltage, s.Phase, s.Frequency, dummyPrice) Then
                countPower = countPower + 1
            End If
            GoTo NextAct2
        End If
        countPower = countPower + 1
        countEnclosure = countEnclosure + 1

        If Not hasGearboxData Then GoTo NextAct2

        ' Now find compatible gearbox
        For j = 2 To gbLastRow
            gb = ReadGearboxRecord(wsGb, j)

            ' Use common helper for gearbox matching (flange, torque limits, stem dim)
            If Not TryMatchGearbox(act, gb, reqTorque, reqStemDim, outputTorque) Then
                ' Track partial matches for error messages
                If gb.InputFlange = act.OutputFlange Then
                    countGbFlange = countGbFlange + 1
                    If act.Torque <= gb.InputTorqueMax Then
                        countGbInputTorque = countGbInputTorque + 1
                        tempOutput = act.Torque * gb.Ratio * gb.Efficiency
                        If tempOutput >= reqTorque And tempOutput <= gb.OutputTorqueMax Then
                            countGbOutputTorque = countGbOutputTorque + 1
                        End If
                    End If
                End If
                GoTo NextGb
            End If
            countGbFlange = countGbFlange + 1
            countGbInputTorque = countGbInputTorque + 1
            countGbOutputTorque = countGbOutputTorque + 1

            ' Calculate operating time with gearbox ratio
            ' Part-turn: pass act.OpTime from DB
            ' Note: Linear doesn't use gearbox, so this function won't be called for Linear
            calcOpTime = CalculateOpTime(act.RPM, reqTurns, s.ActuatorType, gb.Ratio, act.OpTime, act.Speed, act.Stroke)

            ' Check operating time range
            If reqOpTime > 0 Then
                If Not CheckOpTimeRange(calcOpTime, reqOpTime, s.OpTimeMinPct, s.OpTimeMaxPct) Then
                    GoTo NextGb
                End If
            End If
            countGbOpTime = countGbOpTime + 1

            ' Found a match - check if it's the cheapest
            totalPrice = act.Price + gb.Price

            If totalPrice < minPrice Then
                minPrice = totalPrice
                bestAct = act
                bestGb = gb
                found = True
                bestCalcTorque = outputTorque
                bestCalcOpTime = calcOpTime
            End If

NextGb:
        Next j

NextAct2:
    Next i

    If found Then
        result.Success = True
        result.ActuatorModel = bestAct.Model
        result.GearboxModel = bestGb.Model
        result.RPM = bestAct.RPM
        result.Ratio = bestGb.Ratio
        result.OutputFlange = bestGb.OutputFlange
        result.CalcTorque = bestCalcTorque
        result.CalcOpTime = bestCalcOpTime
        result.CalcThrust = bestAct.Thrust  ' Multi-turn only
        result.MaxStemDim = bestGb.MaxStemDim  ' With gearbox: use gearbox's MaxStemDim
        result.MotorPower_kW = bestAct.MotorPower_kW
        result.TotalPrice = minPrice
        result.Status = "OK (with gearbox)"
    Else
        result.Success = False
        result.Status = BuildNoMatchReason( _
            totalModels, countType, countSeries, countPower, countEnclosure, _
            countThrust, countDirectTorque, countDirectOpTime, countGbFlange, countGbInputTorque, _
            countGbOutputTorque, countGbOpTime, reqTorque, reqThrust, _
            reqOpTime, s, hasGearboxData)
    End If

    FindActuatorWithGearbox = result
End Function

' ============================================
' Helper Functions
' ============================================

' Note: ReadActuatorRecord, ReadGearboxRecord, CalculateOpTime, CheckOpTimeRange
' are now in modHelpers.bas as Public functions

Private Sub WriteResult(ws As Worksheet, rowNum As Long, result As SizingResult, errMsg As String, _
    Optional reqTorqueWithSF As Double = 0, Optional safetyFactor As Double = 1)
    ' Use array for batch write (13 columns: COL_MODEL to COL_STATUS)
    Dim arr(1 To 13) As Variant
    Dim rng As Range
    Dim actualSF As Double

    Set rng = ws.Range(ws.Cells(rowNum, COL_MODEL), ws.Cells(rowNum, COL_STATUS))

    ' Set Ratio column to Text format to prevent "4:1" being interpreted as time
    ws.Cells(rowNum, COL_RATIO).NumberFormat = "@"

    If result.Success Then
        ' Calculate actual safety factor: CalcTorque / (reqTorque without SF)
        ' reqTorqueWithSF already has SF applied, so divide by SF to get original
        If reqTorqueWithSF > 0 And safetyFactor > 0 Then
            actualSF = result.CalcTorque / (reqTorqueWithSF / safetyFactor)
        Else
            actualSF = 0
        End If

        arr(1) = result.ActuatorModel                                       ' COL_MODEL
        arr(2) = result.GearboxModel                                        ' COL_GEARBOX
        arr(3) = result.RPM                                                 ' COL_RPM
        arr(4) = IIf(result.Ratio > 0, result.Ratio & ":1", "")            ' COL_RATIO
        arr(5) = result.OutputFlange                                        ' COL_OUTFLANGE
        arr(6) = Round(result.CalcTorque, 2)                               ' COL_CALCTORQUE
        arr(7) = IIf(result.CalcThrust > 0, Round(result.CalcThrust, 2), "") ' COL_CALCTHRUST
        arr(8) = Round(result.CalcOpTime, 2)                               ' COL_CALCOPTIME
        arr(9) = IIf(actualSF > 0, Round(actualSF, 2), "")                 ' COL_ACTUALSF
        arr(10) = IIf(result.MaxStemDim > 0, result.MaxStemDim, "")        ' COL_MAXSTEMDIM
        arr(11) = IIf(result.MotorPower_kW > 0, result.MotorPower_kW, "")  ' COL_KW
        arr(12) = result.TotalPrice                                         ' COL_PRICE
        arr(13) = result.Status                                             ' COL_STATUS
    Else
        arr(1) = ""
        arr(2) = ""
        arr(3) = ""
        arr(4) = ""
        arr(5) = ""
        arr(6) = ""
        arr(7) = ""
        arr(8) = ""
        arr(9) = ""
        arr(10) = ""
        arr(11) = ""
        arr(12) = ""
        arr(13) = IIf(errMsg <> "", errMsg, result.Status)
    End If

    rng.Value = arr
End Sub

' Note: GetCellDouble and GetCellInt are now in modHelpers.bas as Public functions
