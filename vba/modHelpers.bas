Attribute VB_Name = "modHelpers"
Option Explicit

' ============================================
' Noah Actuator Sizing Tool - Helper Functions
' ============================================

' ============================================
' Common Type Definitions
' ============================================

' Model record structure (flat structure - each Model x Freq x kW/RPM is a row)
' Power options and enclosure options are in separate tables
' 18 columns: 1-16 common, 17-18 Linear-specific (Speed, Stroke)
Public Type ModelRecord
    Model As String
    Series As String          ' MA, MS, NA, SA, NL
    ActType As String         ' Multi-turn, Part-turn, Linear
    MotorPower_kW As Double   ' MA series only (0 if N/A)
    ControlType As String     ' SA series only (ONOFF, PCU, SCP)
    Phase As Integer          ' MS series only (1 or 3), other series = 0
    Freq As Long              ' 50 or 60 Hz
    RPM As Double             ' Multi-turn only (Part-turn, Linear use 0)
    Torque As Double          ' Nm (Linear uses 0)
    Thrust As Double          ' kN (Multi-turn, Linear)
    OpTime As Double          ' seconds (Part-turn only, 90 degree operation time)
    DutyCycle As String       ' S2-30min, S4-25%, etc.
    OutputFlange As String    ' Linear: empty (no gearbox)
    MaxStemDim As Double      ' Max valve stem diameter (mm)
    Weight As Double
    BasePrice As Double
    Speed As Double           ' mm/sec (Linear only)
    Stroke As Double          ' mm (Linear only)
End Type

' Resolved actuator record (after joining with power/enclosure options)
Public Type ActuatorRecord
    Model As String
    Series As String
    ActType As String
    MotorPower_kW As Double ' 모터 출력 (kW)
    Torque As Double
    Thrust As Double
    RPM As Double           ' Multi-turn only
    OpTime As Double        ' Part-turn only (seconds for 90 degree)
    Speed As Double         ' mm/sec (Linear only)
    Stroke As Double        ' mm (Linear only)
    Voltage As Integer
    Phase As Integer
    Freq As Integer
    Enclosure As String
    OutputFlange As String
    MaxStemDim As Double
    Weight As Double
    Price As Double         ' BasePrice + PowerAdder + EnclosureAdder
End Type

' Gearbox record structure
Public Type GearboxRecord
    Model As String
    Ratio As Double
    InputTorqueMax As Double
    OutputTorqueMax As Double
    Efficiency As Double
    InputFlange As String
    OutputFlange As String
    MaxStemDim As Double    ' Max valve stem diameter (mm)
    Weight As Double
    Price As Double
End Type

' Alternative record structure (for ShowAlternatives)
Public Type AlternativeRecord
    ActuatorModel As String
    GearboxModel As String
    Torque As Double
    Thrust As Double
    OpTime As Double
    Price As Double
    OutputFlange As String
    RPM As Double
    Ratio As Double
    MaxStemDim As Double      ' 최대 스템 직경 (mm)
    MotorPower_kW As Double   ' 모터 출력 (kW)
End Type

' ============================================
' Constants
' ============================================

' Sizing constants
Public Const MAX_PRICE As Double = 9.9E+99  ' Used as "infinity" for price comparison

' Unit conversion constants
Public Const LBF_FT_TO_NM As Double = 1.35582
Public Const KGF_M_TO_NM As Double = 9.80665
Public Const LBF_TO_KN As Double = 0.00444822
Public Const KGF_TO_KN As Double = 0.00980665

' Sheet names
Public Const SH_SETTINGS As String = "Settings"
Public Const SH_VALVELIST As String = "ValveList"
Public Const SH_CONFIG As String = "Configuration"
' Normalized actuator DB sheets
Public Const SH_MODELS As String = "DB_Models"
Public Const SH_POWER_OPTIONS As String = "DB_PowerOptions"
Public Const SH_ENCLOSURE_OPTIONS As String = "DB_EnclosureOptions"
Public Const SH_ELECTRICAL As String = "DB_ElectricalData"
' Other DB sheets
Public Const SH_GEARBOXES As String = "DB_Gearboxes"
Public Const SH_COUPLINGS As String = "DB_Couplings"
Public Const SH_OPTIONS As String = "DB_Options"
Public Const SH_TEMPLATE As String = "Template_Datasheet"

' ValveList row constants
Public Const ROW_HEADER As Integer = 3      ' Header row (rows 1-2 reserved for buttons)
Public Const ROW_DATA_START As Integer = 4  ' First data row

' ValveList column indices (Input)
Public Const COL_LINENO As Integer = 1
Public Const COL_TAG As Integer = 2
Public Const COL_VALVETYPE As Integer = 3
Public Const COL_SIZE As Integer = 4
Public Const COL_CLASS As Integer = 5
Public Const COL_TORQUE As Integer = 6
Public Const COL_THRUST As Integer = 7
Public Const COL_COUPLINGTYPE As Integer = 8
Public Const COL_COUPLINGDIM As Integer = 9
Public Const COL_LIFT As Integer = 10         ' Lift (mm) - Multi-turn/Linear
Public Const COL_PITCH As Integer = 11        ' Pitch (mm) - Multi-turn only, Turns = Lift / Pitch
Public Const COL_OPTIME As Integer = 12

' ValveList column indices (Result)
Public Const COL_MODEL As Integer = 13
Public Const COL_GEARBOX As Integer = 14
Public Const COL_RPM As Integer = 15
Public Const COL_RATIO As Integer = 16
Public Const COL_OUTFLANGE As Integer = 17
Public Const COL_CALCTORQUE As Integer = 18
Public Const COL_CALCTHRUST As Integer = 19   ' Multi-turn only (추력)
Public Const COL_CALCOPTIME As Integer = 20
Public Const COL_ACTUALSF As Integer = 21     ' 실제 안전율 (CalcTorque / ReqTorque)
Public Const COL_MAXSTEMDIM As Integer = 22   ' 최대 스템 직경 (mm)
Public Const COL_KW As Integer = 23           ' 모터 출력 (kW)
Public Const COL_PRICE As Integer = 24
Public Const COL_STATUS As Integer = 25

' Configuration column indices
Public Const CFG_COL_LINE As Integer = 1
Public Const CFG_COL_TAG As Integer = 2
Public Const CFG_COL_MODEL As Integer = 3
Public Const CFG_COL_GEARBOX As Integer = 4
Public Const CFG_COL_BASEPRICE As Integer = 5
Public Const CFG_COL_HTR As Integer = 6
Public Const CFG_COL_MOD As Integer = 7
Public Const CFG_COL_POS As Integer = 8
Public Const CFG_COL_LMT As Integer = 9
Public Const CFG_COL_EXD As Integer = 10
Public Const CFG_COL_PAINTING As Integer = 11
Public Const CFG_COL_QTY As Integer = 12
Public Const CFG_COL_UNITPRICE As Integer = 13
Public Const CFG_COL_TOTAL As Integer = 14

' ============================================
' Unit Conversion Functions
' ============================================

Public Function ConvertTorqueToNm(value As Double, unit As String) As Double
    Select Case unit
        Case "Nm"
            ConvertTorqueToNm = value
        Case "lbf.ft"
            ConvertTorqueToNm = value * LBF_FT_TO_NM
        Case "kgf.m"
            ConvertTorqueToNm = value * KGF_M_TO_NM
        Case Else
            ConvertTorqueToNm = value
    End Select
End Function

Public Function ConvertThrustToKN(value As Double, unit As String) As Double
    Select Case unit
        Case "kN"
            ConvertThrustToKN = value
        Case "lbf"
            ConvertThrustToKN = value * LBF_TO_KN
        Case "kgf"
            ConvertThrustToKN = value * KGF_TO_KN
        Case Else
            ConvertThrustToKN = value
    End Select
End Function

' ============================================
' Database Helper Functions
' ============================================

Public Function GetLastRow(ws As Worksheet, col As Integer) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    If lastRow < 2 Then lastRow = 1
    GetLastRow = lastRow
End Function

Public Function FindInColumn(ws As Worksheet, col As Integer, value As Variant) As Long
    Dim rng As Range
    Set rng = ws.Columns(col).Find(what:=value, LookIn:=xlValues, LookAt:=xlWhole)
    If rng Is Nothing Then
        FindInColumn = 0
    Else
        FindInColumn = rng.Row
    End If
End Function

' ============================================
' Model Filter (before resolving power/enclosure)
' ============================================

Public Function PassesModelFilters(m As ModelRecord, s As SizingSettings, _
    reqThrust As Double) As Boolean
    ' Filter by model's base specs before resolving power/enclosure options
    ' Power and enclosure compatibility is checked by ResolveActuator

    PassesModelFilters = False

    ' Filter by type
    If m.ActType <> s.ActuatorType Then Exit Function

    ' Filter by model range (series)
    If Not MatchModelRange(m.Series, s.ModelRange) Then Exit Function

    ' Filter by frequency (flat structure: each model row is for specific freq)
    If m.Freq <> s.Frequency Then Exit Function

    ' Filter by phase (MS series only: Phase > 0 means torque depends on phase)
    ' Other series: Phase = 0, no filtering needed
    If m.Phase > 0 Then
        If m.Phase <> s.Phase Then Exit Function
    End If

    ' Filter by thrust (Multi-turn and Linear only)
    If (s.ActuatorType = "Multi-turn" Or s.ActuatorType = "Linear") And reqThrust > 0 Then
        If m.Thrust < reqThrust Then Exit Function
    End If

    ' Filter by Fail-safe (Spring Return = SR series only)
    If InStr(s.Failsafe, "SR") > 0 Then
        If m.Series <> "SR" Then Exit Function
    End If

    ' Filter by Duty Cycle
    If s.DutyCycle <> "Any" Then
        If InStr(s.DutyCycle, "S2") > 0 Then
            ' Intermittent - S2 duty cycle
            If InStr(m.DutyCycle, "S2") = 0 Then Exit Function
        ElseIf InStr(s.DutyCycle, "S4") > 0 Then
            ' Continuous - S4 duty cycle
            If InStr(m.DutyCycle, "S4") = 0 Then Exit Function
        End If
    End If

    ' Filter by Operation Mode (SA series ControlType)
    If m.Series = "SA" Then
        If s.OperationMode = "On-Off" Then
            If m.ControlType <> "ONOFF" Then Exit Function
        ElseIf InStr(s.OperationMode, "High-Speed") > 0 Then
            ' Modulating High-Speed = SCP only
            If m.ControlType <> "SCP" Then Exit Function
        ElseIf s.OperationMode = "Modulating" Then
            ' Modulating = PCU or SCP (exclude ONOFF)
            If m.ControlType = "ONOFF" Then Exit Function
        End If
    End If

    PassesModelFilters = True
End Function

' ============================================
' Enclosure Matching
' ============================================

Public Function MatchEnclosure(dbEnclosure As String, settingEnclosure As String) As Boolean
    ' Waterproof: IP67, IP68
    ' Explosionproof: Exd, Exde, Ex

    If settingEnclosure = "Waterproof" Then
        MatchEnclosure = (InStr(1, dbEnclosure, "IP", vbTextCompare) > 0)
    ElseIf settingEnclosure = "Explosionproof" Then
        MatchEnclosure = (InStr(1, dbEnclosure, "Ex", vbTextCompare) > 0)
    Else
        MatchEnclosure = True
    End If
End Function

Public Function MatchModelRange(dbSeries As String, settingModelRange As String) As Boolean
    ' Filter by model series (NA, SA, etc.)
    ' "All" means no filtering
    
    If settingModelRange = "All" Or settingModelRange = "" Then
        MatchModelRange = True
    Else
        MatchModelRange = (dbSeries = settingModelRange)
    End If
End Function

' ============================================
' Message Helpers
' ============================================

Public Sub ShowInfo(msg As String)
    MsgBox msg, vbInformation, "Noah Sizing Tool"
End Sub

Public Sub ShowWarning(msg As String)
    MsgBox msg, vbExclamation, "Noah Sizing Tool"
End Sub

Public Sub ShowError(msg As String)
    MsgBox msg, vbCritical, "Noah Sizing Tool"
End Sub

Public Function AskYesNo(msg As String) As Boolean
    AskYesNo = (MsgBox(msg, vbYesNo + vbQuestion, "Noah Sizing Tool") = vbYes)
End Function

' ============================================
' Common Cell Value Helpers
' ============================================

Public Function GetCellDouble(cell As Range) As Double
    If IsNumeric(cell.value) And cell.value <> "" Then
        GetCellDouble = CDbl(cell.value)
    Else
        GetCellDouble = 0
    End If
End Function

Public Function GetCellInt(cell As Range) As Integer
    If IsNumeric(cell.value) And cell.value <> "" Then
        GetCellInt = CInt(cell.value)
    Else
        GetCellInt = 0
    End If
End Function

Public Function GetCellString(cell As Range) As String
    GetCellString = Trim(CStr(cell.value))
End Function

' ============================================
' Common Record Readers
' ============================================

' ============================================
' Model Record Reader (DB_Models - normalized)
' ============================================

Public Function ReadModelRecord(ws As Worksheet, rowNum As Long) As ModelRecord
    Dim m As ModelRecord

    ' DB_Models columns (flat structure - 18 columns):
    ' 1:Model, 2:Series, 3:ActType, 4:MotorPower_kW, 5:ControlType, 6:Phase,
    ' 7:Freq, 8:RPM, 9:Torque_Nm, 10:Thrust_kN, 11:OpTime_sec,
    ' 12:DutyCycle, 13:OutputFlange, 14:MaxStemDim_mm, 15:Weight_kg, 16:BasePrice,
    ' 17:Speed_mm_sec, 18:Stroke_mm

    With m
        .Model = CStr(ws.Cells(rowNum, 1).value)
        .Series = CStr(ws.Cells(rowNum, 2).value)
        .ActType = CStr(ws.Cells(rowNum, 3).value)
        .MotorPower_kW = GetCellDouble(ws.Cells(rowNum, 4))
        .ControlType = CStr(ws.Cells(rowNum, 5).value)
        .Phase = GetCellInt(ws.Cells(rowNum, 6))
        .Freq = GetCellInt(ws.Cells(rowNum, 7))
        .RPM = GetCellDouble(ws.Cells(rowNum, 8))
        .Torque = GetCellDouble(ws.Cells(rowNum, 9))
        .Thrust = GetCellDouble(ws.Cells(rowNum, 10))
        .OpTime = GetCellDouble(ws.Cells(rowNum, 11))
        .DutyCycle = CStr(ws.Cells(rowNum, 12).value)
        .OutputFlange = CStr(ws.Cells(rowNum, 13).value)
        .MaxStemDim = GetCellDouble(ws.Cells(rowNum, 14))
        .Weight = GetCellDouble(ws.Cells(rowNum, 15))
        .BasePrice = GetCellDouble(ws.Cells(rowNum, 16))
        .Speed = GetCellDouble(ws.Cells(rowNum, 17))
        .Stroke = GetCellDouble(ws.Cells(rowNum, 18))
    End With

    ReadModelRecord = m
End Function

' ============================================
' Power Option Lookup (DB_PowerOptions)
' ============================================

Public Function HasPowerOption(modelName As String, voltage As Integer, _
    phase As Integer, freq As Integer, ByRef priceAdder As Double) As Boolean

    Dim ws As Worksheet
    Dim i As Long, lastRow As Long

    priceAdder = 0
    HasPowerOption = False

    If Not SheetExists(SH_POWER_OPTIONS) Then Exit Function

    Set ws = ThisWorkbook.Worksheets(SH_POWER_OPTIONS)
    lastRow = GetLastRow(ws, 1)

    For i = 2 To lastRow
        If CStr(ws.Cells(i, 1).value) = modelName Then
            If GetCellInt(ws.Cells(i, 2)) = voltage And _
               GetCellInt(ws.Cells(i, 3)) = phase And _
               GetCellInt(ws.Cells(i, 4)) = freq Then
                priceAdder = GetCellDouble(ws.Cells(i, 5))
                HasPowerOption = True
                Exit Function
            End If
        End If
    Next i
End Function

' ============================================
' Enclosure Option Lookup (DB_EnclosureOptions)
' ============================================

Public Function HasEnclosureOption(modelName As String, settingEnclosure As String, _
    ByRef actualEnclosure As String, ByRef priceAdder As Double) As Boolean

    Dim ws As Worksheet
    Dim i As Long, lastRow As Long
    Dim dbEnclosure As String

    actualEnclosure = ""
    priceAdder = 0
    HasEnclosureOption = False

    If Not SheetExists(SH_ENCLOSURE_OPTIONS) Then Exit Function

    Set ws = ThisWorkbook.Worksheets(SH_ENCLOSURE_OPTIONS)
    lastRow = GetLastRow(ws, 1)

    For i = 2 To lastRow
        If CStr(ws.Cells(i, 1).value) = modelName Then
            dbEnclosure = CStr(ws.Cells(i, 2).value)
            If MatchEnclosure(dbEnclosure, settingEnclosure) Then
                actualEnclosure = dbEnclosure
                priceAdder = GetCellDouble(ws.Cells(i, 3))
                HasEnclosureOption = True
                Exit Function
            End If
        End If
    Next i
End Function

' ============================================
' Resolve Actuator (Join Model + Power + Enclosure)
' ============================================

Public Function ResolveActuator(m As ModelRecord, s As SizingSettings, _
    ByRef act As ActuatorRecord) As Boolean

    Dim powerAdder As Double
    Dim enclosureAdder As Double
    Dim actualEnclosure As String

    ResolveActuator = False

    ' Check power option exists
    If Not HasPowerOption(m.Model, s.Voltage, s.Phase, s.Frequency, powerAdder) Then
        Exit Function
    End If

    ' Check enclosure option exists
    If Not HasEnclosureOption(m.Model, s.Enclosure, actualEnclosure, enclosureAdder) Then
        Exit Function
    End If

    ' Build resolved actuator record
    With act
        .Model = m.Model
        .Series = m.Series
        .ActType = m.ActType
        .MotorPower_kW = m.MotorPower_kW  ' 모터 출력
        .Torque = m.Torque
        .Thrust = m.Thrust
        .RPM = m.RPM
        .OpTime = m.OpTime          ' Part-turn: 90 degree operation time from DB
        .Speed = m.Speed            ' Linear: mm/sec
        .Stroke = m.Stroke          ' Linear: mm
        .Voltage = s.Voltage
        .Phase = s.Phase
        .Freq = s.Frequency
        .Enclosure = actualEnclosure
        .OutputFlange = m.OutputFlange
        .MaxStemDim = m.MaxStemDim
        .Weight = m.Weight
        .Price = m.BasePrice + powerAdder + enclosureAdder
    End With

    ResolveActuator = True
End Function

Public Function ReadGearboxRecord(ws As Worksheet, rowNum As Long) As GearboxRecord
    Dim gb As GearboxRecord
    
    With gb
        .Model = CStr(ws.Cells(rowNum, 1).value)
        .Ratio = GetCellDouble(ws.Cells(rowNum, 2))
        .InputTorqueMax = GetCellDouble(ws.Cells(rowNum, 3))
        .OutputTorqueMax = GetCellDouble(ws.Cells(rowNum, 4))
        .Efficiency = GetCellDouble(ws.Cells(rowNum, 5))
        .InputFlange = CStr(ws.Cells(rowNum, 6).value)
        .OutputFlange = CStr(ws.Cells(rowNum, 7).value)
        .MaxStemDim = GetCellDouble(ws.Cells(rowNum, 8))   ' Max valve stem diameter (mm)
        .Weight = GetCellDouble(ws.Cells(rowNum, 9))
        .Price = GetCellDouble(ws.Cells(rowNum, 10))
    End With
    
    ReadGearboxRecord = gb
End Function

' ============================================
' Coupling Data Helpers
' ============================================

Public Function GetCouplingLimits(couplingType As String, ByRef minDim As Double, ByRef maxDim As Double) As Boolean
    Dim ws As Worksheet
    Dim i As Long
    Dim lastRow As Long

    minDim = 0
    maxDim = 0

    If couplingType = "" Then
        GetCouplingLimits = False
        Exit Function
    End If

    If Not SheetExists(SH_COUPLINGS) Then
        minDim = -1
        maxDim = -1
        GetCouplingLimits = False
        Exit Function
    End If

    Set ws = ThisWorkbook.Worksheets(SH_COUPLINGS)
    lastRow = GetLastRow(ws, 1)

    For i = 2 To lastRow
        If CStr(ws.Cells(i, 1).value) = couplingType Then
            minDim = GetCellDouble(ws.Cells(i, 2))
            maxDim = GetCellDouble(ws.Cells(i, 3))
            GetCouplingLimits = True
            Exit Function
        End If
    Next i

    GetCouplingLimits = False
End Function

' ============================================
' Operating Time Calculation
' ============================================

Public Function CalculateOpTime(rpm As Double, turns As Double, actType As String, _
    Optional gbRatio As Double = 1, Optional actOpTime As Double = 0, _
    Optional actSpeed As Double = 0, Optional actStroke As Double = 0) As Double
    ' Operating time in seconds
    ' Reference: DMRA Formulas for quoting
    '
    ' Multi-turn (direct): Time = (Turns * 60) / RPM
    ' Multi-turn (with gearbox): Time = (Turns * Ratio * 60) / Actuator RPM
    '
    ' Part-turn (flat DB structure):
    ' Part-turn (direct): Use OpTime_sec from DB directly
    ' Part-turn (with gearbox): OpTime_sec * Ratio (gearbox slows down rotation)
    '
    ' Linear: Time = Stroke / Speed (no gearbox)

    If actType = "Multi-turn" Then
        If rpm > 0 And turns > 0 Then
            CalculateOpTime = (turns * gbRatio * 60) / rpm
        Else
            CalculateOpTime = 0
        End If
    ElseIf actType = "Linear" Then
        ' Linear: OpTime = Stroke / Speed (gearbox not used for Linear)
        If actSpeed > 0 And actStroke > 0 Then
            CalculateOpTime = actStroke / actSpeed
        Else
            CalculateOpTime = 0
        End If
    Else ' Part-turn
        ' Use OpTime from DB (actOpTime parameter)
        ' With gearbox: OpTime increases proportionally to ratio
        If actOpTime > 0 Then
            CalculateOpTime = actOpTime * gbRatio
        Else
            ' Fallback to RPM-based calculation (legacy)
            If rpm > 0 Then
                CalculateOpTime = (gbRatio * 60) / (4 * rpm)
            Else
                CalculateOpTime = 0
            End If
        End If
    End If
End Function

' ============================================
' Operating Time Range Check
' ============================================

Public Function CheckOpTimeRange(calcTime As Double, reqTime As Double, _
    minPct As Double, maxPct As Double) As Boolean
    
    Dim minTime As Double, maxTime As Double
    
    ' minPct is typically negative (e.g., -50%), maxPct is positive (e.g., +50%)
    ' Example: reqTime=60s, minPct=-50%, maxPct=+50%
    '          minTime = 60 * (1 - 0.5) = 30s
    '          maxTime = 60 * (1 + 0.5) = 90s
    minTime = reqTime * (1 + minPct / 100)
    maxTime = reqTime * (1 + maxPct / 100)
    
    ' Ensure min <= max
    If minTime > maxTime Then
        Dim temp As Double
        temp = minTime
        minTime = maxTime
        maxTime = temp
    End If
    
    CheckOpTimeRange = (calcTime >= minTime And calcTime <= maxTime)
End Function

' ============================================
' Progress Indicator
' ============================================

Public Sub ShowProgress(current As Long, total As Long, Optional prefix As String = "Processing")
    Application.StatusBar = prefix & " " & current & " of " & total & "..."
End Sub

Public Sub ClearProgress()
    Application.StatusBar = False
End Sub

' ============================================
' Sheet Existence Check
' ============================================

Public Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

' ============================================
' Actuator Lookup Helper
' ============================================

Public Function GetActuatorThrustByModel(actModel As String) As Double
    ' Get actuator thrust from DB_Models by model name
    ' Column 10 = Thrust_kN (DB_Models 18-column structure)
    Dim wsModels As Worksheet
    Dim i As Long, lastRow As Long

    GetActuatorThrustByModel = 0

    If actModel = "" Then Exit Function

    On Error Resume Next
    If Not SheetExists(SH_MODELS) Then Exit Function

    Set wsModels = ThisWorkbook.Worksheets(SH_MODELS)
    lastRow = GetLastRow(wsModels, 1)

    For i = 2 To lastRow
        If CStr(wsModels.Cells(i, 1).value) = actModel Then
            GetActuatorThrustByModel = GetCellDouble(wsModels.Cells(i, 10))  ' Thrust_kN column
            Exit Function
        End If
    Next i
End Function

Public Function GetGearboxRatioByModel(gbModel As String) As Double
    ' Get gearbox ratio from DB_Gearboxes by model name
    Dim wsGb As Worksheet
    Dim i As Long, lastRow As Long

    GetGearboxRatioByModel = 0

    If gbModel = "" Then Exit Function

    On Error Resume Next
    If Not SheetExists(SH_GEARBOXES) Then Exit Function

    Set wsGb = ThisWorkbook.Worksheets(SH_GEARBOXES)
    lastRow = GetLastRow(wsGb, 1)

    For i = 2 To lastRow
        If CStr(wsGb.Cells(i, 1).value) = gbModel Then
            GetGearboxRatioByModel = GetCellDouble(wsGb.Cells(i, 2))  ' Ratio column
            Exit Function
        End If
    Next i
End Function

Public Function GetActuatorWeightByModel(actModel As String, Optional motorKW As Double = 0) As Double
    ' Get actuator weight from DB_Models by model name (and optionally MotorPower_kW)
    ' Column 4 = MotorPower_kW, Column 15 = Weight_kg (DB_Models 18-column structure)
    '
    ' For MA series, same model has different weights by kW. If motorKW is provided,
    ' it will match both Model and MotorPower_kW. Otherwise, returns first match.
    Dim wsModels As Worksheet
    Dim i As Long, lastRow As Long
    Dim dbKW As Double

    GetActuatorWeightByModel = 0

    If actModel = "" Then Exit Function

    On Error Resume Next
    If Not SheetExists(SH_MODELS) Then Exit Function

    Set wsModels = ThisWorkbook.Worksheets(SH_MODELS)
    lastRow = GetLastRow(wsModels, 1)

    For i = 2 To lastRow
        If CStr(wsModels.Cells(i, 1).value) = actModel Then
            ' If motorKW provided, also match MotorPower_kW (column 4)
            If motorKW > 0 Then
                dbKW = GetCellDouble(wsModels.Cells(i, 4))
                If Abs(dbKW - motorKW) < 0.01 Then  ' kW match within tolerance
                    GetActuatorWeightByModel = GetCellDouble(wsModels.Cells(i, 15))  ' Weight_kg column
                    Exit Function
                End If
            Else
                ' No kW specified, return first match
                GetActuatorWeightByModel = GetCellDouble(wsModels.Cells(i, 15))  ' Weight_kg column
                Exit Function
            End If
        End If
    Next i
End Function

Public Function GetGearboxWeightByModel(gbModel As String) As Double
    ' Get gearbox weight from DB_Gearboxes by model name
    ' Column 9 = Weight_kg
    Dim wsGb As Worksheet
    Dim i As Long, lastRow As Long

    GetGearboxWeightByModel = 0

    If gbModel = "" Then Exit Function

    On Error Resume Next
    If Not SheetExists(SH_GEARBOXES) Then Exit Function

    Set wsGb = ThisWorkbook.Worksheets(SH_GEARBOXES)
    lastRow = GetLastRow(wsGb, 1)

    For i = 2 To lastRow
        If CStr(wsGb.Cells(i, 1).value) = gbModel Then
            GetGearboxWeightByModel = GetCellDouble(wsGb.Cells(i, 9))  ' Weight_kg column
            Exit Function
        End If
    Next i
End Function

' ============================================
' Valve Type to Actuator Type Mapping
' ============================================

Public Function GetActuatorTypeFromValve(valveType As String) As String
    ' Determine actuator type based on valve type
    ' Part-turn: Ball, Butterfly, Plug (90 degree rotation)
    ' Multi-turn: Gate, Globe (multi-rotation with linear motion)
    ' Linear: Linear (direct linear motion)

    Select Case valveType
        Case "Ball", "Butterfly", "Plug"
            GetActuatorTypeFromValve = "Part-turn"
        Case "Gate", "Globe"
            GetActuatorTypeFromValve = "Multi-turn"
        Case "Linear"
            GetActuatorTypeFromValve = "Linear"
        Case Else
            GetActuatorTypeFromValve = ""
    End Select
End Function

' ============================================
' Data Validation Helper
' ============================================

Public Sub SetDropdown(cell As Range, options As String)
    ' Set data validation dropdown for a cell
    ' options: comma-separated list (e.g., "Gate,Globe,Ball")
    
    On Error Resume Next
    cell.Validation.Delete
    On Error GoTo 0
    
    With cell.Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:=options
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = True
        .ErrorTitle = "Invalid Input"
        .ErrorMessage = "Please select from the list."
    End With
End Sub

' ============================================
' No-Match Reason Builder (Common)
' ============================================

Public Function BuildNoMatchReason( _
    totalAct As Long, _
    typeCount As Long, _
    seriesCount As Long, _
    powerCount As Long, _
    enclosureCount As Long, _
    thrustCount As Long, _
    torqueCount As Long, _
    opTimeCount As Long, _
    gbFlangeCount As Long, _
    gbInputTorqueCount As Long, _
    gbOutputTorqueCount As Long, _
    gbOpTimeCount As Long, _
    reqTorque As Double, _
    reqThrust As Double, _
    reqOpTime As Double, _
    s As SizingSettings, _
    hasGearboxData As Boolean) As String

    Dim needsThrust As Boolean
    needsThrust = (s.ActuatorType = "Multi-turn" And reqThrust > 0)

    If totalAct = 0 Then
        BuildNoMatchReason = "DB_Actuators is empty."
        Exit Function
    End If

    If typeCount = 0 Then
        BuildNoMatchReason = "No models match Actuator Type: " & s.ActuatorType
        Exit Function
    End If

    If seriesCount = 0 Then
        BuildNoMatchReason = "No models match Model Range: " & s.ModelRange
        Exit Function
    End If

    If powerCount = 0 Then
        BuildNoMatchReason = "No models match " & s.Voltage & "V " & _
            s.Phase & "ph " & s.Frequency & "Hz"
        Exit Function
    End If

    If enclosureCount = 0 Then
        BuildNoMatchReason = "No models match Enclosure: " & s.Enclosure
        Exit Function
    End If

    If needsThrust And thrustCount = 0 Then
        BuildNoMatchReason = "No models meet Thrust >= " & Round(reqThrust, 2) & " kN"
        Exit Function
    End If

    ' Check if direct actuators meet torque (for informational purposes)
    Dim directTorqueOK As Boolean
    directTorqueOK = (torqueCount > 0)

    ' If direct torque fails, we need gearbox. Check gearbox availability first.
    If Not directTorqueOK Then
        If Not hasGearboxData Then
            BuildNoMatchReason = "No direct actuators meet Torque >= " & Round(reqTorque, 2) & _
                " Nm. DB_Gearboxes is empty."
            Exit Function
        End If

        If gbFlangeCount = 0 Then
            BuildNoMatchReason = "No direct actuators meet Torque >= " & Round(reqTorque, 2) & _
                " Nm. No compatible gearboxes (flange mismatch)."
            Exit Function
        End If

        If gbInputTorqueCount = 0 Then
            BuildNoMatchReason = "No direct actuators meet Torque >= " & Round(reqTorque, 2) & _
                " Nm. Gearboxes exceed input torque limit."
            Exit Function
        End If

        If gbOutputTorqueCount = 0 Then
            BuildNoMatchReason = "No models or gearbox combinations meet Torque >= " & _
                Round(reqTorque, 2) & " Nm"
            Exit Function
        End If
    End If

    ' Direct or gearbox passed torque, check OpTime
    If reqOpTime > 0 Then
        If opTimeCount = 0 And gbOpTimeCount = 0 Then
            BuildNoMatchReason = "No actuators meet Op Time range (" & _
                Round(reqOpTime * (1 + s.OpTimeMinPct / 100), 1) & "~" & _
                Round(reqOpTime * (1 + s.OpTimeMaxPct / 100), 1) & " sec)"
            Exit Function
        End If
    End If

    BuildNoMatchReason = "No suitable model found."
End Function

' ============================================
' Common Sizing Helper Functions
' ============================================

Public Function TryResolveActuator(wsModels As Worksheet, rowNum As Long, _
    s As SizingSettings, reqThrust As Double, ByRef act As ActuatorRecord) As Boolean
    ' Combined function: Read model + Apply filters + Resolve power/enclosure
    ' Returns True if actuator passes all filters and is successfully resolved

    Dim m As ModelRecord

    TryResolveActuator = False

    ' Read model record
    m = ReadModelRecord(wsModels, rowNum)
    If Trim$(m.Model) = "" Then Exit Function

    ' Apply model filters (type, series, thrust)
    If Not PassesModelFilters(m, s, reqThrust) Then Exit Function

    ' Resolve power and enclosure options
    If Not ResolveActuator(m, s, act) Then Exit Function

    TryResolveActuator = True
End Function

Public Function TryMatchGearbox(act As ActuatorRecord, gb As GearboxRecord, _
    reqTorque As Double, reqStemDim As Double, ByRef outputTorque As Double) As Boolean
    ' Combined gearbox matching: flange + input torque + output torque + stem dim
    ' Returns True if gearbox is compatible, sets outputTorque

    TryMatchGearbox = False
    outputTorque = 0

    ' Check basic validity
    If Trim$(gb.Model) = "" Then Exit Function
    If gb.Ratio <= 0 Then Exit Function

    ' Check flange compatibility
    If gb.InputFlange <> act.OutputFlange Then Exit Function

    ' Check input torque limit
    If act.Torque > gb.InputTorqueMax Then Exit Function

    ' Calculate and check output torque
    outputTorque = act.Torque * gb.Ratio * gb.Efficiency
    If outputTorque < reqTorque Then Exit Function
    If outputTorque > gb.OutputTorqueMax Then Exit Function

    ' Check stem dimension
    If reqStemDim > 0 And gb.MaxStemDim > 0 Then
        If reqStemDim > gb.MaxStemDim Then Exit Function
    End If

    TryMatchGearbox = True
End Function

Public Function CreateAlternativeDirect(act As ActuatorRecord, calcOpTime As Double) As AlternativeRecord
    ' Create AlternativeRecord for direct actuator (no gearbox)

    Dim alt As AlternativeRecord

    With alt
        .ActuatorModel = act.Model
        .GearboxModel = ""
        .Torque = act.Torque
        .Thrust = act.Thrust
        .OpTime = calcOpTime
        .Price = act.Price
        .OutputFlange = act.OutputFlange
        .RPM = act.RPM
        .Ratio = 1
        .MaxStemDim = act.MaxStemDim      ' Direct: actuator's MaxStemDim
        .MotorPower_kW = act.MotorPower_kW
    End With

    CreateAlternativeDirect = alt
End Function

Public Function CreateAlternativeWithGearbox(act As ActuatorRecord, gb As GearboxRecord, _
    outputTorque As Double, calcOpTime As Double) As AlternativeRecord
    ' Create AlternativeRecord for actuator + gearbox combination

    Dim alt As AlternativeRecord

    With alt
        .ActuatorModel = act.Model
        .GearboxModel = gb.Model
        .Torque = outputTorque
        .Thrust = act.Thrust
        .OpTime = calcOpTime
        .Price = act.Price + gb.Price
        .OutputFlange = gb.OutputFlange
        .RPM = act.RPM
        .Ratio = gb.Ratio
        .MaxStemDim = gb.MaxStemDim       ' With gearbox: use gearbox's MaxStemDim
        .MotorPower_kW = act.MotorPower_kW
    End With

    CreateAlternativeWithGearbox = alt
End Function

Public Function AlternativeToString(alt As AlternativeRecord) As String
    ' Convert AlternativeRecord to pipe-delimited string (for backward compatibility)
    ' Format: Model|Gearbox|Torque|OpTime|Price|Flange|RPM|Ratio|Thrust|MaxStemDim|kW

    AlternativeToString = alt.ActuatorModel & "|" & alt.GearboxModel & "|" & _
        Round(alt.Torque, 1) & "|" & Round(alt.OpTime, 1) & "|" & alt.Price & "|" & _
        alt.OutputFlange & "|" & alt.RPM & "|" & alt.Ratio & "|" & alt.Thrust & "|" & _
        alt.MaxStemDim & "|" & alt.MotorPower_kW
End Function

Public Function StringToAlternative(altString As String) As AlternativeRecord
    ' Parse pipe-delimited string back to AlternativeRecord
    ' Format: Model|Gearbox|Torque|OpTime|Price|Flange|RPM|Ratio|Thrust|MaxStemDim|kW

    Dim parts() As String
    Dim alt As AlternativeRecord

    parts = Split(altString, "|")

    If UBound(parts) >= 10 Then
        With alt
            .ActuatorModel = parts(0)
            .GearboxModel = parts(1)
            .Torque = Val(parts(2))
            .OpTime = Val(parts(3))
            .Price = Val(parts(4))
            .OutputFlange = parts(5)
            .RPM = Val(parts(6))
            .Ratio = Val(parts(7))
            .Thrust = Val(parts(8))
            .MaxStemDim = Val(parts(9))
            .MotorPower_kW = Val(parts(10))
        End With
    ElseIf UBound(parts) >= 8 Then
        ' Backward compatibility: old format without MaxStemDim/kW
        With alt
            .ActuatorModel = parts(0)
            .GearboxModel = parts(1)
            .Torque = Val(parts(2))
            .OpTime = Val(parts(3))
            .Price = Val(parts(4))
            .OutputFlange = parts(5)
            .RPM = Val(parts(6))
            .Ratio = Val(parts(7))
            .Thrust = Val(parts(8))
            .MaxStemDim = 0
            .MotorPower_kW = 0
        End With
    End If

    StringToAlternative = alt
End Function
