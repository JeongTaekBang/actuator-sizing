Attribute VB_Name = "modSettings"
Option Explicit

' ============================================
' Noah Actuator Sizing Tool - Settings Module
' ============================================

' Settings structure
Public Type SizingSettings
    TorqueUnit As String
    ThrustUnit As String
    Enclosure As String
    SafetyFactor As Double
    ActuatorType As String
    OperationMode As String
    Failsafe As String
    DutyCycle As String
    Voltage As Integer
    Phase As Integer
    Frequency As Integer
    OpTimeMinPct As Double
    OpTimeMaxPct As Double
    CouplingType As String
    ModelRange As String
    LinesToAdd As Integer
End Type

' Global settings variable
Public gSettings As SizingSettings

' Settings row positions in Settings sheet
Private Const ROW_TORQUE_UNIT As Integer = 4
Private Const ROW_THRUST_UNIT As Integer = 5
Private Const ROW_ENCLOSURE As Integer = 6
Private Const ROW_SAFETY_FACTOR As Integer = 7
Private Const ROW_ACT_TYPE As Integer = 8
Private Const ROW_OP_MODE As Integer = 9
Private Const ROW_FAILSAFE As Integer = 10
Private Const ROW_DUTY_CYCLE As Integer = 11
Private Const ROW_VOLTAGE As Integer = 12
Private Const ROW_PHASE As Integer = 13
Private Const ROW_FREQUENCY As Integer = 14
Private Const ROW_OPTIME_MIN As Integer = 15
Private Const ROW_OPTIME_MAX As Integer = 16
Private Const ROW_COUPLING_TYPE As Integer = 17
Private Const ROW_MODEL_RANGE As Integer = 18
Private Const ROW_LINES_TO_ADD As Integer = 19

' ============================================
' Load Settings from Sheet
' ============================================

Public Function LoadSettings() As SizingSettings
    Dim ws As Worksheet
    Dim s As SizingSettings

    ' Start with safe defaults
    With s
        .TorqueUnit = "Nm"
        .ThrustUnit = "kN"
        .Enclosure = ""
        .SafetyFactor = 1.25
        .ActuatorType = ""
        .OperationMode = ""
        .Failsafe = "None"
        .DutyCycle = "Any"
        .Voltage = 0
        .Phase = 0
        .Frequency = 0
        .OpTimeMinPct = -50
        .OpTimeMaxPct = 50
        .CouplingType = "Thrust Base - Threaded"
        .ModelRange = "All"
        .LinesToAdd = 10
    End With

    If Not SheetExists(SH_SETTINGS) Then
        ShowError "Settings sheet not found."
        LoadSettings = s
        gSettings = s
        Exit Function
    End If

    Set ws = ThisWorkbook.Worksheets(SH_SETTINGS)

    With s
        .TorqueUnit = GetCellString(ws.Cells(ROW_TORQUE_UNIT, 2))
        .ThrustUnit = GetCellString(ws.Cells(ROW_THRUST_UNIT, 2))
        .Enclosure = GetCellString(ws.Cells(ROW_ENCLOSURE, 2))
        .SafetyFactor = GetCellDouble(ws.Cells(ROW_SAFETY_FACTOR, 2))
        .ActuatorType = GetCellString(ws.Cells(ROW_ACT_TYPE, 2))
        .OperationMode = GetCellString(ws.Cells(ROW_OP_MODE, 2))
        .Failsafe = GetCellString(ws.Cells(ROW_FAILSAFE, 2))
        .DutyCycle = GetCellString(ws.Cells(ROW_DUTY_CYCLE, 2))
        .Voltage = GetCellInt(ws.Cells(ROW_VOLTAGE, 2))
        .Phase = GetCellInt(ws.Cells(ROW_PHASE, 2))
        .Frequency = GetCellInt(ws.Cells(ROW_FREQUENCY, 2))
        .OpTimeMinPct = GetCellDouble(ws.Cells(ROW_OPTIME_MIN, 2))
        .OpTimeMaxPct = GetCellDouble(ws.Cells(ROW_OPTIME_MAX, 2))
        .CouplingType = GetCellString(ws.Cells(ROW_COUPLING_TYPE, 2))
        .ModelRange = GetCellString(ws.Cells(ROW_MODEL_RANGE, 2))
        .LinesToAdd = GetCellInt(ws.Cells(ROW_LINES_TO_ADD, 2))
    End With

    ' Validate and set defaults
    If s.SafetyFactor < 1 Then s.SafetyFactor = 1.25
    If s.TorqueUnit = "" Then s.TorqueUnit = "Nm"
    If s.ThrustUnit = "" Then s.ThrustUnit = "kN"
    If s.Failsafe = "" Then s.Failsafe = "None"
    If s.DutyCycle = "" Then s.DutyCycle = "Any"
    If s.CouplingType = "" Then s.CouplingType = "Thrust Base - Threaded"
    If s.ModelRange = "" Then s.ModelRange = "All"
    If s.LinesToAdd < 1 Then s.LinesToAdd = 10

    LoadSettings = s
    gSettings = s
End Function

' ============================================
' Validate Settings
' ============================================

Public Function ValidateSettings(s As SizingSettings) As Boolean
    Dim isValid As Boolean
    isValid = True

    If s.TorqueUnit = "" Then
        ShowWarning "Torque unit is not selected."
        isValid = False
    End If

    If s.ThrustUnit = "" Then
        ShowWarning "Thrust unit is not selected."
        isValid = False
    End If

    If s.SafetyFactor < 1 Then
        ShowWarning "Safety factor should be >= 1.0"
        isValid = False
    End If

    If s.Voltage <= 0 Then
        ShowWarning "Voltage is not selected."
        isValid = False
    End If

    If s.Phase <= 0 Then
        ShowWarning "Phase is not selected."
        isValid = False
    End If

    If s.Frequency <= 0 Then
        ShowWarning "Frequency is not selected."
        isValid = False
    End If

    If s.ActuatorType = "" Then
        ShowWarning "Actuator type is not selected."
        isValid = False
    End If

    If s.Enclosure = "" Then
        ShowWarning "Enclosure is not selected."
        isValid = False
    End If

    ValidateSettings = isValid
End Function

' ============================================
' Get Settings Display String
' ============================================

Public Function GetSettingsSummary() As String
    Dim s As SizingSettings
    s = LoadSettings()

    GetSettingsSummary = "Settings Summary:" & vbCrLf & _
        "- Type: " & s.ActuatorType & vbCrLf & _
        "- Mode: " & s.OperationMode & vbCrLf & _
        "- Fail-safe: " & s.Failsafe & vbCrLf & _
        "- Duty Cycle: " & s.DutyCycle & vbCrLf & _
        "- Voltage: " & s.Voltage & "V " & s.Phase & "ph " & s.Frequency & "Hz" & vbCrLf & _
        "- Enclosure: " & s.Enclosure & vbCrLf & _
        "- Safety Factor: " & s.SafetyFactor & vbCrLf & _
        "- Op. Time Range: " & s.OpTimeMinPct & "% to " & s.OpTimeMaxPct & "%" & vbCrLf & _
        "- Coupling Type: " & s.CouplingType & vbCrLf & _
        "- Lines to Add: " & s.LinesToAdd
End Function
