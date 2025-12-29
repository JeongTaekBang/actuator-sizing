Attribute VB_Name = "modMain"
Option Explicit

' ============================================
' Noah Actuator Sizing Tool - Main Module
' Button handlers and main procedures
' ============================================

' ============================================
' Constants
' ============================================

Private Const APP_TITLE As String = "Noah Sizing Tool"

' ============================================
' Button Handlers
' ============================================

Public Sub btn_AddLine()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim s As SizingSettings
    Dim i As Long
    Dim startRow As Long
    Dim endRow As Long

    On Error GoTo ErrorHandler
    
    ' Check if sheet exists
    If Not SheetExists(SH_VALVELIST) Then
        ShowError "ValveList sheet not found."
        Exit Sub
    End If
    
    Set ws = ThisWorkbook.Worksheets(SH_VALVELIST)
    
    ' Load settings to get LinesToAdd and CouplingType
    s = LoadSettings()
    
    lastRow = GetLastRow(ws, COL_LINENO)
    If lastRow < ROW_DATA_START Then lastRow = ROW_HEADER
    
    startRow = lastRow + 1
    endRow = startRow + s.LinesToAdd - 1
    
    Application.ScreenUpdating = False
    
    ' Define valve types based on actuator type
    Dim valveTypes As String
    If s.ActuatorType = "Multi-turn" Then
        valveTypes = "Gate,Globe"
    ElseIf s.ActuatorType = "Linear" Then
        valveTypes = "Linear"
    Else ' Part-turn
        valveTypes = "Ball,Butterfly,Plug"
    End If
    
    ' Add lines with common settings pre-filled
    For i = startRow To endRow
        ' Set line number (auto-generated, read-only concept)
        ws.Cells(i, COL_LINENO).value = i - ROW_HEADER
        
        ' Apply common Coupling Type from Settings
        ws.Cells(i, COL_COUPLINGTYPE).value = s.CouplingType
        
        ' Set ValveType dropdown based on Actuator Type
        SetDropdown ws.Cells(i, COL_VALVETYPE), valveTypes
    Next i
    
    Application.ScreenUpdating = True
    
    ' Activate sheet and select the first new row's Tag cell (for user input)
    ws.Activate
    ws.Cells(startRow, COL_TAG).Select

    ShowInfo s.LinesToAdd & " lines added (rows " & startRow & " to " & endRow & ")" & vbCrLf & _
        "Coupling Type: " & s.CouplingType
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    ShowError "Error adding lines: " & Err.Description
End Sub

Public Sub btn_SizingAll()
    If AskYesNo("Run sizing for all lines?") Then
        SizingAll
    End If
End Sub

Public Sub btn_SizingSelected()
    SizingSelected
End Sub

Public Sub btn_Alternative()
    ShowAlternatives
End Sub

Public Sub btn_ExportDatasheet()
    ExportDatasheet
End Sub

Public Sub btn_ClearResults()
    If AskYesNo("Clear all sizing results?") Then
        ClearAllResults
    End If
End Sub

Public Sub btn_ToConfiguration()
    ' Copy selected models from ValveList to Configuration sheet for option selection
    Dim wsValve As Worksheet
    Dim wsConfig As Worksheet
    Dim lastRowValve As Long
    Dim configRow As Long
    Dim i As Long
    Dim copiedCount As Long
    
    On Error GoTo ErrorHandler
    
    ' Check if sheets exist
    If Not SheetExists(SH_VALVELIST) Then
        ShowError "ValveList sheet not found."
        Exit Sub
    End If
    
    If Not SheetExists(SH_CONFIG) Then
        ShowError "Configuration sheet not found."
        Exit Sub
    End If
    
    Set wsValve = ThisWorkbook.Worksheets(SH_VALVELIST)
    Set wsConfig = ThisWorkbook.Worksheets(SH_CONFIG)
    
    lastRowValve = GetLastRow(wsValve, COL_LINENO)
    
    If lastRowValve < ROW_DATA_START Then
        ShowWarning "No data in ValveList."
        Exit Sub
    End If
    
    ' Check if any sizing results exist
    Dim hasResults As Boolean
    hasResults = False
    For i = ROW_DATA_START To lastRowValve
        If Trim(wsValve.Cells(i, COL_MODEL).value) <> "" Then
            hasResults = True
            Exit For
        End If
    Next i
    
    If Not hasResults Then
        ShowWarning "No sizing results to copy. Please run Sizing first."
        Exit Sub
    End If
    
    ' Ask for confirmation
    If Not AskYesNo("Copy sizing results to Configuration sheet?" & vbCrLf & _
        "This will clear existing Configuration data.") Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' Clear existing Configuration data (keep header)
    Dim lastRowConfig As Long
    lastRowConfig = GetLastRow(wsConfig, CFG_COL_LINE)
    If lastRowConfig >= 2 Then
        wsConfig.Range(wsConfig.Cells(2, 1), wsConfig.Cells(lastRowConfig, CFG_COL_TOTAL)).ClearContents
    End If
    
    ' Copy data from ValveList to Configuration
    configRow = 2
    copiedCount = 0
    
    For i = ROW_DATA_START To lastRowValve
        If Trim(wsValve.Cells(i, COL_MODEL).value) <> "" Then
            ' Copy basic info
            wsConfig.Cells(configRow, CFG_COL_LINE).value = wsValve.Cells(i, COL_LINENO).value
            wsConfig.Cells(configRow, CFG_COL_TAG).value = wsValve.Cells(i, COL_TAG).value
            wsConfig.Cells(configRow, CFG_COL_MODEL).value = wsValve.Cells(i, COL_MODEL).value
            wsConfig.Cells(configRow, CFG_COL_GEARBOX).value = wsValve.Cells(i, COL_GEARBOX).value
            wsConfig.Cells(configRow, CFG_COL_BASEPRICE).value = wsValve.Cells(i, COL_PRICE).value
            
            ' Set default values for options
            wsConfig.Cells(configRow, CFG_COL_HTR).value = "No"
            wsConfig.Cells(configRow, CFG_COL_MOD).value = "No"
            wsConfig.Cells(configRow, CFG_COL_POS).value = "No"
            wsConfig.Cells(configRow, CFG_COL_LMT).value = "No"
            wsConfig.Cells(configRow, CFG_COL_EXD).value = "No"
            wsConfig.Cells(configRow, CFG_COL_PAINTING).value = "None"
            wsConfig.Cells(configRow, CFG_COL_QTY).value = 1
            
            ' Set Unit Price formula
            wsConfig.Cells(configRow, CFG_COL_UNITPRICE).Formula = _
                "=" & wsConfig.Cells(configRow, CFG_COL_BASEPRICE).Address(False, False) & _
                "+IF(" & wsConfig.Cells(configRow, CFG_COL_HTR).Address(False, False) & "=""Yes"",IFERROR(VLOOKUP(""OPT-HTR"",DB_Options!A:C,3,FALSE),0),0)" & _
                "+IF(" & wsConfig.Cells(configRow, CFG_COL_MOD).Address(False, False) & "=""Yes"",IFERROR(VLOOKUP(""OPT-MOD"",DB_Options!A:C,3,FALSE),0),0)" & _
                "+IF(" & wsConfig.Cells(configRow, CFG_COL_POS).Address(False, False) & "=""Yes"",IFERROR(VLOOKUP(""OPT-POS"",DB_Options!A:C,3,FALSE),0),0)" & _
                "+IF(" & wsConfig.Cells(configRow, CFG_COL_LMT).Address(False, False) & "=""Yes"",IFERROR(VLOOKUP(""OPT-LMT"",DB_Options!A:C,3,FALSE),0),0)" & _
                "+IF(" & wsConfig.Cells(configRow, CFG_COL_EXD).Address(False, False) & "=""Yes"",IFERROR(VLOOKUP(""OPT-EXD"",DB_Options!A:C,3,FALSE),0),0)" & _
                "+IFERROR(VLOOKUP(" & wsConfig.Cells(configRow, CFG_COL_PAINTING).Address(False, False) & _
                ",DB_Options!A:C,3,FALSE),0)"
            
            ' Set Total Price formula
            wsConfig.Cells(configRow, CFG_COL_TOTAL).Formula = _
                "=" & wsConfig.Cells(configRow, CFG_COL_UNITPRICE).Address(False, False) & "*" & _
                wsConfig.Cells(configRow, CFG_COL_QTY).Address(False, False)
            
            configRow = configRow + 1
            copiedCount = copiedCount + 1
        End If
    Next i
    
    Application.ScreenUpdating = True
    
    ' Activate Configuration sheet
    wsConfig.Activate
    wsConfig.Cells(2, CFG_COL_HTR).Select
    
    ShowInfo copiedCount & " lines copied to Configuration." & vbCrLf & _
        "Select options (Yes/No) and Painting for each line."
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    ShowError "Error copying to Configuration: " & Err.Description
End Sub

' ============================================
' Clear Results
' ============================================

Private Sub ClearAllResults()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets(SH_VALVELIST)
    lastRow = GetLastRow(ws, COL_LINENO)

    For i = ROW_DATA_START To lastRow
        ws.Cells(i, COL_MODEL).value = ""
        ws.Cells(i, COL_GEARBOX).value = ""
        ws.Cells(i, COL_RPM).value = ""
        ws.Cells(i, COL_RATIO).value = ""
        ws.Cells(i, COL_OUTFLANGE).value = ""
        ws.Cells(i, COL_CALCTORQUE).value = ""
        ws.Cells(i, COL_CALCTHRUST).value = ""
        ws.Cells(i, COL_CALCOPTIME).value = ""
        ws.Cells(i, COL_ACTUALSF).value = ""
        ws.Cells(i, COL_MAXSTEMDIM).value = ""
        ws.Cells(i, COL_KW).value = ""
        ws.Cells(i, COL_PRICE).value = ""
        ws.Cells(i, COL_STATUS).value = ""
    Next i

    ShowInfo "All results cleared."
End Sub

' ============================================
' Show Alternatives (with UserForm support)
' ============================================

Public Sub ShowAlternatives()
    Dim ws As Worksheet
    Dim selectedRow As Long
    Dim s As SizingSettings
    Dim alternatives As Collection
    Dim failReason As String

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

    ' Override ActuatorType based on ValveType in selected row
    Dim valveType As String
    valveType = GetCellString(ws.Cells(selectedRow, COL_VALVETYPE))
    If valveType <> "" Then
        Dim derivedActType As String
        derivedActType = GetActuatorTypeFromValve(valveType)
        If derivedActType <> "" Then
            s.ActuatorType = derivedActType
        End If
    End If

    ' Get valve requirements (using common helper functions)
    Dim reqTorque As Double, reqThrust As Double
    Dim reqOpTime As Double, reqTurns As Double
    Dim reqLift As Double, reqPitch As Double
    Dim reqStemDim As Double

    reqTorque = GetCellDouble(ws.Cells(selectedRow, COL_TORQUE))
    reqThrust = GetCellDouble(ws.Cells(selectedRow, COL_THRUST))
    reqOpTime = GetCellDouble(ws.Cells(selectedRow, COL_OPTIME))
    reqLift = GetCellDouble(ws.Cells(selectedRow, COL_LIFT))
    reqPitch = GetCellDouble(ws.Cells(selectedRow, COL_PITCH))
    reqStemDim = GetCellDouble(ws.Cells(selectedRow, COL_COUPLINGDIM))

    ' Calculate Turns from Lift and Pitch (Multi-turn only)
    If reqPitch > 0 Then
        reqTurns = reqLift / reqPitch
    Else
        reqTurns = 0
    End If

    ' Convert and apply safety factor
    reqTorque = ConvertTorqueToNm(reqTorque, s.TorqueUnit) * s.SafetyFactor
    reqThrust = ConvertThrustToKN(reqThrust, s.ThrustUnit) * s.SafetyFactor

    ' Validate based on actuator type
    If s.ActuatorType = "Linear" Then
        If reqThrust <= 0 Then
            ShowWarning "No thrust specified for Linear actuator."
            Exit Sub
        End If
    Else
        If reqTorque <= 0 Then
            ShowWarning "No torque specified."
            Exit Sub
        End If
    End If

    ' Find all alternatives
    Set alternatives = FindAllAlternatives(reqTorque, reqThrust, reqOpTime, reqTurns, reqStemDim, s, failReason)

    If alternatives.Count = 0 Then
        If failReason <> "" Then
            ws.Cells(selectedRow, COL_STATUS).value = failReason
            ShowWarning failReason
        Else
            ShowWarning "No alternative models found."
        End If
        Exit Sub
    End If

    ' Show alternatives using UserForm (pass reqTorque and SafetyFactor for ActualSF calculation)
    ShowAlternativesForm alternatives, selectedRow, ws, reqTorque, s.SafetyFactor
    Exit Sub
    
ErrorHandler:
    ShowError "Error finding alternatives: " & Err.Description
End Sub

' ============================================
' Show Alternatives Form
' ============================================

Private Sub ShowAlternativesForm(alternatives As Collection, selectedRow As Long, ws As Worksheet, _
    Optional reqTorqueWithSF As Double = 0, Optional safetyFactor As Double = 1)
    Dim selectedIdx As Long
    Dim alt As AlternativeRecord
    Dim arr(1 To 13) As Variant
    Dim rng As Range
    Dim actualSF As Double

    On Error GoTo ErrorHandler

    ' Load data into form and show
    frmAlternatives.LoadAlternatives alternatives, selectedRow, ws
    frmAlternatives.Show vbModal

    ' Process selection
    If frmAlternatives.UserCancelled Then
        Unload frmAlternatives
        Exit Sub
    End If

    ' Find selected index
    selectedIdx = GetAlternativeIndex(frmAlternatives.SelectedAlternative, alternatives)
    Unload frmAlternatives

    If selectedIdx < 1 Or selectedIdx > alternatives.Count Then
        Exit Sub  ' Cancelled or invalid
    End If

    ' Parse selected alternative using helper function
    alt = StringToAlternative(CStr(alternatives(selectedIdx)))

    ' Calculate actual safety factor
    If reqTorqueWithSF > 0 And safetyFactor > 0 Then
        actualSF = alt.Torque / (reqTorqueWithSF / safetyFactor)
    Else
        actualSF = 0
    End If

    ' Write to sheet using array (batch write for performance)
    Set rng = ws.Range(ws.Cells(selectedRow, COL_MODEL), ws.Cells(selectedRow, COL_STATUS))

    ' Set Ratio column to Text format to prevent "4:1" being interpreted as time
    ws.Cells(selectedRow, COL_RATIO).NumberFormat = "@"

    arr(1) = alt.ActuatorModel                                          ' COL_MODEL
    arr(2) = alt.GearboxModel                                           ' COL_GEARBOX
    arr(3) = alt.RPM                                                    ' COL_RPM
    arr(4) = IIf(alt.Ratio > 1, alt.Ratio & ":1", "")                  ' COL_RATIO
    arr(5) = alt.OutputFlange                                           ' COL_OUTFLANGE
    arr(6) = alt.Torque                                                 ' COL_CALCTORQUE
    arr(7) = IIf(alt.Thrust > 0, alt.Thrust, "")                       ' COL_CALCTHRUST
    arr(8) = alt.OpTime                                                 ' COL_CALCOPTIME
    arr(9) = IIf(actualSF > 0, Round(actualSF, 2), "")                 ' COL_ACTUALSF
    arr(10) = IIf(alt.MaxStemDim > 0, alt.MaxStemDim, "")              ' COL_MAXSTEMDIM
    arr(11) = IIf(alt.MotorPower_kW > 0, alt.MotorPower_kW, "")        ' COL_KW
    arr(12) = alt.Price                                                 ' COL_PRICE
    arr(13) = "OK (Alternative)"                                        ' COL_STATUS

    rng.value = arr

    ShowInfo "Selected: " & alt.ActuatorModel & IIf(alt.GearboxModel <> "", " + " & alt.GearboxModel, "")
    Exit Sub

ErrorHandler:
    On Error Resume Next
    Unload frmAlternatives
    ShowError "Error showing alternatives: " & Err.Description
End Sub

Private Function GetAlternativeIndex(altString As String, alternatives As Collection) As Long
    Dim i As Long
    
    GetAlternativeIndex = 0
    If altString = "" Then Exit Function
    
    For i = 1 To alternatives.Count
        If CStr(alternatives(i)) = altString Then
            GetAlternativeIndex = i
            Exit Function
        End If
    Next i
End Function

' ============================================
' Find All Alternatives
' ============================================

Public Function FindAllAlternatives(reqTorque As Double, reqThrust As Double, _
    reqOpTime As Double, reqTurns As Double, reqStemDim As Double, _
    s As SizingSettings, Optional ByRef failReason As String = "") As Collection

    Dim wsModels As Worksheet, wsGb As Worksheet
    Dim alternatives As New Collection
    Dim i As Long, j As Long
    Dim modelsLastRow As Long, gbLastRow As Long
    Dim act As ActuatorRecord
    Dim gb As GearboxRecord
    Dim alt As AlternativeRecord
    Dim outputTorque As Double
    Dim calcOpTime As Double
    ' Counters for error tracking
    Dim totalModels As Long
    Dim countType As Long, countSeries As Long
    Dim countPower As Long, countEnclosure As Long
    Dim countTorque As Long, countThrust As Long, countDirectOpTime As Long
    Dim countGbFlange As Long, countGbInputTorque As Long
    Dim countGbOutputTorque As Long, countGbOpTime As Long
    Dim hasGearboxData As Boolean

    On Error GoTo ErrorHandler
    failReason = ""

    ' Check if sheets exist
    If Not SheetExists(SH_MODELS) Or Not SheetExists(SH_GEARBOXES) Then
        failReason = "DB_Models or DB_Gearboxes sheet not found."
        Set FindAllAlternatives = alternatives
        Exit Function
    End If

    Set wsModels = ThisWorkbook.Worksheets(SH_MODELS)
    Set wsGb = ThisWorkbook.Worksheets(SH_GEARBOXES)

    modelsLastRow = GetLastRow(wsModels, 1)
    gbLastRow = GetLastRow(wsGb, 1)
    hasGearboxData = (gbLastRow >= 2)

    ' Phase 1: Direct actuators (no gearbox)
    For i = 2 To modelsLastRow
        ' Use common helper function
        If Not TryResolveActuator(wsModels, i, s, reqThrust, act) Then
            totalModels = totalModels + 1
            GoTo NextAlt1
        End If
        totalModels = totalModels + 1
        countType = countType + 1
        countSeries = countSeries + 1
        countThrust = countThrust + 1
        countPower = countPower + 1
        countEnclosure = countEnclosure + 1

        ' Torque check: skip for Linear (uses Thrust instead)
        If s.ActuatorType <> "Linear" Then
            If act.Torque < reqTorque Then GoTo NextAlt1
        End If
        countTorque = countTorque + 1

        ' Stem dimension check (direct)
        If reqStemDim > 0 And act.MaxStemDim > 0 Then
            If reqStemDim > act.MaxStemDim Then GoTo NextAlt1
        End If

        ' Direct actuator (no gearbox: gbRatio=1)
        ' Part-turn: pass act.OpTime from DB
        ' Linear: pass Speed and Stroke from DB
        calcOpTime = CalculateOpTime(act.RPM, reqTurns, s.ActuatorType, 1, act.OpTime, act.Speed, act.Stroke)

        If reqOpTime > 0 Then
            If Not CheckOpTimeRange(calcOpTime, reqOpTime, s.OpTimeMinPct, s.OpTimeMaxPct) Then
                GoTo NextAlt1
            End If
        End If
        countDirectOpTime = countDirectOpTime + 1

        ' Create alternative using helper function
        alt = CreateAlternativeDirect(act, calcOpTime)
        alternatives.Add AlternativeToString(alt)

NextAlt1:
    Next i

    ' Phase 2: Actuator + Gearbox combinations (skip for Linear - no gearbox support)
    If s.ActuatorType = "Linear" Then
        GoTo SkipGearboxPhase
    End If

    For i = 2 To modelsLastRow
        ' Use common helper function
        If Not TryResolveActuator(wsModels, i, s, reqThrust, act) Then GoTo NextAlt2

        For j = 2 To gbLastRow
            gb = ReadGearboxRecord(wsGb, j)

            ' Use common helper for gearbox matching
            If Not TryMatchGearbox(act, gb, reqTorque, reqStemDim, outputTorque) Then
                ' Track partial matches for error messages
                If Trim$(gb.Model) <> "" And gb.Ratio > 0 Then
                    If gb.InputFlange = act.OutputFlange Then
                        countGbFlange = countGbFlange + 1
                        If act.Torque <= gb.InputTorqueMax Then
                            countGbInputTorque = countGbInputTorque + 1
                        End If
                    End If
                End If
                GoTo NextGbAlt
            End If
            countGbFlange = countGbFlange + 1
            countGbInputTorque = countGbInputTorque + 1
            countGbOutputTorque = countGbOutputTorque + 1

            ' Calculate operating time with gearbox ratio
            ' Part-turn: pass act.OpTime from DB
            ' Note: Linear never reaches here (skipped above)
            calcOpTime = CalculateOpTime(act.RPM, reqTurns, s.ActuatorType, gb.Ratio, act.OpTime, act.Speed, act.Stroke)

            If reqOpTime > 0 Then
                If Not CheckOpTimeRange(calcOpTime, reqOpTime, s.OpTimeMinPct, s.OpTimeMaxPct) Then
                    GoTo NextGbAlt
                End If
            End If
            countGbOpTime = countGbOpTime + 1

            ' Create alternative using helper function
            alt = CreateAlternativeWithGearbox(act, gb, outputTorque, calcOpTime)
            alternatives.Add AlternativeToString(alt)

NextGbAlt:
        Next j

NextAlt2:
    Next i

SkipGearboxPhase:
    If alternatives.Count = 0 And failReason = "" Then
        failReason = BuildNoMatchReason( _
            totalModels, countType, countSeries, countPower, countEnclosure, _
            countThrust, countTorque, countDirectOpTime, _
            countGbFlange, countGbInputTorque, countGbOutputTorque, _
            countGbOpTime, reqTorque, reqThrust, reqOpTime, s, hasGearboxData)
    End If

    Set FindAllAlternatives = alternatives
    Exit Function

ErrorHandler:
    Set FindAllAlternatives = New Collection
End Function

' ============================================
' Note: Helper functions (ReadModelRecord, ReadGearboxRecord, ResolveActuator,
'       CalculateOpTime, GetCellDouble, GetCellInt, BuildNoMatchReason,
'       PassesModelFilters) are centralized in modHelpers.bas
' ============================================
