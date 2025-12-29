# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Noah Actuator Sizing Tool - An Excel VBA automation system for selecting Rotork Noah electric actuators and gearbox combinations based on valve requirements (torque, thrust, operating time).

**Tech Stack**: VBA (Excel macros) + Python (workbook generation)

## Commands

### Python 환경
```bash
conda activate actuator-sizing
```
Python 스크립트 실행 전에 반드시 conda 환경 활성화 필요.

### Python - Generate Base Excel Workbook

**PowerShell (권장)**:
```powershell
cd "C:\Users\since\Dropbox\bjtPersonalProjects\actuatorSizing"; & "C:\Users\since\anaconda3\envs\actuator-sizing\python.exe" create_workbook.py
```

**Anaconda Prompt**:
```bash
conda activate actuator-sizing
python create_workbook.py
```

Creates `NoahSizing.xlsx` with all sheets, sample data, and formatting. Requires `openpyxl`.

### Excel VBA Setup
1. Convert `NoahSizing.xlsx` to `.xlsm` (macro-enabled)
2. Import VBA modules via Alt+F11 in this order:
   - `vba/modHelpers.bas` (must be first - contains shared types)
   - `vba/modSettings.bas`
   - `vba/modSizing.bas`
   - `vba/modMain.bas`
   - `vba/modDatasheet.bas`
3. Create UserForm `frmAlternatives` manually and paste code from `vba/frmAlternatives.frm`

## Architecture

### VBA Module Dependencies
```
modHelpers.bas (types: ActuatorRecord, GearboxRecord, constants, utilities)
       ↓
modSettings.bas (LoadSettings, ValidateSettings)
modSizing.bas (FindBestActuator, SizeLine - core algorithm)
modMain.bas (button handlers, Alternative selection)
       ↓
modDatasheet.bas (Excel export)
```

### Core Algorithm: Two-Phase Selection
The sizing engine in `modSizing.bas` uses a two-phase approach:

1. **Phase 1 - Direct Drive**: Find actuator without gearbox
   - Filters: ActType, Series, Voltage/Phase/Freq, Enclosure, Torque, Thrust, OpTime range, MaxStemDim

2. **Phase 2 - Gearbox Combination**: Find actuator + gearbox pair
   - Additional filters: Flange compatibility, input/output torque limits, efficiency
   - Output Torque = Actuator Torque × Ratio × Efficiency

3. **Phase 3 - Decision**: Select cheaper option between Phase 1 and Phase 2

### ValveType → ActuatorType 자동 결정

Sizing 및 Alternative 검색 시 **Settings의 "Actuator Type" 설정을 무시**하고, ValveList 각 행의 **ValveType**에 따라 ActuatorType을 자동 결정:

| ValveType | → ActuatorType |
|-----------|----------------|
| Ball, Butterfly, Plug | Part-turn |
| Gate, Globe | Multi-turn |
| Linear | Linear |

- `GetActuatorTypeFromValve()` 함수 (modHelpers.bas)
- `SizeLine()` 및 `ShowAlternatives()`에서 ValveType 읽어 오버라이드

**Settings의 "Actuator Type"**: Add Lines 버튼에서 ValveType 드롭다운 초기값 설정용으로만 사용 (편의 기능)

### Settings 전역 설정 vs ValveList 개별 입력

**Settings에서 전역으로 설정 (모든 행 공유):**
- Voltage, Phase, Frequency (전원 사양)
- Enclosure (Waterproof / Explosionproof)
- Torque Unit, Thrust Unit
- Safety Factor
- Operation Mode (On-Off / Modulating / Modulating High-Speed)
- Fail-safe (None / Close-on-Fail SR)
- Duty Cycle (Any / Intermittent S2 / Continuous S4)
- Op. Time Min/Max %

**ValveList에서 행별로 입력:**
- ValveType (→ ActuatorType 자동 결정)
- Torque, Thrust, Lift, Pitch, Op. Time (Turns = Lift / Pitch 자동 계산)
- Coupling Type, Coupling Dim

**워크플로우:**
1. Settings에서 전원/Enclosure 등 프로젝트 공통 설정
2. ValveList에 직접 데이터 입력 (Add Lines 버튼 없이도 가능)
3. Sizing 실행 → 각 행의 ValveType에 따라 적절한 모델 선택

**참고:** 한 프로젝트 내 모든 밸브가 동일 전원/Enclosure 사용 → 일반적인 플랜트 프로젝트 사용 케이스에 맞는 설계

### 가격 흐름 (Sizing → Configuration)

```
Sizing 결과:
├── 직접 연결: TotalPrice = Actuator.Price
└── Gearbox 조합: TotalPrice = Actuator.Price + Gearbox.Price

ValveList COL_PRICE (Actuator + Gearbox 합산)
    ↓ btn_ToConfiguration
Configuration CFG_COL_BASEPRICE
    ↓
Unit Price = BasePrice + 옵션들 (HTR, MOD, POS, LMT, EXD, Painting)
    ↓
Total = Unit Price × Qty
```

- **Gearbox 가격**: ValveList Price에 이미 포함
- **Configuration**: 최종 견적 시트 (옵션 추가 + 수량 반영)

### Operating Time Formulas (플랫 DB 구조 반영)
- **Multi-turn**: `Turns = Lift / Pitch` (자동 계산)
- **Multi-turn direct**: `(Turns × 60) / RPM`
- **Multi-turn + gearbox**: `(Turns × Ratio × 60) / RPM`
- **Part-turn direct**: DB의 `OpTime_sec` 직접 사용 (계산 불필요)
- **Part-turn + gearbox**: `OpTime_sec × Ratio` (기어박스가 회전 속도 감소)
- **Linear direct**: `Stroke_mm / Speed_mm_sec` (기어박스 미사용)

### Excel Sheet Structure
- **Settings**: Configuration inputs (units, safety factor, power specs)
- **ValveList**: Main input/output workspace (columns 1-12 input, 13-22 results)
- **Configuration**: Option selection and pricing with auto-calculated totals
- **Flat Actuator DB** (hidden):
  - `DB_Models`: Flat structure - each Model × Freq × kW/RPM is a row (18 columns: Model, Series, ActType, MotorPower_kW, ControlType, Phase, Freq, RPM, Torque_Nm, Thrust_kN, OpTime_sec, DutyCycle, OutputFlange, MaxStemDim_mm, Weight_kg, BasePrice, Speed_mm_sec, Stroke_mm)
  - `DB_PowerOptions`: Power configurations (Model, Voltage, Phase, Freq, PriceAdder)
  - `DB_EnclosureOptions`: Enclosure options (Model, Enclosure, PriceAdder)
  - `DB_ElectricalData`: Electrical specs for datasheet (11 columns: Model, Voltage, Phase, Freq, StartingCurrent_A, StartingPF, RatedCurrent_A, AvgCurrent_A, AvgPF, AvgPower_kW, MotorPoles) - 일부 필드는 placeholder
- **DB_Gearboxes**: Gearbox specs (flat structure - no normalization needed)
- **DB_Couplings, DB_Options**: Other database sheets
- **Template_Datasheet**: Export template

### Key Type Definitions (modHelpers.bas)
- `ModelRecord`: 플랫 구조 - Model, Series, ActType, MotorPower_kW, ControlType, Phase, Freq, RPM, Torque, Thrust, OpTime, DutyCycle, OutputFlange, MaxStemDim, Weight, BasePrice, Speed, Stroke (18 fields)
- `ActuatorRecord`: Resolved record after joining Model + PowerOption + EnclosureOption (includes RPM, OpTime, Speed, Stroke, Voltage, Phase, Freq, Enclosure, calculated Price)
- `GearboxRecord`: Model, Ratio, InputTorqueMax, OutputTorqueMax, Efficiency, InputFlange, OutputFlange, MaxStemDim, Weight, Price
- `AlternativeRecord`: ActuatorModel, GearboxModel, Torque, Thrust, OpTime, Price, OutputFlange, RPM, Ratio (for Alternative selection UI)
- `SizingSettings`: All configuration from Settings sheet

### Key Constants (modHelpers.bas)
- `MAX_PRICE`: 9.9E+99 (used as "infinity" for price comparison)

### Key Functions for Flat DB Structure
- `ReadModelRecord()`: Read from DB_Models (18 columns including Phase, Freq, OpTime, Speed, Stroke)
- `PassesModelFilters()`: Filter by ActType, Series, Phase (MS only), Freq, Thrust
- `HasPowerOption()`: Check DB_PowerOptions for model + voltage/phase/freq
- `HasEnclosureOption()`: Check DB_EnclosureOptions for model + enclosure type
- `ResolveActuator()`: Join Model + Power + Enclosure to build ActuatorRecord (includes OpTime, Speed, Stroke)
- `CalculateOpTime()`: Multi-turn uses RPM formula, Part-turn uses DB OpTime directly, Linear uses Stroke/Speed

### Common Helper Functions (modHelpers.bas)
중복 코드 제거를 위한 공통 함수:
- `TryResolveActuator()`: 모델 읽기 + 필터 + Resolve 통합 (ReadModelRecord + PassesModelFilters + ResolveActuator)
- `TryMatchGearbox()`: 기어박스 호환성 검사 통합 (flange, input/output torque, stem dim)
- `CreateAlternativeDirect()`: ActuatorRecord → AlternativeRecord (직접 연결)
- `CreateAlternativeWithGearbox()`: ActuatorRecord + GearboxRecord → AlternativeRecord (기어박스 조합)
- `AlternativeToString()`: AlternativeRecord → 파이프 구분 문자열 변환
- `StringToAlternative()`: 파이프 구분 문자열 → AlternativeRecord 파싱

**참고**: VBA에서 UDT(User Defined Type)는 Optional 매개변수로 사용 불가 → 두 함수로 분리

### Unit Conversion Constants
- lbf.ft → Nm: 1.35582
- kgf.m → Nm: 9.80665
- lbf → kN: 0.00444822
- kgf → kN: 0.00980665

## Current Limitations
- Contains sample/placeholder data (needs real Noah product specs)
- Weight output not implemented in datasheet export
- 60Hz models not fully verified

## Performance Considerations

### 선형 검색 vs Dictionary
VBA에서 Dictionary 대신 선형 검색 사용:
- 데이터 규모가 작음 (Models 200개, Gearboxes 50개 수준)
- Dictionary는 `Scripting.Runtime` 의존성 추가 필요
- 복합 키(Model+Voltage+Phase+Freq) 처리가 복잡해짐
- 현재 규모에서 성능 차이 무의미

### 예상 데이터 규모
| 시트 | 예상 행 수 | 비고 |
|------|-----------|------|
| DB_Models | 50~200 | Rotork Noah 제품군 한정 |
| DB_PowerOptions | 150~600 | 모델당 3~5개 전원옵션 |
| DB_EnclosureOptions | 100~400 | 모델당 2~3개 |
| DB_Gearboxes | 20~50 | 기어박스 종류 제한적 |
| ValveList | 10~200 | 프로젝트당 |

### 성능 병목
Excel Range I/O가 주 병목 → 배열 일괄 읽기/쓰기로 최적화됨

### 향후 최적화 (필요시)
데이터 1000+ 행으로 증가 시 고려:
- Dictionary 도입 (O(n) → O(1) 검색)
- 결과 캐싱
- Early exit 최적화

## Design Decisions (설계 결정)

### ValveList 행 레이아웃
```
Row 1-2: 버튼 영역 (Sizing All, Clear Results 등)
Row 3:   헤더 (ROW_HEADER = 3)
Row 4~:  데이터 시작 (ROW_DATA_START = 4)
```
- `COL_LINENO` 컬럼 기준으로 데이터 행 판단
- Line No.는 `행번호 - ROW_HEADER`로 자동 계산

### Default Values (기본값)
Settings 시트가 비어있거나 값이 없을 때 적용:
| 설정 | 기본값 | 비고 |
|-----|--------|-----|
| SafetyFactor | 1.25 | 최소 1.0 이상 필수 |
| OpTimeMinPct | -50% | 요구 시간의 50% ~ 150% 범위 허용 |
| OpTimeMaxPct | +50% | |
| LinesToAdd | 10 | Add Lines 버튼 클릭 시 |
| TorqueUnit | Nm | |
| ThrustUnit | kN | |
| Failsafe | None | SR 선택 시 Spring Return만 검색 |
| DutyCycle | Any | S2/S4 필터 |
| CouplingType | Thrust Base - Threaded | |
| ModelRange | All | 모든 시리즈 검색 |

### Selection Tie-breaker (동점 처리)
가격이 동일한 모델이 여러 개일 때:
1. **1순위**: 가격 (낮을수록 우선)
2. **2순위**: 토크 마진 (작을수록 우선) - `Actuator.Torque - reqTorque`

```vba
If act.Price < minDirectPrice Or _
   (act.Price = minDirectPrice And torqueMargin < minTorqueMargin) Then
```

### Alternative 표시 방식
- **정렬**: 없음 (발견 순서대로 표시)
- **개수 제한**: 없음 (모든 유효한 조합 표시)
- **데이터 형식**: 파이프(`|`) 구분 문자열
  - VBA Collection에 UDT 저장 불가 → 문자열 변환 필요
  - Format: `Model|Gearbox|Torque|OpTime|Price|Flange|RPM|Ratio|Thrust`

### Multi-turn vs Part-turn vs Linear 차이
| 항목 | Multi-turn | Part-turn | Linear |
|-----|-----------|-----------|--------|
| Torque (토크) | 사용 (Nm) | 사용 (Nm) | 사용 안 함 (0) |
| Thrust (추력) | 사용 (kN) | 사용 안 함 (0) | 사용 (kN) |
| Lift/Pitch | 사용자 입력 → Turns 자동계산 | 무시 (90° 고정) | Lift만 사용 (Stroke) |
| RPM (DB) | 사용 | 사용 안 함 (0) | 사용 안 함 (0) |
| OpTime_sec (DB) | 사용 안 함 (0) | 사용 (90° 동작시간) | 사용 안 함 (0) |
| Speed_mm_sec (DB) | 사용 안 함 (0) | 사용 안 함 (0) | 사용 (mm/sec) |
| Stroke_mm (DB) | 사용 안 함 (0) | 사용 안 함 (0) | 사용 (mm) |
| Op Time 계산 | `(Lift/Pitch × Ratio × 60) / RPM` | `OpTime_sec × Ratio` | `Stroke_mm / Speed_mm_sec` |
| Gearbox 사용 | 가능 | 가능 | 불가 (직접 연결만) |

### Ratio 표시 형식
- 기어박스 있음: `"10:1"` 형식
- 기어박스 없음 (직접 연결): 빈 문자열 `""`
- Ratio = 1일 때도 빈 문자열 (직접 연결 의미)

### Error Handling 패턴
```vba
Public Sub SomeFunction()
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' ... 작업 수행 ...

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True   ' cleanup
    Application.EnableEvents = True
    ShowError "Error: " & Err.Description
End Sub
```
- `On Error GoTo ErrorHandler` 패턴 사용
- ErrorHandler에서 Application 상태 복원 (cleanup)

### Batch Write 패턴 (성능 최적화)
```vba
' 개별 셀 쓰기 (느림) ❌
ws.Cells(row, 1).Value = "A"
ws.Cells(row, 2).Value = "B"

' 배열로 일괄 쓰기 (빠름) ✅
Dim arr(1 To 10) As Variant
arr(1) = "A"
arr(2) = "B"
ws.Range(ws.Cells(row, 1), ws.Cells(row, 10)).Value = arr
```
- `WriteResult()` 함수에서 10개 컬럼 일괄 쓰기
- Alternative 선택 시에도 동일 패턴 적용

### Application 상태 관리
장시간 작업 시 Excel UI 업데이트 비활성화:
```vba
Application.ScreenUpdating = False  ' 화면 깜빡임 방지
Application.EnableEvents = False    ' 이벤트 트리거 방지
```
- `SizingAll()`: 전체 Sizing 시 적용
- `btn_AddLine()`: 다수 행 추가 시 적용
- 작업 완료 또는 에러 시 반드시 `True`로 복원

### Global Settings 변수
```vba
Public gSettings As SizingSettings  ' modSettings.bas
```
- `LoadSettings()` 호출 시 자동 캐싱
- 실제로는 매번 시트에서 다시 로드 (캐시 무효화 불필요)
- 향후 성능 최적화 시 활용 가능

### ISO 5211 플랜지 호환성 (자동 보장)

**핵심 원리**: 토크 요구사항 충족 = 플랜지 호환 자동 보장

ISO 5211 규격에서 플랜지 크기(F07, F10, F14, F16, F25 등)는 토크 범위에 따라 정의됨:
- 낮은 토크 → 작은 플랜지 (F07, F10)
- 높은 토크 → 큰 플랜지 (F14, F16, F25)

```
토크 60Nm 요구 → NA006 선택 (60Nm) → OutputFlange = F07
토크 1000Nm 요구 → NA100 선택 (1000Nm) → OutputFlange = F14
```

**결론**:
- 밸브 플랜지 요구사항 별도 입력 불필요
- 토크 기준 모델 선택 → 모델의 플랜지가 자동으로 적합
- 현재 로직에서 플랜지 검증 로직 추가 불필요

**현재 플랜지 관련 체크**:
| 체크 항목 | 위치 | 설명 |
|----------|------|------|
| 액추에이터 ↔ 기어박스 | `TryMatchGearbox()` | `gb.InputFlange = act.OutputFlange` |
| 결과 출력 | COL_OUTFLANGE (17) | 최종 출력 플랜지 표시 |

### No-Match Reason 추적
Sizing 실패 시 구체적 원인 제공을 위해 각 필터 단계별 카운터 유지:
```vba
Dim countType As Long       ' ActType 매칭 개수
Dim countSeries As Long     ' Series 매칭 개수
Dim countPower As Long      ' Power option 매칭 개수
Dim countEnclosure As Long  ' Enclosure 매칭 개수
...
```
- `BuildNoMatchReason()` 함수에서 0인 카운터 기준으로 실패 원인 결정
- 사용자에게 "No models match 380V 3ph 50Hz" 같은 구체적 메시지 제공

**에러 메시지 우선순위** (먼저 실패한 단계의 메시지 출력):

| 순서 | 조건 | 메시지 예시 |
|------|------|------------|
| 1 | totalAct = 0 | DB_Actuators is empty |
| 2 | typeCount = 0 | No models match Actuator Type: Part-turn |
| 3 | seriesCount = 0 | No models match Model Range |
| 4 | powerCount = 0 | No models match 380V 3ph 50Hz |
| 5 | enclosureCount = 0 | No models match Enclosure: Waterproof |
| 6 | thrustCount = 0 | No models meet Thrust >= X kN |
| 7 | torqueCount = 0 (직접+기어박스) | No models or gearbox combinations meet Torque >= X Nm |
| 8 | opTimeCount = 0 AND gbOpTimeCount = 0 | No actuators meet Op Time range (5~15 sec) |
| 9 | (default) | No suitable model found |

**카운터 흐름**:
- `countDirectTorque`: FindBestActuator Phase 1에서 토크 통과한 직접 연결 개수
- `countDirectOpTime`: 직접 연결 중 OpTime 범위 통과 개수
- `countGb*`: FindActuatorWithGearbox에서 기어박스 조합 단계별 통과 개수

**설계 원칙**: 직접 연결/기어박스 모두 시도 후, 실패 단계에 맞는 메시지 출력

## Real Product Data Integration Checklist

When replacing sample data with actual Noah product specifications, verify:

### DB_Models (플랫 구조 - 18 컬럼)
- All columns: Model, Series, ActType, MotorPower_kW, ControlType, Phase, Freq, RPM, Torque_Nm, Thrust_kN, OpTime_sec, DutyCycle, OutputFlange, MaxStemDim_mm, Weight_kg, BasePrice, Speed_mm_sec, Stroke_mm
- 플랫 구조: 각 Model × Freq × kW/RPM 조합이 별도 행
- **Phase 컬럼**: MS 시리즈 전용 (1 또는 3), 다른 시리즈는 0 (Phase가 토크에 영향 없음)
- Multi-turn: RPM 사용, OpTime_sec = 0, Torque + Thrust
- Part-turn: OpTime_sec 사용 (90° 동작시간), RPM = 0, Torque만
- Linear: Speed_mm_sec + Stroke_mm 사용, Thrust만, Torque = 0
- Thrust_kN: Multi-turn, Linear만 사용 (Part-turn은 0)
- OutputFlange must match gearbox InputFlange for compatibility (e.g., F07, F10, F14, F16, F25)
- Linear는 OutputFlange 없음 (기어박스 미사용)

### DB_PowerOptions (Power Configurations)
- Each Model × Voltage × Phase × Freq combination needs a row
- PriceAdder: Additional cost for this power option (can be negative for cheaper options)
- Standard option typically has PriceAdder = 0

### DB_EnclosureOptions (Enclosure Types)
- Each Model × Enclosure combination needs a row
- Enclosure 매칭 로직 (부분 문자열 검색):
  - "IP67", "IP68" 등 → Settings "Waterproof"와 매칭 ("IP" 포함 여부)
  - "Exd", "Exde" 등 → Settings "Explosionproof"와 매칭 ("Ex" 포함 여부)
- PriceAdder: Additional cost for this enclosure (e.g., 0 for standard IP67, +300 for Exd)

### DB_ElectricalData (For Datasheet Export)
- Each Model × Voltage × Phase × Freq combination needs electrical data
- **11 Columns** (4 lookup keys + 7 data fields):

| Column | Description | Data Status |
|--------|-------------|-------------|
| Model | Actuator model | Key |
| Voltage | V | Key |
| Phase | 1, 3, or DC | Key |
| Freq | Hz (50/60) | Key |
| StartingCurrent_A | Starting current | ⚠️ Placeholder (None) |
| StartingPF | Starting power factor | ⚠️ Placeholder (None) |
| RatedCurrent_A | Rated load current | ✅ Available |
| AvgCurrent_A | Current at average load | ⚠️ Placeholder (None) |
| AvgPF | Power factor at average load | ⚠️ Placeholder (None) |
| AvgPower_kW | Motor power at average load | ✅ Available |
| MotorPoles | Number of motor poles | ✅ Available |

**Placeholder 필드 처리**:
- ⚠️ 표시된 필드는 카탈로그 데이터 확보 시 채워넣을 수 있음
- 데이터 없으면 Datasheet에서 빈칸으로 출력됨
- 필요없는 필드는 향후 삭제 가능 (modDatasheet.bas 수정 필요)

**유지보수 방법**:
1. **필드 추가**: create_workbook.py의 `add_row()` 함수 및 headers 수정 → modDatasheet.bas 컬럼 인덱스 업데이트
2. **필드 삭제**: 동일하게 양쪽 파일 동기화 필요
3. **데이터 직접 입력**: Excel에서 DB_ElectricalData 시트 직접 편집 가능

### DB_Gearboxes
- InputFlange/OutputFlange must use same naming convention as actuators
- Efficiency as decimal (0.85-0.95 typical)
- InputTorqueMax: Maximum actuator torque the gearbox can accept
- OutputTorqueMax: Maximum output torque (should be > InputTorqueMax × Ratio × Efficiency)

### DB_Couplings
- CouplingType names must exactly match dropdown values in ValveList
- MinDimension_mm/MaxDimension_mm define valid valve stem diameter range

### Sizing Validation
- Test with edge cases: maximum torque, minimum RPM, boundary operating times
- Verify Phase 1 (direct) vs Phase 2 (gearbox) selection chooses cheaper option correctly
- Confirm OpTime range filtering works (Settings: Op. Time Min/Max %)
- Check MaxStemDim validation: CouplingDim ≤ Actuator.MaxStemDim (direct) or Gearbox.MaxStemDim (with gearbox)
- Verify price calculation: BasePrice + PowerAdder + EnclosureAdder

## Option 추가 요청 시

새 옵션 추가가 필요하면 다음 정보 제공:
- 옵션 코드 (예: `OPT-ENCODER`)
- 옵션 설명 (예: `Absolute Encoder`)
- 타입: `Yes/No` 또는 `드롭다운`
- (드롭다운인 경우) 선택지 목록 (예: `None, Type-A, Type-B`)

Claude가 업데이트할 파일:
- `create_workbook.py` - DB_Options 시트에 옵션 행 추가
- `vba/modMain.bas` - Configuration 수식에 VLOOKUP 추가
- `vba/modHelpers.bas` - 컬럼 상수 추가

가격은 Excel에서 직접 DB_Options 시트의 Price 컬럼 수정.

## Actuator DB 구조 (플랫 구조)

Noah 시리즈별 특성이 다르므로 플랫 구조를 사용합니다.

**Model 컬럼 규칙**: 카탈로그 원래 모델명 사용 (예: NA006, MA01, SR05)
- kW, Hz, RPM 등은 별도 컬럼으로 구분
- 같은 Model이 여러 행에 존재 (Freq, kW, RPM 조합별로)
- VBA에서 Model + Freq + 기타 조건으로 필터링

### 플랫 구조 선택 이유
- VBA 코드 단순화 (PassesModelFilters에서 Freq 등 필터링)
- 데이터 규모 적절함 (~260행, 선형 검색에 문제없음)
- 사용자 입력 변경 없음 (kW, RPM은 자동 선정)
- Excel 데이터 관리 용이 (카탈로그와 모델명 일치)

### Noah 시리즈별 특성

| 시리즈 | 타입 | 복합 키 | 토크/추력 범위 | 전원 특성 |
|--------|------|---------|----------------|-----------|
| MA | Multi-turn | Model + kW + Hz + RPM | 40~15,680 Nm | 3상 전용 |
| MS | Multi-turn | Model + Phase + Hz + RPM | 45~110 Nm | 1상/3상 |
| NA | Part-turn | Model + Hz | 60~3,500 Nm | DC/AC |
| SA | Part-turn | Model + Type + Hz | 30~90 Nm | 1상 전용 |
| SR | Part-turn (Spring Return) | Model + Hz | 50~500 Nm | AC/DC |
| NL | Linear | Model + Hz | 4~35 kN | DC/AC |

**참조 문서**: `docs/noah_torque_tables.md`

### DB_Models (플랫 구조)

각 조합이 별도 행으로 등록 (18 컬럼):
```
| Model | Series | ActType | MotorPower_kW | ControlType | Phase | Freq | RPM | Torque_Nm | Thrust_kN | OpTime_sec | DutyCycle | OutputFlange | MaxStemDim_mm | Weight_kg | BasePrice | Speed_mm_sec | Stroke_mm |
```

- 컬럼 1~16: 기존 (Multi-turn, Part-turn 공용)
- 컬럼 17~18: Linear 전용 (Speed_mm_sec, Stroke_mm)
- **Phase 컬럼**: MS 시리즈만 1 또는 3 사용, 다른 시리즈는 0

#### 시리즈별 예시

**MA (Multi-turn)**: 같은 Model이 kW × Hz × RPM 조합별로 여러 행
```
| MA01 | MA | Multi-turn | 0.2 | - | 0 | 50 | 16 | 88 | 50 | - | S2-30min | F10 | 65 | 45 | 1000 |
| MA01 | MA | Multi-turn | 0.2 | - | 0 | 50 | 20 | 88 | 50 | - | S2-30min | F10 | 65 | 45 | 1000 |
| MA01 | MA | Multi-turn | 0.4 | - | 0 | 50 | 16 | 138 | 60 | - | S2-30min | F10 | 65 | 45 | 1200 |
```
- Phase=0 (MA는 3상 전용, Phase가 토크에 영향 없음)

**MS (Multi-turn, 소형)**: Phase에 따라 Torque 다름 → Phase 컬럼으로 구분
```
| MS01 | MS | Multi-turn | - | - | 3 | 50 | 20.5 | 110 | 30 | - | S2-30min | F07 | 40 | 15 | 600 |
| MS01 | MS | Multi-turn | - | - | 1 | 50 | 19 | 45 | 20 | - | S2-15min | F07 | 40 | 15 | 500 |
```
- Phase=3 (3상, 110Nm), Phase=1 (1상, 45Nm)
- 동일 Model명 "MS01" 사용, Phase 컬럼으로 구분 (일관성 있는 모델명 체계)

**NA (Part-turn)**: 같은 Model이 Hz별로 여러 행
```
| NA006 | NA | Part-turn | - | - | 0 | 50 | - | 60 | - | 18 | S4-50% | F07 | 22 | 11 | 500 |
| NA006 | NA | Part-turn | - | - | 0 | 60 | - | 60 | - | 16 | S4-50% | F07 | 22 | 11 | 500 |
| NA100 | NA | Part-turn | - | - | 0 | 50 | - | 1000 | - | 38 | S4-25% | F14 | 42 | 25 | 1500 |
```
- Phase=0 (Phase가 토크에 영향 없음)
- RPM 컬럼 비워둠 (Part-turn은 OpTime_sec 직접 사용)

**SA (Part-turn, 소형)**: 카탈로그 모델명 그대로 사용
```
| SA005 | SA | Part-turn | - | ONOFF | 0 | 50 | - | 50 | - | 17 | S2-15min | F07 | 20 | 2.8 | 400 | - | - |
| SA005L | SA | Part-turn | - | SCP | 0 | 50 | - | 50 | - | 8 | S2-15min | F07 | 20 | 4.3 | 550 | - | - |
| SA05X | SA | Part-turn | - | ONOFF | 0 | 50 | - | 50 | - | 17 | S2-15min | F07 | 20 | 5 | 600 | - | - |
```
- Phase=0 (SA는 1상 전용, Phase가 토크에 영향 없음)
- **ControlType**: ONOFF, PCU, SCP (제어 방식)
  - ONOFF: 기본 개폐제어
  - PCU: 비례제어 (Proportional Control Unit)
  - SCP: 고속 비례제어 (Stepping Control Panel) - 동작시간 절반
- **방폭형 모델 (SA05X, SA09X)**: ControlType은 ONOFF, Enclosure만 Exd

**SR (Part-turn, Spring Return)**: 같은 Model이 Hz별로 여러 행
```
| SR05 | SR | Part-turn | - | SR | 0 | 50 | - | 50 | - | 17 | S2-15min | F07 | 20 | 26.5 | 600 | - | - |
| SR05 | SR | Part-turn | - | SR | 0 | 60 | - | 50 | - | 14 | S2-15min | F07 | 20 | 26.5 | 600 | - | - |
| SR10 | SR | Part-turn | - | SR | 0 | 50 | - | 100 | - | 20 | S2-15min | F07 | 22 | 35 | 700 | - | - |
| SR50 | SR | Part-turn | - | SR | 0 | 50 | - | 500 | - | 116 | S2-15min | F10 | 42 | 82 | 1200 | - | - |
```
- Phase=0 (Phase가 토크에 영향 없음)
- **ControlType**: SR (Spring Return 구분용)
- **OpTime**: 모터 구동 시간 (Open 방향) - Close는 스프링으로 1-2초
- **Spring Return Time**: 카탈로그 참조 (1-2초, 사이징에서 미사용)
- **방폭형 (X 접미사)**: SR05X, SR10X 등 - Enclosure=Exd로 처리

**NL (Linear)**: 같은 Model이 Hz별로 여러 행
```
| NL04 | NL | Linear | - | - | 0 | 50 | - | - | 4 | - | S4-50% | - | 20 | 16 | 500 | 0.8 | 40 |
| NL04 | NL | Linear | - | - | 0 | 60 | - | - | 4 | - | S4-50% | - | 20 | 16 | 500 | 0.93 | 40 |
| NL20 | NL | Linear | - | - | 0 | 50 | - | - | 20 | - | S4-30% | - | 24 | 31 | 1000 | 0.85 | 100 |
| NL35 | NL | Linear | - | - | 0 | 50 | - | - | 35 | - | S4-20% | - | 24 | 31 | 1200 | 0.4 | 100 |
```
- Phase=0 (Phase가 토크/추력에 영향 없음)
- **Output**: Torque 대신 **Thrust (kN)** 사용
- **OpTime 계산**: `Stroke_mm / Speed_mm_sec`
- **Gearbox 미사용**: 직접 연결만 지원
- **OutputFlange**: 없음 (직접 마운트)

### ModelRecord 타입 (modHelpers.bas)

```vba
Public Type ModelRecord
    Model As String           ' 카탈로그 모델명 (예: MA01, MS01, NA006, NL04) - 중복 가능
    Series As String          ' MA, MS, NA, SA, NL, SR
    ActType As String         ' Multi-turn, Part-turn, Linear
    MotorPower_kW As Double   ' MA 전용 (0이면 N/A)
    ControlType As String     ' SA 전용 (ONOFF, PCU, SCP)
    Phase As Integer          ' MS 전용 (1 또는 3), 다른 시리즈는 0
    Freq As Long              ' 50 또는 60
    RPM As Double             ' Multi-turn용 (Part-turn, Linear는 0)
    Torque As Double          ' Nm (Linear는 0)
    Thrust As Double          ' kN (Multi-turn, Linear)
    OpTime As Double          ' 초 (Part-turn 전용, 90° 동작시간)
    DutyCycle As String       ' S2-30min, S4-25% 등
    OutputFlange As String    ' Linear는 빈값
    MaxStemDim As Double      ' mm
    Weight As Double          ' kg
    BasePrice As Double
    Speed As Double           ' mm/sec (Linear 전용)
    Stroke As Double          ' mm (Linear 전용)
End Type
```

### DB_PowerOptions (전압/주파수)

모델별 지원 전원 옵션:
```
| Model | Voltage | Phase | Freq | PriceAdder |
|-------|---------|-------|------|------------|
| NA006 | 380 | 3 | 50 | 0 |
| NA006 | 220 | 1 | 50 | 0 |
| NA006 | 12 | DC | 50 | 50 |
| SA005 | 220 | 1 | 50 | 0 |
| MS01 | 380 | 3 | 50 | 0 |
| MS01 | 220 | 1 | 50 | 0 |
```
- SA는 1상 전용 (380V/440V 불가)
- MS는 Phase에 따라 Torque가 다름 → DB_Models의 Phase 컬럼으로 구분

### DB_EnclosureOptions (보호등급)

모델별 지원 Enclosure:
```
| Model | Enclosure | PriceAdder |
|-------|-----------|------------|
| NA006 | IP67 | 0 |
| NA006 | Exd | 300 |
| SA005 | IP67 | 0 |
| SA005 | Exd | 200 |
| SA05X | Exd | 0 |
| MS01 | IP67 | 0 |
| MS01 | Exd | 250 |
```
- 방폭형 모델 (SA05X, SA09X): Exd만 지원 (BasePrice에 포함)
- 일반 모델: IP67 기본, Exd 옵션
- MS01: 동일 Enclosure 옵션이 3상/1상 모두에 적용

### DB_ElectricalData (Datasheet용)

모델 × 전원 조합별 전기 특성 (11컬럼):
```
| Model | Voltage | Phase | Freq | StartingCurrent_A | StartingPF | RatedCurrent_A | AvgCurrent_A | AvgPF | AvgPower_kW | MotorPoles |
|-------|---------|-------|------|-------------------|------------|----------------|--------------|-------|-------------|------------|
| NA006 | 380 | 3 | 50 | (비어있음) | (비어있음) | 0.15 | (비어있음) | (비어있음) | 0.015 | 4 |
| NA006 | 220 | 1 | 50 | (비어있음) | (비어있음) | 0.42 | (비어있음) | (비어있음) | 0.015 | 4 |
| MS01 | 380 | 3 | 50 | (비어있음) | (비어있음) | 0.8 | (비어있음) | (비어있음) | 0.2 | 4 |
```

**참고**:
- `(비어있음)` 필드는 카탈로그에서 데이터 확보 시 채워넣을 수 있음
- Datasheet 출력 시 빈칸으로 표시됨
- 불필요한 필드는 DB와 modDatasheet.bas 모두에서 삭제 가능

### 가격 계산

최종 가격 = BasePrice + PowerAdder + EnclosureAdder

### 사이징 로직 변경 사항

1. **Operating Time 계산**
   - Multi-turn: `(Turns × 60) / RPM` (기존 그대로)
   - Part-turn: DB의 `OpTime_sec` 직접 사용 (계산 불필요)

2. **필터 조건**
   - Freq 필터 추가 (Settings의 Frequency와 매칭)
   - Phase 필터 추가 (MS 시리즈만: DB_Models.Phase = Settings.Phase)
   - Operation Mode 필터: On-Off vs Modulating
     - On-Off: SA 시리즈의 ONOFF 모델
     - Modulating: SA 시리즈의 PCU/SCP 모델

3. **비례제어 제품**
   - MA/MS: 최대 토크의 50%로 선정 (카탈로그 지침)
   - Settings "Operation Mode" = Modulating 시 적용

### 데이터 입력 시 주의사항

- Model 컬럼: 카탈로그 모델명 그대로 사용 (예: MS01, NA006) - 복합 키로 중복 가능
- Phase 컬럼: MS 시리즈만 1 또는 3 입력, 다른 시리즈는 0
- 숫자 컬럼은 숫자 형식으로 (텍스트 "100" ❌ → 숫자 100 ✅)
- `ActType`: 정확히 `Multi-turn`, `Part-turn`, 또는 `Linear`
- RPM: Multi-turn만 입력, Part-turn/Linear는 0 또는 비워둠
- OpTime_sec: Part-turn만 입력, Multi-turn/Linear는 0 또는 비워둠
- Speed_mm_sec, Stroke_mm: Linear만 입력, Multi-turn/Part-turn은 0 또는 비워둠
- Torque_Nm: Multi-turn/Part-turn만 입력, Linear는 0 또는 비워둠
- Thrust_kN: Multi-turn/Linear만 입력, Part-turn은 0 또는 비워둠

### DB_Gearboxes (삼보 기어박스)

삼보산업(Sambo) 기어박스 데이터. 순수 기계 부품으로 시리즈/전원 옵션 없음.

**3가지 타입:**

| 타입 | 시리즈 | 특징 | Efficiency |
|------|--------|------|------------|
| Bevel Gear | SB-V, SB-VS | Part-turn/Multi-turn, 고효율 | 0.90~0.97 |
| Spur Gear | SB-SR | Multi-turn, 고비율 | 0.83~0.92 |
| Worm Gear | SBWG | Multi-turn, 초고비율, 셀프락킹 | 0.30~0.37 |

**DB 구조:**
```
| Model | Ratio | InputTorqueMax | OutputTorqueMax | Efficiency | InputFlange | OutputFlange | MaxStemDim_mm | Weight_kg | Price |
```

**대표 모델 예시:**
```
| SB-VS10 | 2.5 | 92.6 | 220 | 0.96 | F10 | F10 | 30 | 0 | 200 |
| SB-V3 | 5 | 526.3 | 2500 | 0.96 | F25 | F25 | 72 | 0 | 900 |
| SB-SR100 | 12 | 88.8 | 980 | 0.92 | F16 | F16 | 55 | 0 | 500 |
| SBWG-00 | 40 | 85.0 | 1200 | 0.35 | F12 | F12 | 36 | 0 | 550 |
```

**참고:**
- InputFlange = OutputFlange 가정 (실제 데이터 확인 시 수정 필요)
- Weight: 0 (수기 입력 필요)
- Price: 플레이스홀더 (실제 가격 입력 필요)
- Efficiency = Mechanical Advantage / Ratio (카탈로그 데이터)
- 상세 스펙: `docs/Sambo_Gearbox_Specifications.md`

## DB_Couplings 실제 데이터 입력

### 현재 구조

```
| CouplingType              | MinDimension_mm | MaxDimension_mm |
|---------------------------|-----------------|-----------------|
| Thrust Base - Threaded    | 20              | 120             |
| Standard (Part-turn)      | 0               | 0               |
```

### 주의사항

- **문자열 정확히 일치**: Settings 드롭다운, ValveList 드롭다운과 완전 일치 필요
- **0, 0 = 검증 스킵**: 치수 검증 불필요한 경우 (예: Part-turn 직접 플랜지)

### 새 커플링 타입 추가 요청 시

다음 정보 제공:
```
추가할 커플링: "Rising Stem - Flanged"
- MinDimension: 30mm
- MaxDimension: 200mm
```

Claude가 업데이트할 파일:
- `create_workbook.py` - DB_Couplings 데이터 + Settings/ValveList 드롭다운

## Error Handling Process (오류 처리 프로세스)

모든 오류 발생 시 다음 단계를 따른다.

### 1단계: 오류 기록

| 항목 | 내용 |
|------|------|
| 발생 일시 | YYYY-MM-DD |
| 오류 유형 | 데이터 모델링 / 코드 로직 / VBA / Python / 기타 |
| 증상 | 무엇이 잘못되었는가 |
| 관측 데이터 | 어떻게 발견했는가 (사용자 피드백, 테스트 결과 등) |

### 2단계: 근본 원인 분석

- **왜 발생했는가?** (표면적 원인)
- **어떤 가정이 잘못되었는가?** (근본 원인)
- **어떤 검증 단계가 누락되었는가?** (프로세스 결함)

### 3단계: 수정 및 검증

- Before/After 비교 명시
- 수정된 파일 목록
- 테스트 방법 및 결과

### 4단계: 재발 방지

- 체크리스트에 항목 추가
- CLAUDE.md에 사례 기록
- 필요시 검증 자동화

---

## Error Case Log (오류 사례 기록)

### [2025-12-27] 데이터 모델링: SA 시리즈 EXP 분류 오류

**오류 유형**: 데이터 모델링

**증상**:
- SA 시리즈의 "EXP"를 `ControlType` 컬럼에 분류
- 실제로는 `Enclosure` 옵션 (방폭형 = Exd)

**관측 데이터**:
- 사용자 피드백: "EXP PCU 이런건 control type이 아닌데?"
- 카탈로그 재검토 결과:
  - SA05X, SA09X: 모델명에 'X' 접미사 (방폭형)
  - SA005L, SA009L: 모델명에 'L' 접미사 (비례제어)

**근본 원인 분석**:
1. **표면적 유사성**: 카탈로그 Type 컬럼에 `ON-OFF, PCU, SCP, EXP`가 함께 나열
2. **도메인 지식 부족**: EXP = Explosionproof = Enclosure 옵션이라는 것 미인지
3. **검증 단계 누락**: 같은 컬럼에 있어도 실제 의미가 다를 수 있음

**수정 내용**:
```
Before: ControlType = [ONOFF, PCU, SCP, EXP]
After:  ControlType = [ONOFF, PCU, SCP]  # 제어 방식
        Enclosure = [IP67, Exd]           # 보호등급 (EXP = Exd)
```

**수정된 파일**:
- `create_workbook.py`: SA 모델 데이터, Settings, PowerOptions, EnclosureOptions, ElectricalData
- `CLAUDE.md`: 문서 업데이트

**재발 방지**: 데이터 모델링 검증 체크리스트 추가 (아래)

---

### [2025-12-27] VBA 수정: 플랫 DB 구조 지원

**오류 유형**: 코드 로직

**증상**:
- VBA 코드가 기존 10컬럼 DB_Models 구조 기준
- 새로운 15컬럼 플랫 구조와 불일치

**수정 내용**:

1. **modHelpers.bas - ModelRecord 타입**
```vba
' Before: 10 fields
' After: 15 fields (추가: MotorPower_kW, ControlType, Freq, OpTime, DutyCycle)
```

2. **modHelpers.bas - ActuatorRecord 타입**
```vba
' 추가: OpTime As Double  ' Part-turn용 (DB에서 직접 읽음)
```

3. **modHelpers.bas - ReadModelRecord()**
```vba
' 15개 컬럼 읽기로 수정
' 1:Model, 2:Series, 3:ActType, 4:MotorPower_kW, 5:ControlType,
' 6:Freq, 7:RPM, 8:Torque, 9:Thrust, 10:OpTime,
' 11:DutyCycle, 12:OutputFlange, 13:MaxStemDim, 14:Weight, 15:BasePrice
```

4. **modHelpers.bas - PassesModelFilters()**
```vba
' Freq 필터 추가
If m.Freq <> s.Frequency Then Exit Function
```

5. **modHelpers.bas - ResolveActuator()**
```vba
' OpTime 필드 전달 추가
.OpTime = m.OpTime
```

6. **modHelpers.bas - CalculateOpTime()**
```vba
' 시그니처 변경: Optional actOpTime As Double = 0 추가
' Part-turn: actOpTime > 0이면 actOpTime * gbRatio 반환
```

7. **modSizing.bas, modMain.bas**
```vba
' CalculateOpTime 호출 시 act.OpTime 전달
calcOpTime = CalculateOpTime(act.RPM, reqTurns, s.ActuatorType, gbRatio, act.OpTime)
```

**수정된 파일**:
- `vba/modHelpers.bas`: 타입 정의, ReadModelRecord, PassesModelFilters, ResolveActuator, CalculateOpTime
- `vba/modSizing.bas`: FindBestActuator, FindActuatorWithGearbox
- `vba/modMain.bas`: ShowAlternatives

---

### [2025-12-27] 데이터 모델링: MS 시리즈 모델명 일관성 개선

**오류 유형**: 데이터 모델링

**증상**:
- MS 시리즈 모델명에 Phase 접미사 포함 (`MS01-3P`, `MS01-1P`)
- 다른 시리즈는 접미사 없음 (`NA006`, `SA005`)
- 데이터 입력 시 일관성 부족, 유지보수 어려움

**관측 데이터**:
- 사용자 피드백: "일관성을 위해 표시하는거 수정되어야 하지 않아?"
- 유지보수 관점: 직접 DB 시트에 데이터 추가 시 규칙 혼란

**근본 원인 분석**:
1. **MS 시리즈 특수성**: Phase에 따라 토크가 다름 (3상: 110Nm, 1상: 45Nm)
2. **초기 설계 결정**: 별도 제품으로 처리하기 위해 모델명에 Phase 포함
3. **일관성 부족**: 다른 시리즈와 네이밍 규칙 불일치

**수정 내용**:
```
Before: Model = "MS01-3P", "MS01-1P" (17 컬럼)
After:  Model = "MS01", Phase = 3 또는 1 (18 컬럼)
```

**변경 사항**:
1. DB_Models에 Phase 컬럼 추가 (6번째 컬럼)
2. MS 시리즈: Model = "MS01", Phase = 3 또는 1
3. 다른 시리즈: Phase = 0 (Phase가 토크에 영향 없음)
4. PassesModelFilters: Phase > 0이면 Settings.Phase와 매칭 필터 추가

**수정된 파일**:
- `create_workbook.py`: DB_Models 18컬럼, MS 모델명 변경, PowerOptions/EnclosureOptions/ElectricalData 업데이트
- `vba/modHelpers.bas`: ModelRecord에 Phase 필드 추가, ReadModelRecord/PassesModelFilters 수정, 컬럼 인덱스 업데이트
- `CLAUDE.md`: 문서 업데이트

**재발 방지**: 모델명 네이밍 시 일관성 우선 고려

---

### [2025-12-28] DB 조회: MA 시리즈 Weight 조회 실패

**오류 유형**: 코드 로직 / DB 구조 불일치

**증상**:
- MA03 (MotorPower: 2.2kW) 선정 후 Datasheet 출력 시 Weight가 표시되지 않음
- DB_Models에는 Weight 데이터가 분명히 존재

**관측 데이터**:
- 사용자 피드백: "MA03, MotorPower: 2.2인데 무게는 분명 DB에 있는데, datasheet에는 안따라와"
- DB 확인: MA03은 kW별로 여러 행 존재 (2.2, 3.7, 5.5, 7.5kW)

**근본 원인 분석**:
1. **플랫 DB 구조**: MA 시리즈는 같은 Model명에 여러 MotorPower_kW 옵션 존재
   - MA03 + 2.2kW → Weight = 106kg
   - MA03 + 7.5kW → Weight = 132kg
2. **조회 함수 불완전**: `GetActuatorWeightByModel(actModel)` 함수가 Model명만으로 검색
3. **첫 번째 매칭 반환**: 동일 Model의 첫 행 Weight를 반환 (kW 무관)
4. **kW 정보 미활용**: ValveList에 COL_KW(23)로 kW가 저장되어 있지만 전달되지 않음

**수정 내용**:
```vba
' Before: Model만으로 조회
actWeight = GetActuatorWeightByModel(actModel)

' After: Model + kW로 조회
motorKW = GetCellDouble(wsValve.Cells(valveRow, COL_KW))
actWeight = GetActuatorWeightByModel(actModel, motorKW)

' GetActuatorWeightByModel 함수 시그니처 변경
Public Function GetActuatorWeightByModel(actModel As String, Optional motorKW As Double = 0) As Double
    ' motorKW > 0이면 Model + MotorPower_kW 동시 매칭
```

**수정된 파일**:
- `vba/modHelpers.bas`: GetActuatorWeightByModel에 Optional motorKW 파라미터 추가
- `vba/modDatasheet.bas`: FillDatasheetLine에서 COL_KW 읽어 전달

**재발 방지 체크리스트**:
- [ ] 플랫 DB에서 같은 Model이 여러 행으로 존재하는 경우, 조회 시 추가 키 필요
- [ ] MA 시리즈: Model + MotorPower_kW
- [ ] 기타 시리즈: Model만으로 충분 (kW 없거나 고정)

**추가 고려사항**:
- Datasheet "Equipment Offered" 섹션에 MotorPower(kW) 표시 여부
- 현재 Template에는 해당 행 없음 → 필요시 Template_Datasheet 수정 필요

---

## Validation Checklists (검증 체크리스트)

### 데이터 모델링 검증

새로운 제품 데이터 입력 시 반드시 확인:

1. **분류 기준 확인**
   - [ ] 같은 컬럼에 나열된 항목들이 정말 같은 카테고리인가?
   - [ ] 각 항목의 실제 의미/기능이 무엇인가?

2. **모델명 패턴 분석**
   - [ ] 접미사/접두사에 의미가 있는가?
   - [ ] 예: 'L' = PCU/SCP (비례제어), 'X' = 방폭형

3. **직교성(Orthogonality) 검증**
   - [ ] 한 옵션이 다른 옵션과 독립적인가?
   - [ ] 예: ControlType(ONOFF/PCU)과 Enclosure(IP67/Exd)는 독립적으로 선택 가능해야 함
   - [ ] 반례: EXP 모델은 Exd만 지원 → 완전 독립은 아님 (제약조건)

4. **가격 구조 검증**
   - [ ] 옵션별 가격 추가가 어디서 발생하는가?
   - [ ] 중복 계산되지 않는가?

5. **사용자 검토 요청**
   - [ ] 도메인 전문가(사용자)에게 분류 확인 요청
   - [ ] "이 항목들이 같은 카테고리가 맞나요?" 질문

### 코드 변경 검증

- [ ] 변경된 함수의 호출처 모두 확인
- [ ] 관련 상수/타입 정의 동기화
- [ ] 에러 핸들링 유지

### VBA 모듈 검증

- [ ] 타입 정의 변경 시 모든 모듈에서 사용처 확인
- [ ] Excel 시트 컬럼 인덱스 동기화
- [ ] Application 상태 복원 (ScreenUpdating, EnableEvents)

---

## Domain Knowledge Reference (도메인 지식)

### 액추에이터 옵션 분류 원칙

| 카테고리 | 예시 | 특징 | 핵심 질문 |
|----------|------|------|-----------|
| 제어 방식 (Control) | ONOFF, PCU, SCP | 동작 방식, 소프트웨어적 | 어떻게 움직이는가? |
| 보호등급 (Enclosure) | IP67, Exd | 물리적 하우징, 인증 | 어디서 사용 가능한가? |
| 전원 (Power) | 380V/3상/50Hz | 전기적 사양 | 무슨 전원이 필요한가? |
| 기계적 (Mechanical) | Flange, Torque, RPM | 물리적 성능 | 무엇을 움직일 수 있는가? |

### Noah 모델명 접미사 규칙

| 접미사 | 의미 | 예시 |
|--------|------|------|
| L | 비례제어 (PCU/SCP) | SA005L, SA009L |
| X | 방폭형 (Exd) | SA05X, SA09X |
| 숫자 | 토크 등급 (kg.m) | NA006 = 6 kg.m |
