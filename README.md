# Actuator Sizing Tool

전동 액추에이터 + 기어박스 모델 선정 프로그램

## 설치 방법

### 1단계: Excel 파일 준비

1. `NoahSizing.xlsx` 파일을 Excel에서 엽니다
2. **다른 이름으로 저장** → 파일 형식을 **Excel 매크로 사용 통합 문서 (*.xlsm)** 로 선택
3. `NoahSizing.xlsm`으로 저장

### 2단계: VBA 모듈 Import

1. Excel에서 **Alt + F11** 을 눌러 VBA 편집기 열기
2. 좌측 **프로젝트 탐색기**에서 `VBAProject (NoahSizing.xlsm)` 우클릭
3. **파일 가져오기(Import File)** 선택
4. `vba` 폴더에서 아래 파일들을 순서대로 가져오기:
   - `modHelpers.bas` ← **반드시 첫 번째로** (공통 타입/함수 정의)
   - `modSettings.bas`
   - `modSizing.bas`
   - `modMain.bas`
   - `modDatasheet.bas`
5. **UserForm 생성** (아래 2-1단계 참조)
6. VBA 편집기 닫기 (Alt + Q)
7. 파일 저장

> **주의**: `modHelpers.bas`를 먼저 가져와야 합니다. 다른 모듈들이 이 모듈의 타입과 함수를 참조합니다.

### 2-1단계: UserForm 생성 (필수)

Alternative 선택 시 사용되는 UserForm입니다.

1. VBA 편집기에서 **삽입(Insert)** → **UserForm** 선택
2. Properties 창에서 Name을 `frmAlternatives`로 변경
3. 폼 속성 설정:
   - Caption: `Select Alternative Model`
   - Width: `680`, Height: `440`
4. Toolbox에서 컨트롤 추가:

| 컨트롤 | Name | 설명 |
|--------|------|------|
| Label | `lblHeader` | 상단 헤더 (컬럼명 표시) |
| ListBox | `lstAlternatives` | 메인 목록 (ColumnCount: 9) |
| Label | `lblInfo` | 하단 정보 텍스트 |
| CommandButton | `btnOK` | OK 버튼 (Default: True) |
| CommandButton | `btnCancel` | Cancel 버튼 (Cancel: True) |

5. 폼을 더블클릭하여 **코드 창** 열기
6. `vba/frmAlternatives.frm` 파일을 메모장으로 열기
7. `Option Explicit`부터 파일 끝까지 복사
8. VBA 코드 창에 붙여넣기

> **주의**: `frmAlternatives.frm`은 직접 Import할 수 없습니다. 코드만 복사해서 사용하세요.

> **TIP**: 컨트롤 속성 상세는 [TECHNICAL_GUIDE.md](TECHNICAL_GUIDE.md)의 "5.5 frmAlternatives" 섹션 참조

### 3단계: 버튼 추가 (선택사항)

ValveList 시트에 버튼을 추가하여 매크로를 연결합니다:

1. **개발 도구** 탭 → **삽입** → **양식 컨트롤** → **단추**
2. 시트에 버튼 그리기
3. 매크로 지정 대화상자에서 해당 매크로 선택:

| 버튼 이름 | 매크로 | 설명 |
|----------|--------|------|
| Add Line | `btn_AddLine` | Settings 기준으로 라인 추가 (Coupling Type, ValveType 드롭다운 자동 설정) |
| Sizing All | `btn_SizingAll` | 모든 라인 사이징 |
| Sizing Selected | `btn_SizingSelected` | 선택한 라인만 사이징 |
| Alternative | `btn_Alternative` | 대체 모델 조회 |
| To Configuration | `btn_ToConfiguration` | 선정된 모델을 Configuration 시트로 복사 (옵션 선택용) |
| Export Datasheet | `btn_ExportDatasheet` | Datasheet 엑셀 출력 |
| Clear Results | `btn_ClearResults` | 결과 초기화 |

또는 **Alt + F8** 을 눌러 매크로 목록에서 직접 실행할 수 있습니다.

---

## 사용 방법

### 1. Settings 시트에서 기본 설정

- **Torque Unit**: Nm / lbf.ft / kgf.m
- **Thrust Unit**: kN / lbf / kgf
- **Enclosure**: Waterproof / Explosionproof
- **Safety Factor**: 안전율 (기본 1.25)
- **Actuator Type**: Multi-turn / Part-turn / Linear
  - Multi-turn: Gate, Globe 밸브용 (다회전)
  - Part-turn: Ball, Butterfly, Plug 밸브용 (90° 회전)
  - Linear: Linear 밸브용 (직선 운동)
- **Operation Mode**: On-Off / Modulating / Modulating (High-Speed)
  - On-Off: 기본 개폐 제어
  - Modulating: 비례 제어 (PCU)
  - Modulating (High-Speed): 고속 비례 제어 (SCP)
- **Fail-safe**: None / Close-on-Fail (SR)
  - SR 선택 시 Spring Return 시리즈만 검색
- **Duty Cycle**: Any / Intermittent (S2) / Continuous (S4)
- **Voltage/Phase/Freq**: 전원 사양
- **Op. Time Range**: 허용 Operating Time 범위 (%)
- **Coupling Type**: 커플링 타입 (공통 설정, 라인 추가시 자동 적용)
  - Multi-turn: Thrust Base - Threaded
  - Part-turn: Standard (Part-turn)
- **Model Range**: 모델 시리즈 필터 (All / NA / SA / SR / MA / MS / NL)
- **Lines to Add**: Add Line 실행 시 추가할 라인 수

### 2. ValveList 시트에서 밸브 정보 입력

1. **Add Line** 실행 → Settings에서 지정한 수만큼 라인 자동 생성
2. 각 라인의 밸브 정보 입력:

| 컬럼 | 설명 | 비고 |
|------|------|------|
| Line No. | 라인 번호 | 자동 생성 (수정 불필요) |
| Tag | 밸브 태그 | 사용자 입력 |
| ValveType | 밸브 종류 | Gate, Globe, Ball, Butterfly, Plug, Linear |
| Size | 밸브 사이즈 | |
| Class | 압력 등급 | |
| Torque | 필요 토크 | Multi-turn/Part-turn용 (Linear는 0) |
| Thrust | 필요 추력 | Multi-turn/Linear용 (Part-turn은 0) |
| CouplingType | 커플링 타입 | Settings에서 자동 적용, Part-turn은 "Standard" 선택 |
| CouplingDim | 커플링 치수 | mm (DB_Couplings 범위 검증) |
| Lift | 밸브 리프트 | mm, Multi-turn용 (Turns = Lift / Pitch) |
| Pitch | 스템 피치 | mm, Multi-turn용 |
| Op.Time | 요구 작동시간 | 초 |

> **ValveType → ActuatorType 자동 결정**: ValveType 선택 시 ActuatorType이 자동으로 결정됩니다.
> - Ball, Butterfly, Plug → Part-turn
> - Gate, Globe → Multi-turn
> - Linear → Linear

### 3. 사이징 실행

- **Sizing All**: 모든 라인 사이징
- **Sizing Selected**: 선택한 라인만 사이징
- 스펙 미충족 시 Status 컬럼에 사유 표시

### 4. Alternative 모델 확인

1. ValveList 시트에서 라인 선택 후 **Alternative** 실행
2. 스펙을 만족하는 모델 목록이 UserForm에 표시됨 (9개 컬럼 리스트박스)
3. 원하는 모델 선택 후 **OK** 클릭 (또는 더블클릭)
4. 선택하면 ValveList에 즉시 반영되고 Status는 `OK (Alternative)`로 표시됨
5. **Cancel**/닫기(X)는 변경 없음, 결과가 없으면 Status 컬럼에 사유가 표시됨

### 5. Configuration에서 옵션 선택

사이징 완료 후 옵션을 추가하고 최종 가격을 계산합니다:

1. **To Configuration** 실행 → 선정된 모델이 Configuration 시트로 복사됨
2. 각 라인별 옵션 선택:

| 컬럼 | 설명 | 입력 방식 |
|------|------|-----------|
| Line, Tag, Model, Gearbox | ValveList에서 자동 복사 | 자동 |
| Base | 모델 + 기어박스 가격 | 자동 |
| HTR | Space Heater ($50) | Yes/No 드롭다운 |
| MOD | Modulating Control ($200) | Yes/No 드롭다운 |
| POS | Position Transmitter ($150) | Yes/No 드롭다운 |
| LMT | Limit Switch ($80) | Yes/No 드롭다운 |
| EXD | Explosionproof Upgrade ($300) | Yes/No 드롭다운 |
| Painting | 도장 옵션 | None/EP/PU/SPEC 드롭다운 |
| Qty | 수량 | 수동 입력 (기본 1) |
| Unit | Base + 선택 옵션 합계 | 자동 계산 |
| Total | Unit × Qty | 자동 계산 |

3. 하단에 Grand Total 자동 합산

### 6. Datasheet 출력

- **Export Datasheet** 실행
- 저장 위치 선택
- 새 Excel 파일로 Datasheet 생성

---

## 파일 구조

```
actuatorSizing/
├── NoahSizing.xlsx            # 기본 Excel 파일 (생성됨)
├── NoahSizing.xlsm            # 매크로 포함 파일 (사용자가 변환)
├── create_workbook.py         # Excel 파일 생성 스크립트
├── vba/
│   ├── modHelpers.bas         # 공통 타입, 상수, 유틸리티 함수
│   ├── modSettings.bas        # 설정 로드/검증
│   ├── modSizing.bas          # 사이징 엔진
│   ├── modMain.bas            # 버튼 핸들러, Alternative 선택
│   ├── modDatasheet.bas       # Datasheet 엑셀 출력
│   └── frmAlternatives.frm    # Alternative 선택 UserForm (코드 참조용)
├── README.md                  # 이 문서 (사용자 가이드)
└── TECHNICAL_GUIDE.md         # 기술 문서 (코드 로직 설명)
```

> **참고**: 코드 로직, 알고리즘, 데이터 흐름에 대한 상세 설명은 [TECHNICAL_GUIDE.md](TECHNICAL_GUIDE.md)를 참조하세요.

---

## VBA 모듈 설명

| 모듈 | 설명 |
|------|------|
| `modHelpers.bas` | 공통 타입(`ModelRecord`, `ActuatorRecord`, `GearboxRecord`), 상수, 유틸리티 함수, DB 조회(`ReadModelRecord`, `ResolveActuator`) |
| `modSettings.bas` | Settings 시트에서 설정 로드 및 검증 |
| `modSizing.bas` | 사이징 알고리즘 (직접 선정 + 기어박스 조합) |
| `modMain.bas` | 버튼 핸들러, Alternative 조회 및 선택 처리 |
| `modDatasheet.bas` | Datasheet 엑셀 파일 출력 |
| `frmAlternatives.frm` | Alternative 선택 UserForm 코드 (수동 생성 필요) |

---

## DB 시트 구조 (플랫 + 옵션 테이블)

액추에이터 DB는 **플랫 구조**를 사용합니다. 각 Model × Freq × kW/RPM 조합이 별도 행으로 등록되며, 전원/Enclosure 옵션은 별도 테이블에서 관리됩니다.

### DB_Models (기본 사양 - 플랫 구조, 17컬럼)
| Model | Series | ActType | MotorPower_kW | ControlType | Freq | RPM | Torque_Nm | Thrust_kN | OpTime_sec | DutyCycle | OutputFlange | MaxStemDim_mm | Weight_kg | BasePrice | Speed_mm_sec | Stroke_mm |

> **참고**:
> - **플랫 구조**: 각 Model × Freq × kW/RPM 조합이 별도 행으로 등록
> - Multi-turn: RPM 사용, Torque + Thrust
> - Part-turn: OpTime_sec 사용 (90° 동작시간), Torque만
> - Linear: Speed_mm_sec + Stroke_mm 사용, Thrust만
> - `MaxStemDim_mm`은 액추에이터가 수용 가능한 최대 밸브 스템 직경(mm)

### DB_PowerOptions (전원 옵션)
| Model | Voltage | Phase | Freq | PriceAdder |

> **참고**: 모델별 지원 전원 조합. PriceAdder는 기본가 대비 추가/감소 금액

### DB_EnclosureOptions (보호등급 옵션)
| Model | Enclosure | PriceAdder |

> **참고**: 모델별 지원 Enclosure. IP67→Waterproof, Exd→Explosionproof 매칭

### DB_ElectricalData (Datasheet용 전기 데이터)
| Model | Voltage | Phase | Freq | StartingCurrent_A | StartingPF | RatedCurrent_A | AvgCurrent_A | AvgPF | AvgPower_kW | MotorPoles |

### DB_Gearboxes (삼보 기어박스)
| Model | Ratio | InputTorqueMax | OutputTorqueMax | Efficiency | InputFlange | OutputFlange | MaxStemDim_mm | Weight_kg | Price |

> **참고**:
> - 삼보산업(Sambo) 기어박스 데이터 포함
> - **Bevel Gear (SB-V)**: 고효율 (0.90~0.97)
> - **Spur Gear (SB-SR)**: Multi-turn용, 고비율
> - **Worm Gear (SBWG)**: 초고비율, 셀프락킹, 저효율 (0.30~0.37)
> - `MaxStemDim_mm`은 기어박스 출력부가 수용 가능한 최대 밸브 스템 직경(mm)
> - Linear 액추에이터는 기어박스 미사용 (직접 연결만)

### DB_Couplings
| CouplingType | MinDimension_mm | MaxDimension_mm |

### DB_Options
| Code | Description | Price |

### Configuration (옵션 선택 시트)
| Line | Tag | Model | Gearbox | Base | HTR | MOD | POS | LMT | EXD | Painting | Qty | Unit | Total |

- HTR~EXD: Yes/No 드롭다운
- Painting: None/PAINT-EP/PAINT-PU/PAINT-SPEC 드롭다운
- Unit, Total: 수식 자동 계산
- **가격 계산**: BasePrice + PowerAdder + EnclosureAdder

---

## 실제 데이터 입력

가상 데이터를 실제 Noah 제품 데이터로 교체하려면:

1. 시트 숨기기 해제 (시트 탭 우클릭 → 숨기기 취소)
2. **DB_Models**: 각 모델 기본 사양 입력 (모델당 1행)
3. **DB_PowerOptions**: 모델별 지원 전원 조합 입력
4. **DB_EnclosureOptions**: 모델별 지원 Enclosure 입력
5. **DB_ElectricalData**: Datasheet용 전기 데이터 입력
6. **DB_Gearboxes**: 기어박스 스펙 입력
7. **DB_Couplings**: 커플링 타입별 치수 범위 입력
8. **DB_Options**: 옵션 코드 및 가격 입력
9. 시트 다시 숨기기

> **주의**: 숫자는 숫자 형식으로 입력 (텍스트 "100" ❌ → 숫자 100 ✅)

---

## 사이징 공식

> 참조: DMRA "Formulas to help when quoting"

### Multi-turn (다회전)

**회전수 계산**
```
Turns = Lift / Pitch
```
- Lift: 밸브 스템 이동 거리 (mm)
- Pitch: 스템 나사 피치 (mm)

**직접 구동 (액추에이터만)**
```
Operating Time (sec) = (Turns × 60) / RPM
```

**기어박스 조합**
```
Operating Time (sec) = (Turns × Ratio × 60) / Actuator RPM
출력 토크 = 액추에이터 토크 × 기어비 × 효율
```

### Part-turn (1/4 Turn = 90°)

**직접 구동 (액추에이터만)**
```
Operating Time (sec) = DB의 OpTime_sec 직접 사용 (90° 동작시간)
```

**기어박스 조합**
```
Operating Time (sec) = OpTime_sec × Ratio
출력 토크 = 액추에이터 토크 × 기어비 × 효율
```

> **참고**: Part-turn 액추에이터는 DB에 90° 동작시간이 직접 저장되어 있음

### Linear (직선 운동)

**직접 구동만 (기어박스 미사용)**
```
Operating Time (sec) = Stroke_mm / Speed_mm_sec
```

> **참고**: Linear 액추에이터는 Torque 대신 Thrust(추력)를 사용하며, 기어박스를 사용하지 않음

### 토크/추력 계산
```
필요 토크 = 밸브 토크 × 안전율
필요 추력 = 밸브 추력 × 안전율
```

### 회전수 계산 참고 (Multi-turn)
```
Turns = Lift (mm) / Pitch (mm)
```
- 입력: Lift, Pitch → 프로그램이 Turns 자동 계산
- 예: Lift=100mm, Pitch=5mm → Turns=20

### 커플링 검증
- CouplingDim은 DB_Couplings 범위 내여야 함 (범위 미충족 시 Status 표시)
- **MaxStemDim 검증**:
  - 직접 연결 시: CouplingDim ≤ Actuator.MaxStemDim
  - 기어박스 사용 시: CouplingDim ≤ Gearbox.MaxStemDim

---

## Best Practice 체크리스트

- 단위/안전율 설정 후 입력값 확인
- Multi-turn은 Torque + Thrust 모두 만족해야 함
- Op. Time 범위(%) 검증
- Gearbox 조합 시 InputTorqueMax/OutputTorqueMax/효율/플랜지 호환 확인
- CouplingDim 범위(DB_Couplings) 확인
- **MaxStemDim 확인**: 액추에이터/기어박스가 밸브 스템 치수를 수용 가능한지 확인
- 미충족 사유는 Status 컬럼에서 확인

## 버전 정보

- **v3.0**: 전체 시리즈 지원 및 삼보 기어박스 (2025.12)
  - **Linear 액추에이터 지원**: NL 시리즈 (직선 운동, Thrust 기반)
  - **Spring Return 지원**: SR 시리즈 (Fail-safe 옵션)
  - **삼보 기어박스 통합**: Bevel (SB-V), Spur (SB-SR), Worm (SBWG)
  - **DB_Models 17컬럼 확장**: MotorPower_kW, ControlType, Freq, OpTime_sec, Speed_mm_sec, Stroke_mm 추가
  - **Part-turn OpTime**: DB에서 직접 읽기 (계산 대신)
  - **Operation Mode 확장**: Modulating (High-Speed) 추가
  - **필터 추가**: Fail-safe, Duty Cycle
  - **Datasheet Weight 출력**: 액추에이터/기어박스/조합 무게 표시

- **v2.0**: 플랫 DB 구조 + 옵션 테이블 적용 (2025.12)
  - DB_Actuators → DB_Models (플랫 구조: 같은 Model이 Freq/kW/RPM 조합별로 여러 행)
  - 옵션 테이블 분리 (DB_PowerOptions, DB_EnclosureOptions, DB_ElectricalData)
  - 가격 계산: BasePrice + PowerAdder + EnclosureAdder

- **v1.7**: MaxStemDim 검증 추가 (2024.12)
  - DB_Models에 `MaxStemDim_mm` 컬럼 추가 (직접 연결 시 밸브 스템 치수 검증)
  - DB_Gearboxes에 `MaxStemDim_mm` 컬럼 추가 (기어박스 사용 시 밸브 스템 치수 검증)
  - 사이징 시 CouplingDim ≤ MaxStemDim 자동 검증

- **v1.6**: Alternative 선택 UserForm 추가 (2024.12)
  - Alternative 선택 시 UserForm UI 사용 (`frmAlternatives`)
  - 9개 컬럼 리스트박스로 모델 비교/선택 가능

- **v1.5**: CalcThrust 컬럼 추가 및 UI 개선 (2024.12)
  - ValveList 결과에 CalcThrust (계산된 추력) 컬럼 추가 (Multi-turn용)
  - 입력/결과 컬럼 색상 구분 (파란색: 입력, 녹색: 결과)
  - TECHNICAL_GUIDE.md 기술 문서 추가

- **v1.4**: Configuration 및 공식 개선 (2024.12)
  - Configuration 시트 기능 구현: 옵션별 컬럼 (HTR, MOD, POS, LMT, EXD, Painting)
  - `btn_ToConfiguration`: 선정된 모델을 Configuration으로 복사
  - Unit Price / Total Price 자동 계산 (수식)
  - Operating Time 공식 업데이트 (DMRA 참조 공식 반영)
  - Part-turn + Gearbox 조합 공식 추가: `(Ratio × 60) / (4 × RPM)`

- **v1.3**: 검증/Alternative 개선 (2024.12)
  - Alternative는 스펙 충족 모델만 표시, 미충족 사유 Status 표시
  - Gearbox 조합에서 Thrust/Op Time 검증 강화
  - CouplingDim 범위 검증 (DB_Couplings)
  - Datasheet에 Calculated Thrust/안전율 표시

- **v1.2**: 워크플로우 개선 (2024.12)
  - Settings에 Coupling Type 추가 (공통 설정)
  - Settings에 Lines to Add 추가 (일괄 라인 생성)
  - Add Line 기능 개선: 지정 수만큼 라인 생성, 공통 설정 자동 적용

- **v1.1**: 코드 개선 (2024.12)
  - 에러 처리 추가 (시트 누락, 런타임 오류 대응)
  - 진행률 표시 (상태바에 작업 진행 상황 표시)
  - 코드 중복 제거 (공통 함수 `modHelpers.bas`로 통합)
  - Operating Time 범위 검증 로직 개선
  - 기본값 자동 설정 (Safety Factor, 단위 등)

- **v1.0**: 초기 버전 (2024.12)
  - 사이징 엔진
  - Alternative 모델 조회
  - Datasheet 엑셀 출력

---

## 문제 해결

### "User-defined type not defined" 오류
- `modHelpers.bas`가 먼저 Import 되었는지 확인
- VBA 편집기에서 **디버그** → **컴파일** 실행

### "Subscript out of range" 오류
- 필요한 시트(Settings, ValveList, DB_Actuators 등)가 존재하는지 확인
- 시트 이름이 정확한지 확인 (대소문자 구분)

### 사이징 결과가 나오지 않을 때
1. Settings 시트의 설정값 확인 (Voltage, Phase, Freq)
2. DB_Actuators에 해당 조건의 액추에이터가 있는지 확인
3. 요구 토크가 모든 액추에이터 용량을 초과하는지 확인
4. Status 컬럼의 사유 메시지 확인 (Op Time 범위/Enclosure/DB_Couplings 등)

### 진행률이 표시되지 않을 때
- Excel 상태바가 보이는지 확인 (하단)
- 다른 매크로가 상태바를 사용 중인지 확인

### Alternative 버튼 클릭 시 오류가 발생할 때
- UserForm(`frmAlternatives`)이 생성되지 않았습니다
- VBA 편집기에서 UserForm을 만들고 이름을 `frmAlternatives`로 정확히 설정하세요
- 2-1단계의 UserForm 생성 가이드를 참조하세요
