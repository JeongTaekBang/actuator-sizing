# 데이터 입력 가이드 (Data Entry Guide)

이 문서는 Noah Actuator Sizing Tool의 DB 시트에 실제 제품 데이터를 입력하는 담당자를 위한 가이드입니다.

---

## 1. DB 시트 간 관계도

```
┌─────────────────────────────────────────────────────────────────────────┐
│                              DB_Models                                   │
│  (메인 테이블 - 액추에이터 기본 사양)                                      │
│  ┌─────────────────────────────────────────────────────────────────┐    │
│  │ Model │ Series │ ActType │ ... │ Torque │ OutputFlange │ Price │    │
│  └───┬───────────────────────────────────────────────────────┬─────┘    │
└──────┼───────────────────────────────────────────────────────┼──────────┘
       │ Model (1:N)                                           │
       ▼                                                       │
┌──────────────────┐  ┌──────────────────┐  ┌─────────────────┐│
│ DB_PowerOptions  │  │DB_EnclosureOptions│  │DB_ElectricalData││
│ (전원 옵션)       │  │ (보호등급 옵션)    │  │ (Datasheet용)   ││
├──────────────────┤  ├──────────────────┤  ├─────────────────┤│
│ Model            │  │ Model            │  │ Model           ││
│ Voltage          │  │ Enclosure        │  │ Voltage         ││
│ Phase            │  │ PriceAdder       │  │ Phase           ││
│ Freq             │  │                  │  │ Freq            ││
│ PriceAdder       │  │                  │  │ StartingCurrent ││
└──────────────────┘  └──────────────────┘  │ ...             ││
                                            └─────────────────┘│
                                                               │
       ┌───────────────────────────────────────────────────────┘
       │ OutputFlange ↔ InputFlange (호환성 체크)
       ▼
┌─────────────────────────────────────────────────────────────────────────┐
│                            DB_Gearboxes                                  │
│  (기어박스 - 독립 테이블)                                                 │
│  ┌─────────────────────────────────────────────────────────────────┐    │
│  │ Model │ Ratio │ InputFlange │ OutputFlange │ MaxStemDim │ Price │    │
│  └─────────────────────────────────────────────────────────────────┘    │
└─────────────────────────────────────────────────────────────────────────┘

┌──────────────────┐  ┌──────────────────┐
│   DB_Couplings   │  │    DB_Options    │
│ (커플링 치수)     │  │ (추가 옵션)       │
├──────────────────┤  ├──────────────────┤
│ CouplingType     │  │ Code             │
│ MinDimension_mm  │  │ Description      │
│ MaxDimension_mm  │  │ Price            │
└──────────────────┘  └──────────────────┘
```

---

## 2. 사이징 로직에서 데이터 사용 흐름

### Step 1: DB_Models 필터링
```
사용자 입력 (Settings + ValveList)
    ↓
DB_Models 검색
    ├── ActType 매칭 (ValveType → Multi-turn/Part-turn/Linear)
    ├── Series 매칭 (Model Range 설정)
    ├── Freq 매칭 (Settings: 50Hz/60Hz)
    ├── Phase 매칭 (MS 시리즈만: 1상/3상)
    ├── Torque ≥ 요구토크 × 안전율
    ├── Thrust ≥ 요구추력 × 안전율 (Multi-turn, Linear만)
    └── MaxStemDim ≥ CouplingDim
```

### Step 2: 전원/Enclosure 옵션 확인
```
DB_Models에서 후보 모델 발견
    ↓
DB_PowerOptions 조회
    └── Model + Voltage + Phase + Freq 조합 존재 여부
    ↓
DB_EnclosureOptions 조회
    └── Model + Enclosure 조합 존재 여부
    ↓
없으면 → 해당 모델 제외
있으면 → 가격 계산 진행
```

### Step 3: 가격 계산
```
최종 가격 = BasePrice (DB_Models)
          + PowerAdder (DB_PowerOptions)
          + EnclosureAdder (DB_EnclosureOptions)
```

### Step 4: 기어박스 조합 (필요시)
```
직접 연결로 조건 만족 못할 경우
    ↓
DB_Gearboxes 검색
    ├── InputFlange = Actuator.OutputFlange
    ├── InputTorqueMax ≥ Actuator.Torque
    ├── OutputTorqueMax ≥ 요구토크 × 안전율
    ├── MaxStemDim ≥ CouplingDim
    └── 출력토크 = Actuator.Torque × Ratio × Efficiency
    ↓
조합 가격 = Actuator.Price + Gearbox.Price
```

---

## 3. 시트별 입력 가이드

### 3.1 DB_Models (메인 테이블)

**18개 컬럼 - 모두 필수 입력 (해당 없으면 0 또는 빈값)**

| 컬럼 | 설명 | 입력 규칙 | 예시 |
|------|------|----------|------|
| Model | 모델명 | 카탈로그 원본 그대로 | `NA006`, `MS01`, `MA01` |
| Series | 시리즈 | NA, SA, MA, MS, SR, NL 중 하나 | `NA` |
| ActType | 액추에이터 타입 | `Multi-turn`, `Part-turn`, `Linear` 정확히 | `Part-turn` |
| MotorPower_kW | 모터 출력 | MA 시리즈만 입력, 나머지 0 | `0.2` |
| ControlType | 제어 방식 | SA: ONOFF/PCU/SCP, SR: SR, 나머지 빈값 | `ONOFF` |
| Phase | 전원 Phase | MS만 1 또는 3, 나머지 0 | `0` |
| Freq | 주파수 | 50 또는 60 | `50` |
| RPM | 회전속도 | Multi-turn만, Part-turn/Linear는 0 | `16` |
| Torque_Nm | 토크 | Multi-turn/Part-turn, Linear는 0 | `60` |
| Thrust_kN | 추력 | Multi-turn/Linear만, Part-turn은 0 | `50` |
| OpTime_sec | 동작시간 | Part-turn만 (90° 시간), 나머지 0 | `18` |
| DutyCycle | 사용률 | S2-30min, S4-25% 등 | `S4-50%` |
| OutputFlange | 출력 플랜지 | F07, F10, F14 등 (Linear는 빈값) | `F07` |
| MaxStemDim_mm | 최대 스템 직경 | mm 단위 | `22` |
| Weight_kg | 무게 | kg 단위 | `11` |
| BasePrice | 기본 가격 | 숫자 (통화 단위 없이) | `500` |
| Speed_mm_sec | 이동 속도 | Linear만, 나머지 0 | `0.8` |
| Stroke_mm | 스트로크 | Linear만, 나머지 0 | `40` |

#### 시리즈별 필수 입력 컬럼

| 시리즈 | 필수 컬럼 | 0 또는 빈값 |
|--------|----------|-------------|
| **MA** | MotorPower_kW, RPM, Torque, Thrust, OutputFlange | Phase=0, OpTime=0, Speed=0, Stroke=0 |
| **MS** | Phase(1/3), RPM, Torque, Thrust, OutputFlange | MotorPower=0, OpTime=0, Speed=0, Stroke=0 |
| **NA** | Torque, OpTime_sec, OutputFlange | Phase=0, RPM=0, Thrust=0, Speed=0, Stroke=0 |
| **SA** | ControlType, Torque, OpTime_sec, OutputFlange | Phase=0, RPM=0, Thrust=0, Speed=0, Stroke=0 |
| **SR** | ControlType=SR, Torque, OpTime_sec, OutputFlange | Phase=0, RPM=0, Thrust=0, Speed=0, Stroke=0 |
| **NL** | Thrust, Speed_mm_sec, Stroke_mm | Phase=0, RPM=0, Torque=0, OpTime=0, OutputFlange=빈값 |

#### 같은 모델의 복수 행 등록

같은 모델이 여러 조건으로 존재하면 **각각 별도 행**으로 등록:

```
예: NA006 (50Hz, 60Hz 둘 다 지원)
┌────────┬────────┬───────────┬──────┬────────┬───────┬──────────┐
│ Model  │ Series │ ActType   │ Freq │ Torque │OpTime │ ...      │
├────────┼────────┼───────────┼──────┼────────┼───────┼──────────┤
│ NA006  │ NA     │ Part-turn │ 50   │ 60     │ 18    │ ...      │
│ NA006  │ NA     │ Part-turn │ 60   │ 60     │ 16    │ ...      │
└────────┴────────┴───────────┴──────┴────────┴───────┴──────────┘

예: MA01 (같은 Hz에서 kW/RPM 조합이 여러 개)
┌────────┬──────┬──────────────┬──────┬─────┬────────┐
│ Model  │ Freq │ MotorPower_kW│ RPM  │Torque│ ...   │
├────────┼──────┼──────────────┼──────┼──────┼───────┤
│ MA01   │ 50   │ 0.2          │ 16   │ 88   │ ...   │
│ MA01   │ 50   │ 0.2          │ 20   │ 88   │ ...   │
│ MA01   │ 50   │ 0.4          │ 16   │ 138  │ ...   │
│ MA01   │ 60   │ 0.2          │ 19   │ 88   │ ...   │
└────────┴──────┴──────────────┴──────┴──────┴───────┘
```

---

### 3.2 DB_PowerOptions (전원 옵션)

**DB_Models의 각 모델이 지원하는 전원 조합 등록**

| 컬럼 | 설명 | 예시 |
|------|------|------|
| Model | DB_Models.Model과 정확히 일치 | `NA006` |
| Voltage | 전압 (V) | `380`, `220`, `24` |
| Phase | 상수 (1, 3) 또는 DC | `3`, `1`, `DC` |
| Freq | 주파수 (Hz), DC면 0 | `50`, `60`, `0` |
| PriceAdder | 추가 가격 (기본 옵션은 0) | `0`, `50` |

#### 입력 규칙

1. **DB_Models에 있는 모든 Model에 대해** 최소 1개 이상의 전원 옵션 등록 필수
2. Model 값은 DB_Models의 Model과 **정확히 일치**해야 함 (대소문자, 공백 주의)
3. 기본 옵션 (가장 흔한 전원)은 PriceAdder = 0

```
예: NA006이 지원하는 전원 옵션들
┌────────┬─────────┬───────┬──────┬────────────┐
│ Model  │ Voltage │ Phase │ Freq │ PriceAdder │
├────────┼─────────┼───────┼──────┼────────────┤
│ NA006  │ 380     │ 3     │ 50   │ 0          │  ← 기본
│ NA006  │ 380     │ 3     │ 60   │ 0          │
│ NA006  │ 220     │ 1     │ 50   │ 0          │
│ NA006  │ 220     │ 1     │ 60   │ 0          │
│ NA006  │ 24      │ DC    │ 0    │ 100        │  ← DC 옵션 추가비
└────────┴─────────┴───────┴──────┴────────────┘
```

#### ⚠️ 누락 시 문제

**DB_PowerOptions에 없는 조합은 사이징에서 선택되지 않음**

예: NA006이 DB_Models에 있어도, DB_PowerOptions에 `NA006 + 440V + 3ph + 60Hz` 행이 없으면:
- Settings에서 440V/3상/60Hz 선택 시 NA006은 후보에서 제외됨

---

### 3.3 DB_EnclosureOptions (보호등급 옵션)

**DB_Models의 각 모델이 지원하는 Enclosure 등록**

| 컬럼 | 설명 | 예시 |
|------|------|------|
| Model | DB_Models.Model과 정확히 일치 | `NA006` |
| Enclosure | 보호등급 | `IP67`, `IP68`, `Exd`, `Exde` |
| PriceAdder | 추가 가격 | `0`, `300` |

#### Enclosure 매칭 규칙

Settings의 Enclosure 설정과 DB 값 매칭:

| Settings 값 | DB에서 찾는 패턴 | 예시 |
|-------------|-----------------|------|
| Waterproof | "IP" 포함 | IP67, IP68 |
| Explosionproof | "Ex" 포함 | Exd, Exde |

```
예: NA006 Enclosure 옵션
┌────────┬───────────┬────────────┐
│ Model  │ Enclosure │ PriceAdder │
├────────┼───────────┼────────────┤
│ NA006  │ IP67      │ 0          │  ← 기본 (Waterproof)
│ NA006  │ Exd       │ 300        │  ← 방폭 추가비
└────────┴───────────┴────────────┘

예: 방폭 전용 모델 (SA05X)
┌────────┬───────────┬────────────┐
│ Model  │ Enclosure │ PriceAdder │
├────────┼───────────┼────────────┤
│ SA05X  │ Exd       │ 0          │  ← 방폭만 지원 (기본가에 포함)
└────────┴───────────┴────────────┘
```

#### ⚠️ 누락 시 문제

**DB_EnclosureOptions에 없는 조합은 사이징에서 선택되지 않음**

---

### 3.4 DB_ElectricalData (Datasheet용)

**Datasheet 출력 시 사용되는 전기 사양 데이터**

| 컬럼 | 설명 | 단위 |
|------|------|------|
| Model | DB_Models.Model과 일치 | - |
| Voltage | 전압 | V |
| Phase | 상수 | 1, 3, DC |
| Freq | 주파수 | Hz |
| StartingCurrent_A | 기동 전류 | A |
| StartingPF | 기동 역률 | 0~1 |
| RatedCurrent_A | 정격 전류 | A |
| AvgCurrent_A | 평균 전류 | A |
| AvgPF | 평균 역률 | 0~1 |
| AvgPower_kW | 평균 전력 | kW |
| MotorPoles | 모터 극수 | 2, 4, 6 등 |

#### 입력 규칙

1. DB_PowerOptions에 등록된 모든 Model + Voltage + Phase + Freq 조합에 대해 등록
2. **사이징 로직에는 영향 없음** - Datasheet 출력용
3. 데이터 없으면 Datasheet 해당 셀 비워짐

---

### 3.5 DB_Gearboxes (기어박스)

**액추에이터와 조합 가능한 기어박스 목록**

| 컬럼 | 설명 | 단위 | 예시 |
|------|------|------|------|
| Model | 기어박스 모델명 | - | `SB-VS10` |
| Ratio | 기어비 | - | `2.5` |
| InputTorqueMax | 최대 입력 토크 | Nm | `92.6` |
| OutputTorqueMax | 최대 출력 토크 | Nm | `220` |
| Efficiency | 효율 | 0~1 | `0.96` |
| InputFlange | 입력 플랜지 | - | `F10` |
| OutputFlange | 출력 플랜지 | - | `F10` |
| MaxStemDim_mm | 최대 스템 직경 | mm | `30` |
| Weight_kg | 무게 | kg | `5` |
| Price | 가격 | - | `200` |

#### 플랜지 호환성 규칙

```
Actuator.OutputFlange = Gearbox.InputFlange 일 때만 조합 가능

예:
- MA01 (OutputFlange = F10) + SB-VS10 (InputFlange = F10) → 호환 ✓
- MA01 (OutputFlange = F10) + SB-V3 (InputFlange = F25) → 호환 ✗
```

#### 토크 제한 규칙

```
1. Actuator.Torque ≤ Gearbox.InputTorqueMax (입력 제한)
2. 출력토크 = Actuator.Torque × Ratio × Efficiency
3. 출력토크 ≤ Gearbox.OutputTorqueMax (출력 제한)
```

---

### 3.6 DB_Couplings (커플링)

**ValveList의 Coupling Type 드롭다운 및 치수 검증용**

| 컬럼 | 설명 |
|------|------|
| CouplingType | 커플링 타입명 (드롭다운에 표시) |
| MinDimension_mm | 최소 스템 직경 |
| MaxDimension_mm | 최대 스템 직경 |

```
예:
┌──────────────────────────┬─────────────────┬─────────────────┐
│ CouplingType             │ MinDimension_mm │ MaxDimension_mm │
├──────────────────────────┼─────────────────┼─────────────────┤
│ Thrust Base - Threaded   │ 20              │ 120             │
│ Standard (Part-turn)     │ 0               │ 0               │  ← 0,0 = 검증 안 함
└──────────────────────────┴─────────────────┴─────────────────┘
```

---

### 3.7 DB_Options (추가 옵션)

**Configuration 시트에서 사용하는 옵션 가격표**

| 컬럼 | 설명 | 예시 |
|------|------|------|
| Code | 옵션 코드 | `OPT-HTR` |
| Description | 옵션 설명 | `Space Heater` |
| Price | 옵션 가격 | `50` |

```
예:
┌───────────┬─────────────────────┬───────┐
│ Code      │ Description         │ Price │
├───────────┼─────────────────────┼───────┤
│ OPT-HTR   │ Space Heater        │ 50    │
│ OPT-MOD   │ Modulating Control  │ 200   │
│ OPT-POS   │ Position Transmitter│ 150   │
│ PAINT-EP  │ Epoxy Painting      │ 100   │
└───────────┴─────────────────────┴───────┘
```

---

## 4. 데이터 입력 체크리스트

### 새 모델 추가 시

```
□ Step 1: DB_Models에 기본 사양 입력
    □ 18개 컬럼 모두 확인 (해당 없으면 0 또는 빈값)
    □ Series, ActType 정확한 값 사용
    □ 50Hz/60Hz 별도 행으로 등록 (둘 다 지원하면)
    □ kW/RPM 조합별로 별도 행 (MA 시리즈)
    □ Phase 값 설정 (MS만 1/3, 나머지 0)

□ Step 2: DB_PowerOptions에 전원 옵션 등록
    □ 지원하는 모든 Voltage/Phase/Freq 조합 등록
    □ Model 값이 DB_Models와 정확히 일치하는지 확인
    □ 기본 옵션 PriceAdder = 0

□ Step 3: DB_EnclosureOptions에 보호등급 등록
    □ 지원하는 모든 Enclosure 등록 (IP67, Exd 등)
    □ Model 값 일치 확인
    □ 기본 옵션 PriceAdder = 0

□ Step 4: DB_ElectricalData에 전기 사양 등록 (Datasheet용)
    □ DB_PowerOptions의 각 조합에 대해 등록
    □ (선택사항 - 없어도 사이징 동작)
```

### 검증 방법

1. **테스트 사이징 실행**
   - 새 모델이 조건에 맞을 때 선택되는지 확인
   - Alternative 버튼으로 후보 목록에 나타나는지 확인

2. **가격 확인**
   - BasePrice + PowerAdder + EnclosureAdder 계산 확인

3. **누락 확인**
   - 모델이 선택 안 되면 → PowerOptions 또는 EnclosureOptions 누락 의심
   - "No models match..." 메시지 확인

---

## 5. 흔한 실수 및 해결

| 증상 | 원인 | 해결 |
|------|------|------|
| 모델이 사이징에서 안 나옴 | DB_PowerOptions에 해당 전원 조합 없음 | 전원 옵션 추가 |
| 모델이 사이징에서 안 나옴 | DB_EnclosureOptions에 해당 Enclosure 없음 | Enclosure 옵션 추가 |
| 모델이 사이징에서 안 나옴 | Freq 불일치 (50Hz 모델만 있는데 60Hz 선택) | 60Hz 행 추가 |
| 가격이 이상함 | PriceAdder 값 오류 | PowerOptions/EnclosureOptions 확인 |
| 기어박스 조합 안 됨 | OutputFlange ↔ InputFlange 불일치 | 플랜지 코드 확인 |
| Datasheet 전기 사양 비어있음 | DB_ElectricalData에 해당 조합 없음 | ElectricalData 추가 |

---

## 6. 데이터 형식 주의사항

1. **숫자는 숫자로**: 텍스트 "100"이 아닌 숫자 100
2. **문자열 정확히**: "Multi-turn" (O) vs "multi-turn" (X) vs "Multiturn" (X)
3. **공백 주의**: 앞뒤 공백 없이 입력
4. **소수점**: Efficiency는 0.96 (96% 아님)
5. **빈값 vs 0**: 해당 없는 숫자 컬럼은 0, 해당 없는 문자 컬럼은 빈값

---

**버전**: 2025-12-27
**작성**: Claude Code
