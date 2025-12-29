# Actuator Sizing Tool - Pilot 프로그램 안내

## 개요

안녕하세요,

밸브용 전동 액추에이터 선정 업무를 지원하기 위한 **Sizing Tool Pilot 버전**을 개발했습니다.
실무에 적용하기 전에 여러분의 피드백을 받고자 합니다.

---

## 1. 프로그램 목적

### 해결하고자 하는 문제
- 밸브 요구사항(토크, 추력, 작동시간)에 맞는 액추에이터 + 기어박스 조합을 **수작업으로 선정**하는 번거로움
- 여러 조합 중 **최적(최저가) 모델**을 찾는 시간 소요
- 견적서/Datasheet 작성 시 **반복 작업**

### 이 프로그램이 하는 일
1. **자동 사이징**: 밸브 요구사항 입력 → 최적 액추에이터(+기어박스) 자동 선정
2. **대체 모델 조회**: 조건을 만족하는 모든 모델 목록 표시, 수동 선택 가능
3. **Configuration**: 옵션(Heater, Modulating 등) 선택 및 가격 자동 계산
4. **Datasheet 출력**: 선정 결과를 Excel Datasheet로 자동 생성

---

## 2. 사용 방법 요약

### Step 1: Settings 설정
| 항목 | 설명 |
|------|------|
| Actuator Type | Multi-turn / Part-turn / Linear |
| Voltage / Phase / Freq | 전원 사양 (예: 380V 3ph 50Hz) |
| Enclosure | Waterproof / Explosionproof |
| Safety Factor | 안전율 (기본 1.25) |
| Operation Mode | On-Off / Modulating / High-Speed |
| Fail-safe | None / Close-on-Fail (SR) |
| Duty Cycle | Any / Intermittent (S2) / Continuous (S4) |
| Torque/Thrust Unit | 단위 선택 |

### Step 2: ValveList에 밸브 정보 입력
| 입력 항목 | 설명 |
|-----------|------|
| Tag | 밸브 태그 번호 |
| ValveType | Gate, Globe, Ball, Butterfly, Plug, Linear |
| Torque | 필요 토크 (Multi-turn/Part-turn) |
| Thrust | 필요 추력 (Multi-turn/Linear) |
| Lift (mm) | 밸브 스템 이동거리 (Multi-turn) |
| Pitch (mm) | 스템 나사 피치 (Multi-turn), Turns = Lift / Pitch |
| Op.Time | 요구 작동시간 (초) |
| CouplingDim | 밸브 스템 직경 (mm) |

> **ValveType → ActuatorType 자동 결정**:
> - Ball, Butterfly, Plug → Part-turn
> - Gate, Globe → Multi-turn
> - Linear → Linear

### Step 3: Sizing 실행
- **Sizing All**: 전체 라인 사이징
- **Sizing Selected**: 선택한 라인만 사이징
- 결과가 자동으로 채워짐 (Model, Gearbox, Price 등)

### Step 4: 필요 시 대체 모델 선택
- **Alternative** 버튼 → 조건 만족하는 모든 모델 목록 표시
- 원하는 모델 선택 가능

### Step 5: Configuration에서 옵션 선택
- **To Configuration** 버튼 → 선정 결과 복사
- HTR, MOD, POS 등 옵션 Yes/No 선택
- Unit Price, Total 자동 계산

### Step 6: Datasheet 출력
- **Export Datasheet** 버튼 → Excel 파일로 저장

---

## 3. 사이징 로직 요약

### 선정 기준
```
1. Actuator Type 일치 (Multi-turn / Part-turn / Linear)
2. 전원 사양 일치 (Voltage, Phase, Frequency)
3. Enclosure 일치 (Waterproof → IP67 / Explosionproof → Exd)
4. Fail-safe 일치 (SR 선택 시 Spring Return 시리즈만)
5. Duty Cycle 일치 (S2/S4 필터)
6. 토크 ≥ 요구토크 × 안전율 (Multi-turn/Part-turn)
7. 추력 ≥ 요구추력 × 안전율 (Multi-turn/Linear)
8. 작동시간이 허용 범위 내 (기본 ±50%)
9. 스템 직경 ≤ MaxStemDim
```

### 기어박스 조합 시 (Multi-turn/Part-turn)
```
출력 토크 = 액추에이터 토크 × 기어비 × 효율

Multi-turn: 작동시간 = (회전수 × 기어비 × 60) / 액추에이터 RPM
Part-turn:  작동시간 = DB OpTime × 기어비
Linear:     기어박스 미사용 (직접 연결만)
```

### 최적 모델 선정
- 조건 만족하는 모델 중 **최저 가격** 선택
- 직접 구동 vs 기어박스 조합 중 저렴한 쪽 선택

---

## 4. 현재 상태 (Pilot)

### 완성된 기능
- [x] 자동 사이징 (직접 구동 + 기어박스 조합)
- [x] 대체 모델 조회 및 선택
- [x] Configuration 옵션 선택 및 가격 계산
- [x] Datasheet Excel 출력
- [x] 단위 변환 (Nm/lbf.ft/kgf.m, kN/lbf/kgf)
- [x] 작동시간 범위 검증
- [x] 스템 직경 검증 (MaxStemDim)

### 미완성 / 추후 개선 필요
| 항목 | 상태 | 비고 |
|------|------|------|
| DB 데이터 | **구조 완성** | NA, SA, MA, MS, SR, NL 전 시리즈 지원 |
| Weight 출력 | ✅ 구현됨 | Datasheet에 Actuator/Gearbox/조합 무게 출력 |
| 60Hz 모델 | ✅ 추가됨 | 각 시리즈별 50Hz/60Hz 데이터 포함 |
| 기어박스 | ✅ 삼보 통합 | Bevel, Spur, Worm 31개 대표 모델 |
| 가격 데이터 | **placeholder** | 실제 단가로 업데이트 필요 |
| Weight 수치 | **0** | 실제 무게 데이터 수동 입력 필요 |
| 전기 데이터 | 샘플 값 | Datasheet 전기 사양란 실제 값 필요 |

---

## 5. DB 시트 구조 (플랫 + 옵션 테이블, 실제 데이터 입력 필요)

### DB_Models (숨김 시트) - 모델 기본 사양 (플랫 구조, 18컬럼)
```
Model | Series | ActType | MotorPower_kW | ControlType | Phase | Freq | RPM | Torque_Nm | Thrust_kN | OpTime_sec | DutyCycle | OutputFlange | MaxStemDim_mm | Weight_kg | BasePrice | Speed_mm_sec | Stroke_mm
```
> **플랫 구조**: 같은 Model이 Freq/kW/RPM 조합별로 여러 행 등록

### DB_PowerOptions (숨김 시트) - 전원 옵션
```
Model | Voltage | Phase | Freq | PriceAdder
```

### DB_EnclosureOptions (숨김 시트) - 보호등급 옵션
```
Model | Enclosure | PriceAdder
```

### DB_ElectricalData (숨김 시트) - Datasheet용 전기 데이터
```
Model | Voltage | Phase | Freq | StartingCurrent_A | StartingPF | RatedCurrent_A | AvgCurrent_A | AvgPF | AvgPower_kW | MotorPoles
```

### DB_Gearboxes (숨김 시트)
```
Model | Ratio | InputTorqueMax | OutputTorqueMax | Efficiency | InputFlange | OutputFlange | MaxStemDim_mm | Weight_kg | Price
```

### DB_Options (숨김 시트)
```
Code | Description | Price
```
예: OPT-HTR, OPT-MOD, PAINT-EP 등

### 가격 계산
```
최종 가격 = BasePrice + PowerAdder + EnclosureAdder
```

---

## 6. 피드백 요청 사항

이 Pilot 프로그램을 검토하시고 아래 사항에 대해 의견 부탁드립니다:

### 기능 관련
- [ ] 사이징 로직이 실무와 맞는지?
- [ ] 누락된 검증 조건이 있는지?
- [ ] 추가로 필요한 기능이 있는지?

### 데이터 관련
- [ ] DB 컬럼 구조가 적절한지?
- [ ] 추가로 필요한 데이터 항목이 있는지?

### UI/UX 관련
- [ ] 입력 항목이 직관적인지?
- [ ] 결과 표시가 명확한지?
- [ ] Datasheet 양식이 적절한지?

### 기타
- [ ] 개선이 필요한 부분
- [ ] 우선순위가 높은 기능

---

## 7. 실행 방법

### Step 1: 파일 열기
`NoahSizing.xlsm` 파일을 Excel에서 엽니다.

### Step 2: 매크로 활성화
파일을 열면 상단에 **보안 경고** 메시지가 표시됩니다:

```
보안 경고: 매크로가 사용하지 않도록 설정되었습니다.  [콘텐츠 사용]
```

**[콘텐츠 사용]** 버튼을 클릭하여 매크로를 활성화합니다.

> **참고**: 보안 경고가 표시되지 않으면 아래 방법으로 매크로를 활성화하세요.

### 매크로가 차단된 경우
1. **파일** → **옵션** → **보안 센터** → **보안 센터 설정**
2. **매크로 설정** → **모든 매크로 포함** 또는 **알림을 표시하고 사용자가 선택** 선택
3. Excel 재시작 후 파일 다시 열기

### 인터넷에서 다운로드한 경우 (이메일 첨부 등)
1. 파일 우클릭 → **속성**
2. 하단의 **차단 해제** 체크 → **확인**
3. 파일 다시 열기

---

## 8. 문의

질문이나 피드백이 있으시면 언제든 연락 주세요.

감사합니다.

---

**버전**: Pilot v3.0 (Linear 지원, SR 시리즈, 삼보 기어박스 통합)
**작성일**: 2025년 12월
