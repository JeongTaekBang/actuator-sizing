# Actuator Sizing Tool - Product Requirements Document (PRD)

## 1. 개요

### 1.1 제품명
Noah Actuator Sizing Tool

### 1.2 버전
v2.0 (2025.12)

### 1.3 목적
밸브용 전동 액추에이터 선정 업무를 자동화하여 엔지니어의 시간을 절약하고 선정 오류를 방지합니다.

### 1.4 대상 사용자
- 액추에이터 영업 엔지니어
- 견적 담당자
- 기술 지원팀

---

## 2. 문제 정의

### 2.1 현재 업무 프로세스의 문제점

| 문제 | 영향 |
|------|------|
| **수동 선정** | 밸브 요구사항에 맞는 액추에이터 + 기어박스 조합을 수작업으로 검토 |
| **시간 소요** | 여러 조합 중 최적(최저가) 모델을 찾는 시간 소요 |
| **인적 오류** | 토크/추력/작동시간 계산 오류 가능성 |
| **반복 작업** | 견적서/Datasheet 작성 시 동일 정보 반복 입력 |
| **일관성 부족** | 담당자마다 다른 선정 기준 적용 가능 |

### 2.2 해결 방안
Excel VBA 기반 자동화 도구를 제공하여:
1. 밸브 요구사항 입력 → 최적 모델 자동 선정
2. 대체 모델 목록 제공 → 사용자 선택 가능
3. 옵션 선택 및 가격 자동 계산
4. Datasheet 자동 생성

---

## 3. 기능 요구사항 (Functional Requirements)

### FR-1: 설정 관리

| ID | 요구사항 | 우선순위 | 상태 |
|:--:|----------|:--------:|:----:|
| FR-1.1 | 토크 단위 선택 (Nm, lbf.ft, kgf.m) | P0 | 완료 |
| FR-1.2 | 추력 단위 선택 (kN, lbf, kgf) | P0 | 완료 |
| FR-1.3 | 안전율 설정 (기본 1.25) | P0 | 완료 |
| FR-1.4 | 전원 사양 설정 (Voltage, Phase, Frequency) | P0 | 완료 |
| FR-1.5 | Enclosure 선택 (Waterproof, Explosionproof) | P0 | 완료 |
| FR-1.6 | Actuator Type 선택 (Multi-turn, Part-turn) | P0 | 완료 |
| FR-1.7 | Operating Time 허용 범위 설정 (%) | P0 | 완료 |
| FR-1.8 | Model Range 필터 (All, NA, SA) | P1 | 완료 |
| FR-1.9 | 커플링 타입 기본값 설정 | P1 | 완료 |
| FR-1.10 | 라인 추가 개수 설정 | P2 | 완료 |

### FR-2: 밸브 정보 입력

| ID | 요구사항 | 우선순위 | 상태 |
|:--:|----------|:--------:|:----:|
| FR-2.1 | 밸브 태그/사이즈/등급 입력 | P0 | 완료 |
| FR-2.2 | 요구 토크 입력 | P0 | 완료 |
| FR-2.3 | 요구 추력 입력 (Multi-turn) | P0 | 완료 |
| FR-2.4 | 회전수 입력 (Multi-turn) | P0 | 완료 |
| FR-2.5 | 요구 작동시간 입력 | P0 | 완료 |
| FR-2.6 | 커플링 타입/치수 입력 | P1 | 완료 |
| FR-2.7 | 밸브 타입 드롭다운 | P2 | 완료 |
| FR-2.8 | 일괄 라인 추가 기능 | P2 | 완료 |

### FR-3: 자동 사이징

| ID | 요구사항 | 우선순위 | 상태 |
|:--:|----------|:--------:|:----:|
| FR-3.1 | 토크/추력 조건 필터링 | P0 | 완료 |
| FR-3.2 | 전원 사양 필터링 | P0 | 완료 |
| FR-3.3 | Enclosure 필터링 | P0 | 완료 |
| FR-3.4 | Operating Time 범위 검증 | P0 | 완료 |
| FR-3.5 | 직접 구동 vs 기어박스 조합 비교 | P0 | 완료 |
| FR-3.6 | 최저가 모델 자동 선택 | P0 | 완료 |
| FR-3.7 | 기어박스 플랜지 호환성 검증 | P0 | 완료 |
| FR-3.8 | MaxStemDim 검증 | P1 | 완료 |
| FR-3.9 | 커플링 치수 범위 검증 | P1 | 완료 |
| FR-3.10 | 실패 사유 상세 메시지 | P1 | 완료 |

### FR-4: Alternative 선택

| ID | 요구사항 | 우선순위 | 상태 |
|:--:|----------|:--------:|:----:|
| FR-4.1 | 조건 충족 모든 모델 조회 | P0 | 완료 |
| FR-4.2 | 모델 목록 UserForm 표시 | P1 | 완료 |
| FR-4.3 | 모델 선택 시 결과 반영 | P1 | 완료 |
| FR-4.4 | 더블클릭 선택 지원 | P2 | 완료 |

### FR-5: Configuration (옵션 선택)

| ID | 요구사항 | 우선순위 | 상태 |
|:--:|----------|:--------:|:----:|
| FR-5.1 | 선정 결과 → Configuration 복사 | P0 | 완료 |
| FR-5.2 | Yes/No 옵션 선택 (HTR, MOD, POS, LMT, EXD) | P0 | 완료 |
| FR-5.3 | Painting 옵션 드롭다운 | P1 | 완료 |
| FR-5.4 | 수량 입력 | P0 | 완료 |
| FR-5.5 | Unit Price 자동 계산 | P0 | 완료 |
| FR-5.6 | Total Price 자동 계산 | P0 | 완료 |
| FR-5.7 | Grand Total 합산 | P1 | 완료 |

### FR-6: Datasheet 출력

| ID | 요구사항 | 우선순위 | 상태 |
|:--:|----------|:--------:|:----:|
| FR-6.1 | Excel 파일로 Datasheet 생성 | P0 | 완료 |
| FR-6.2 | 템플릿 기반 출력 | P1 | 완료 |
| FR-6.3 | 전기 데이터 포함 | P1 | 완료 |
| FR-6.4 | 무게 정보 출력 | P2 | 미완료 |

---

## 4. 비기능 요구사항 (Non-Functional Requirements)

### NFR-1: 사용성

| ID | 요구사항 | 상태 |
|:--:|----------|:----:|
| NFR-1.1 | 별도 설치 없이 Excel에서 실행 | 완료 |
| NFR-1.2 | 드롭다운으로 입력 오류 방지 | 완료 |
| NFR-1.3 | 입력/결과 컬럼 색상 구분 | 완료 |
| NFR-1.4 | 진행률 상태바 표시 | 완료 |
| NFR-1.5 | 오류 시 명확한 메시지 | 완료 |

### NFR-2: 성능

| ID | 요구사항 | 상태 |
|:--:|----------|:----:|
| NFR-2.1 | 100개 라인 사이징 < 10초 | 완료 |
| NFR-2.2 | 화면 업데이트 최소화 | 완료 |

### NFR-3: 유지보수성

| ID | 요구사항 | 상태 |
|:--:|----------|:----:|
| NFR-3.1 | 플랫 DB 구조 + 옵션 테이블 분리 | 완료 |
| NFR-3.2 | 모듈화된 VBA 코드 | 완료 |
| NFR-3.3 | 문서화 (README, TECHNICAL_GUIDE) | 완료 |

### NFR-4: 확장성

| ID | 요구사항 | 상태 |
|:--:|----------|:----:|
| NFR-4.1 | 새 모델 추가 용이 | 완료 |
| NFR-4.2 | 새 전원 옵션 추가 용이 | 완료 |
| NFR-4.3 | 새 Enclosure 옵션 추가 용이 | 완료 |
| NFR-4.4 | 새 옵션 추가 가이드 제공 | 완료 |

---

## 5. 데이터 요구사항

### 5.1 플랫 DB 구조 + 옵션 테이블

DB_Models는 **플랫 구조**(같은 Model이 Freq/kW/RPM 조합별로 여러 행)이며, 전원/Enclosure 옵션은 별도 테이블로 분리됩니다.

```
DB_Models (기본 사양 - 플랫 구조)
    ├── DB_PowerOptions (전원 옵션)
    ├── DB_EnclosureOptions (보호등급 옵션)
    └── DB_ElectricalData (전기 데이터)

DB_Gearboxes (기어박스)

DB_Couplings (커플링)

DB_Options (가격 옵션)
```

### 5.2 필수 데이터

| 테이블 | 필수 컬럼 |
|--------|----------|
| DB_Models | Model, Series, Type, Torque_Nm, RPM, OutputFlange, BasePrice |
| DB_PowerOptions | Model, Voltage, Phase, Freq, PriceAdder |
| DB_EnclosureOptions | Model, Enclosure, PriceAdder |
| DB_Gearboxes | Model, Ratio, InputTorqueMax, Efficiency, InputFlange, Price |

### 5.3 가격 계산

```
액추에이터 가격 = BasePrice + PowerAdder + EnclosureAdder
총 가격 = 액추에이터 가격 + 기어박스 가격 + 선택 옵션 합계
```

---

## 6. 제약 조건 및 가정

### 6.1 제약 조건

| 항목 | 설명 |
|------|------|
| 플랫폼 | Microsoft Excel (Windows) |
| 매크로 | VBA 매크로 활성화 필요 |
| 데이터 | 샘플 데이터 포함 (실제 데이터로 교체 필요) |

### 6.2 가정

| 항목 | 설명 |
|------|------|
| ISO 5211 | 토크 만족 시 플랜지 호환 (별도 검증 불필요) |
| 밸브 호환 | 밸브사 책임 (액추에이터 출력 토크 정보 제공) |
| 단위 | DB는 SI 단위 (Nm, kN, mm) |

---

## 7. 로드맵

### 7.1 현재 버전 (v2.0)

- [x] 자동 사이징 엔진
- [x] Alternative 선택 UI
- [x] Configuration 옵션/가격 관리
- [x] Datasheet 출력
- [x] 플랫 DB 구조 + 옵션 테이블

### 7.2 계획된 기능 (Future)

| 기능 | 우선순위 | 설명 |
|------|:--------:|------|
| 무게 출력 | P1 | Datasheet에 액추에이터/기어박스 무게 표시 |
| 60Hz 지원 | P1 | 60Hz 모델 데이터 완비 및 검증 |
| Operation Mode | P2 | On-Off / Modulating 별 필터링 |
| 온도 범위 | P2 | Ambient Temperature 필터링 |
| 인증 정보 | P3 | CE, ATEX, UL 등 인증 필터 |
| Duty Cycle | P3 | S2, S4 등 운전 주기 필터 |

---

## 8. 용어 정의

| 용어 | 정의 |
|------|------|
| Multi-turn | 다회전 액추에이터 (Gate, Globe 밸브용) |
| Part-turn | 1/4회전 액추에이터 (Ball, Butterfly 밸브용) |
| Torque | 회전력 (단위: Nm) |
| Thrust | 추력 (단위: kN) - Multi-turn만 해당 |
| Operating Time | 밸브 전개/폐쇄 시간 (단위: 초) |
| Gearbox Ratio | 기어비 - 입력 회전수 대비 출력 회전수 비율 |
| Efficiency | 기어박스 효율 (0.85~0.95) |
| Safety Factor | 안전율 - 요구 토크에 곱하는 여유 계수 |
| MaxStemDim | 액추에이터/기어박스가 수용 가능한 최대 밸브 스템 직경 |
| Enclosure | 보호등급 (IP67: Waterproof, Exd: Explosionproof) |

---

## 9. 관련 문서

| 문서 | 설명 |
|------|------|
| [README.md](README.md) | 사용자 가이드 (설치, 사용법) |
| [TECHNICAL_GUIDE.md](TECHNICAL_GUIDE.md) | 기술 문서 (알고리즘, 데이터 흐름) |
| [QA_GUIDE.md](QA_GUIDE.md) | Q&A 가이드 (자주 묻는 질문) |
| [PILOT_GUIDE.md](PILOT_GUIDE.md) | Pilot 프로그램 안내 |
| [CLAUDE.md](CLAUDE.md) | 개발자 가이드 (Claude Code용) |

---

## 문서 버전

- **v2.0** (2025.12) - 초기 작성 (플랫 DB 구조 + 옵션 테이블 반영)
