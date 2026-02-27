# k-grant-apply

**Claude Code Skill** — 한국연구재단(NRF) 연구개발과제 연구계획서 자동 생성 스킬

## 개요

사용자의 연구 내용을 입력하면 NRF 표준 양식(HWPX)에 맞는 연구계획서를 자동으로 생성합니다.
RFP(사업공고문) 기반 맞춤 작성, 연차·단계 자동 확장, 수행일정 표 자동 채움을 지원합니다.

## 주요 기능

- **RFP 자동 분석**: 사업공고문에서 연구기간·연차·평가기준·분량 기준 자동 추출
- **연차 블록 자동 확장**: `N차년도` 플레이스홀더를 실제 연차(1차년도·2차년도·3차년도)로 자동 교체
- **단일세포 오믹스 파이프라인 특화**: scRNA-seq, snATAC-seq, WGBS, Hi-C 분석 파이프라인 상세 기술 지원
- **수행일정 표 자동 채움**: 연차별 추진내용·결과물 Gantt 테이블 자동 입력
- **2단계 레벨 구조**: `○` 명사형 소제목 + `   -` 세부 항목(최소 3~4개)

## 파일 구조

```
k-grant-apply/
├── SKILL.md                   ← Claude Code 스킬 정의 (트리거·워크플로우·작성 기준)
├── README.md
├── assets/
│   ├── application.hwpx       ← NRF 표준 연구계획서 양식 (2·3차년도 Gantt 9행 pre-expanded)
│   └── application.hwpx.bak   ← 원본 백업
└── scripts/
    ├── write_hwpx.py          ← HWPX 생성 Python 스크립트 (텍스트 채움 전용)
    └── expand_template.py     ← 템플릿 Gantt 행 사전 확장 스크립트
```

## 설치 방법

```bash
# Claude Code 스킬 디렉토리에 복제
git clone https://github.com/joonan-lab/k-grant-apply \
    ~/.claude/skills/k-grant-apply

# 의존성 설치
pip install lxml
```

## 사용 방법

Claude Code에서 다음과 같이 요청하면 스킬이 자동으로 활성화됩니다:

```
NRF 연구계획서 작성해줘
```

또는

```
한국연구재단 과제 신청서 작성 도와줘
```

### HWPX 직접 생성

```bash
python3 ~/.claude/skills/k-grant-apply/scripts/write_hwpx.py \
    --template ~/.claude/skills/k-grant-apply/assets/application.hwpx \
    --output ~/Desktop/NRF_연구계획서.hwpx \
    --data-json /path/to/data.json
```

### JSON 데이터 형식

```json
{
  "_meta": {
    "total_years": 3,
    "stage": 1,
    "year1_months": 6,
    "year2_months": 12,
    "year3_months": 12
  },
  "necessity": [
    "○ 명사형 소제목",
    "   - 세부 내용 1",
    "   - 세부 내용 2",
    "   - 세부 내용 3"
  ],
  "final_goal": ["○ 최종 목표"],
  "yearly_goals": {
    "year1_main": "1차년도 주관과제 목표",
    "year1_joint": "",
    "year1_contracted": "",
    "year2_main": "2차년도 주관과제 목표",
    "year3_main": "3차년도 주관과제 목표"
  },
  "yearly_contents": {
    "year1_main": ["연구내용 1 (파이프라인 포함)", "연구내용 2"],
    "year2_main": ["..."],
    "year3_main": ["..."]
  },
  "schedule": {
    "year1": [{"task": "추진내용", "result": "결과물"}],
    "year2": [],
    "year3": []
  },
  "strategy": ["○ 명사형 전략", "   - 세부 내용"],
  "system": ["○ 주관기관 역할", "   - 세부 역할"],
  "utilization": ["○ 활용방안"],
  "effects": ["○ 기대효과"],
  "commercialization": {
    "market_size": [], "demand": [], "competition": [], "ip": [],
    "standardization": [], "biz_strategy": [], "investment": [], "production": []
  }
}
```

## 글쓰기 원칙

- `○` 항목: **명사형**으로 끝남 (예: `○ ASD 유병률 현황 및 사회경제적 부담`)
- `   -` 항목: **문장형**, `임/음/함/됨`으로 끝남, 각 `○` 아래 **최소 3~4개**
- `□` 소제목 레벨 사용 안 함

## 요구사항

- Python 3.8+
- [lxml](https://lxml.de/) (`pip install lxml`)
- [Claude Code](https://claude.ai/claude-code)
- 한컴오피스 (생성된 HWPX 파일 열기)

## 라이선스

MIT License
