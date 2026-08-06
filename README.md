# Claude Code 팀 표준 — `/new-project` 스킬

새 프로젝트를 시작할 때 팀 표준 워크플로우 파일과 표준 문서를 자동으로 세팅해 주는 Claude Code 스킬입니다.

## 사용 방법

Claude Code에서 아래처럼 입력하면 됩니다.

```
/new-project 프로젝트명
```

그러면 자동으로:

1. `프로젝트명/` 폴더 생성
2. **GOAL.md** — 목적·목표·완료 기준을 간단히 인터뷰해서 채워줌 (프로젝트의 "무엇을, 왜")
3. **PROGRESS.md** — Phase별 진행 상황 + 검증 로그 템플릿 (검증 없는 진행은 진행이 아니다)
4. **CLAUDE.md** — 프로젝트 규칙 등록: 작업 전 GOAL/PROGRESS 읽기, Phase 단위 진행 + 승인, 슬라이드 작업 시 디자인 가이드 자동 적용
5. `docs/DESIGN_GUIDE.md` — PPTX 슬라이드 디자인 가이드 복사
6. `docs/MAC_SETUP_GUIDE.md` — Mac 셋업 가이드 복사
7. README 생성 + git 초기화

프로젝트명 없이 `/new-project`만 입력하면 이름을 물어봅니다. "현재 폴더에 세팅해줘"라고 하면 새 폴더 없이 현재 위치에 세팅합니다.

## 어디서나 쓰려면 (전역 설치)

이 저장소 안에서는 `/new-project`가 바로 동작합니다. **다른 폴더에서도** 쓰고 싶으면, 본인 Mac에서 스킬을 홈 디렉토리로 복사하세요.

```bash
mkdir -p ~/.claude/skills
cp -R .claude/skills/new-project ~/.claude/skills/
```

이후 어떤 폴더에서 Claude Code를 열어도 `/new-project`를 쓸 수 있습니다.

## 구성

```
.claude/skills/new-project/
├── SKILL.md                        # 스킬 정의 (수행 절차)
├── templates/                      # 프로젝트 루트에 복사되는 워크플로우 템플릿
│   ├── CLAUDE.md                   # 프로젝트 규칙 (디자인 가이드 규칙 포함)
│   ├── GOAL.md                     # 목표·완료 기준 정의
│   └── PROGRESS.md                 # Phase별 진행·검증 로그
└── references/                     # docs/로 복사되는 표준 문서
    ├── design-guide.md             # PPTX 슬라이드 디자인 가이드
    └── mac-setup-guide.md          # Mac 셋업 가이드
```

표준 문서나 템플릿을 업데이트하려면 해당 파일을 수정하고 커밋하면 됩니다. 전역 설치한 경우 `~/.claude/skills/new-project`에도 다시 복사해야 반영됩니다.
