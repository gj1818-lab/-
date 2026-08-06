# Claude Code 팀 표준 — `/new-project` 스킬

새 프로젝트를 시작할 때 팀 표준 문서(PPTX 디자인 가이드, Mac 셋업 가이드)를 자동으로 심어주는 Claude Code 스킬입니다.

## 사용 방법

Claude Code에서 아래처럼 입력하면 됩니다.

```
/new-project 프로젝트명
```

그러면 자동으로:

1. `프로젝트명/` 폴더 생성
2. `docs/DESIGN_GUIDE.md` — PPTX 슬라이드 디자인 가이드 복사
3. `docs/MAC_SETUP_GUIDE.md` — Mac 셋업 가이드 복사
4. `CLAUDE.md` 생성 — 이후 이 프로젝트에서 슬라이드/PPTX 작업 시 디자인 가이드가 **자동 적용**되도록 규칙 등록
5. `README.md` 생성 + git 초기화

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
└── references/
    ├── design-guide.md             # PPTX 슬라이드 디자인 가이드 (원본)
    └── mac-setup-guide.md          # Mac 셋업 가이드 (원본)
```

표준 문서를 업데이트하려면 `references/` 안의 파일을 수정하고 커밋하면 됩니다. 전역 설치한 경우 `~/.claude/skills/new-project`에도 다시 복사해야 반영됩니다.
