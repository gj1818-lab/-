# Mac 셋업 가이드 — Claude Code 팀 환경

> **대상**: SCM 팀 신규 MacBook Pro
> **방식**: Homebrew(맥용 앱 설치 관리자)로 모든 것을 설치
> 

---

## ⚠️ 시작 전 필수 확인

사내 네트워크에서 아래 도메인이 **허용**되어야 정상 작동합니다. 막혀 있으면 IT팀에 먼저 요청하세요.

- `api.anthropic.com` — Claude Code 작동 필수
- `claude.ai` — 로그인
- `github.com`, `registry.npmjs.org` — 설치용

각 단계는 위에서 아래 순서대로 진행하세요. 명령어 박스를 그대로 복사해 터미널에 붙여넣으면 됩니다.

---

## STEP 1. Homebrew 설치

**무엇인가**: Mac용 앱 설치 관리자입니다. 앞으로 모든 프로그램을 이걸로 설치합니다.
**참고**: 중간에 Mac 로그인 비밀번호를 물어볼 수 있습니다.

```bash
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
```

설치가 끝나면 아래 명령으로 PATH를 등록합니다 (Apple Silicon 기준). 이걸 해야 `brew` 명령이 인식됩니다.

```bash
echo 'eval "$(/opt/homebrew/bin/brew shellenv)"' >> ~/.zprofile
eval "$(/opt/homebrew/bin/brew shellenv)"
```

정상 설치 확인:

```bash
brew --version
```

---

## STEP 2. Claude Code 핵심 (필수)

### 2-1. git + node

**git**: 코드·문서의 변경 이력 관리 (표준 파일 배포에도 필수)
**node**: 일부 개발 도구가 의존하는 런타임 (npm 포함)

```bash
brew install git node
```

### 2-2. Claude Code (데스크탑 앱 + CLI 통합)

**무엇인가**: Claude Code 본체입니다. 이 하나로 **데스크탑 앱과 터미널 CLI가 모두 설치**됩니다. 별도로 두 번 깔 필요 없습니다.

```bash
brew install --cask claude-code
```

> 💡 최신 버전을 항상 받고 싶다면 `claude-code` 대신 `claude-code@latest` 를 쓰세요.
> Homebrew 설치는 자동 업데이트가 안 되므로, 가끔 `brew upgrade --cask claude-code` 로 직접 업데이트합니다.

설치 확인:

```bash
claude --version
```

> ⚠️ `claude` 명령이 안 보이면 터미널을 껐다 켠 뒤 다시 확인하세요.

---

## STEP 3. 터미널 · 편집기 (권장)

### 3-1. iTerm2 + VS Code

**iTerm2**: Mac 기본 터미널보다 편한 개발자 표준 터미널
**VS Code**: 코드·문서 편집기 (Claude Code와 연동 좋음)

```bash
brew install --cask iterm2 visual-studio-code
```

### 3-2. Oh My Zsh (선택적 권장)

**무엇인가**: Mac 기본 셸(zsh)에 자동완성·테마·색상을 더해 훨씬 쓰기 편하게 만듭니다.

```bash
sh -c "$(curl -fsSL https://raw.githubusercontent.com/ohmyzsh/ohmyzsh/master/tools/install.sh)"
```

---

## STEP 4. CLI 유틸리티 (권장)

**무엇인가**: 터미널 작업을 편하게 해주는 작은 도구 묶음입니다.

- **gh**: GitHub 저장소 clone·인증을 쉽게
- **ripgrep**: 초고속 텍스트 검색 (Claude Code가 활용)
- **jq**: JSON 설정 파일 다루기
- **tree**: 폴더 구조를 트리 형태로 보기
- **wget**: 파일·데이터 다운로드

```bash
brew install gh ripgrep jq tree wget
```

---

## STEP 5. 앱 · 생산성 도구

**무엇인가**: 업무·협업·화면 관리에 쓰는 데스크탑 앱들입니다.

- **Google Drive**: 파일 동기화·공유
- **Stats**: 메뉴바에서 CPU·메모리·네트워크 실시간 모니터
- **Google Chrome**: 표준 브라우저 (Claude in Chrome 등 활용)
- **Rectangle**: 단축키로 창 정렬·분할 (생산성 필수템)
- **Raycast**: Spotlight 강화판 런처 (검색·클립보드·단축 실행)

```bash
brew install --cask google-drive stats google-chrome rectangle raycast
```

---

## STEP 6. 선택 설치 (원하는 팀원만)

### 6-1. Kaku — AI 코딩용 터미널

**무엇인가**: AI 코딩에 특화된 신생 터미널 앱입니다. Homebrew 공식이 아니라 **별도 tap**으로 설치합니다. (macOS 전용)

```bash
brew install --cask tw93/tap/kakuku
```

### 6-2. Python 개발 환경

**무엇인가**: Python 프로젝트(예: costing-system)를 다루는 팀원만 설치합니다.
**uv**: pip보다 훨씬 빠른 최신 Python 패키지 관리자

```bash
brew install python@3.12 uv
```

### 6-3. 터미널 강화 유틸

- **bat**: `cat` 개선판 (문법 하이라이트)
- **eza**: `ls` 개선판 (색상·아이콘)
- **fzf**: 터미널 파일·명령 퍼지 검색

```bash
brew install bat eza fzf
```

---

## STEP 7. Claude Code 로그인

터미널에 아래를 입력하면 브라우저가 열리며 로그인합니다.

```bash
claude
```

> ⚠️ 로그인이 안 되면 사내 네트워크에서 `api.anthropic.com` 이 차단된 것입니다. IT팀에 요청하세요.

---

## STEP 8. 최종 검증

설치가 정상인지 버전을 한 번에 확인합니다.

```bash
echo "brew:  $(brew --version | head -1)"
echo "git:   $(git --version)"
echo "node:  $(node --version)"
echo "npm:   $(npm --version)"
echo "gh:    $(gh --version | head -1)"
echo "claude: $(claude --version 2>/dev/null || echo '확인 필요 - 터미널 재시작')"
```

모든 항목에 버전 번호가 나오면 셋업 완료입니다. 🎉

---

## 다음 단계

1. 테스트용 폴더를 하나 만들고 그 안에서 `claude` 를 실행해 봅니다.
2. "이 폴더 구조 설명해줘" 같은 간단한 요청으로 응답이 오는지 확인합니다.
3. 팀 표준 설정 파일(CLAUDE.md 등)은 관리자가 별도 배포합니다.

---

## 부록: 전체 한 번에 설치 (숙련자용)

위 단계를 이해했다면, 핵심+권장 항목을 한 번에 설치할 수도 있습니다.

```bash
# CLI 도구
brew install git node gh ripgrep jq tree wget

# Claude Code + 앱
brew install --cask claude-code iterm2 visual-studio-code \
  google-drive stats google-chrome rectangle raycast
```

---

## 참고: 업데이트 & 삭제 명령어

```bash
# 전체 업데이트 확인
brew update && brew outdated

# Claude Code 업데이트
brew upgrade --cask claude-code

# 특정 앱 삭제 (예시)
brew uninstall --cask stats
```
