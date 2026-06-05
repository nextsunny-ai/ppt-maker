# PPT 메이커 — PPT MAKER

**HTML 슬라이드 덱을 PowerPoint에서 직접 편집 가능한 `.pptx`로 변환합니다.**
화면 그대로, 모든 글자가 진짜 텍스트로 들어갑니다.

> AI 사용 안 함 · API 키/로그인/과금 없음 · 로컬에서만 동작

---

## 무엇을 하나

- HTML 덱(1920×1080, 16:9)을 읽어 **네이티브 PowerPoint 요소**로 재구성
  - 글자 → 편집 가능한 텍스트 박스 (폰트·크기·색·자간 유지)
  - 박스·라인 → 네이티브 도형 · 사진 → 네이티브 이미지
  - 아웃라인(속 빈) 헤드라인도 그대로 재현
- 덱이 쓰는 **폰트 자동 감지** → 없는 폰트 원클릭 설치 → `.pptx`에 **임베드**
  (폰트 없는 PC에서도 동일하게 보임)
- 다크 / 라이트 테마 UI

## 화면

좌측 메뉴 — `변환 · 폰트 · 설정 · 정보`. 파일 선택 → 폰트 확인 → 변환 → 폴더에서 보기.

---

## 설치

```bash
git clone https://github.com/nextsunny-ai/ppt-maker.git
cd ppt-maker
npm install
```

필요 환경:

| 항목 | 용도 | 필수 |
|---|---|---|
| **Node.js** | 변환 엔진 실행 | 필수 |
| **Google Chrome** (또는 Edge) | 슬라이드 렌더링 | 필수 |
| **Microsoft PowerPoint** | 폰트 임베드 | 선택 (없어도 .pptx 생성) |

## 실행

```bash
npm start        # 로컬 서버 시작 (http://127.0.0.1:39217)
```
브라우저로 위 주소를 열거나, Windows에서는 `start.ps1` / `PPT메이커.bat` 으로 앱 창을 띄웁니다.

## 쓰는 법

1. **HTML 파일 선택** (여러 개 가능)
2. 자동 **폰트 확인** → 없는 폰트 [설치]
3. **PPT로 변환** → 원본 HTML과 같은 폴더에 같은 이름 `.pptx` 생성
4. **폴더에서 보기**

---

## 구조

```
server.js       로컬 서버 (비동기, 무거운 작업은 자식 프로세스)
html2pptx.js    변환 엔진 (puppeteer 추출 → pptxgenjs)
fontkit.js      폰트 감지 / 설치
fontcli.js      폰트 작업 CLI 래퍼
ui/             화면 (index.html · style.css · app.js)
start.ps1       서버 기동 + 앱 창 (Windows)
```

## 동작 원리

설치된 Chrome로 HTML을 1:1(noscale)로 렌더 → 각 요소의 **실제 좌표·폰트·색을 계산**해서 추출 → `pptxgenjs`로 동일 위치에 네이티브 요소를 다시 그립니다. 그래서 보이는 그대로, 그러나 편집 가능한 PPT가 나옵니다.

---

## 라이선스

© 2026 Sunny Ryu · All rights reserved.

이 소프트웨어의 모든 권리는 Sunny Ryu에게 있습니다. 무단 복제 · 재배포 · 상업적 이용을 금합니다.
