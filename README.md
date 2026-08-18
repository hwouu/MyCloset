# MyCloset

브라우저에서 옷장, 날짜별 아웃핏, 룩북, 위시리스트를 한곳에서 관리하는 개인용 웹 애플리케이션입니다.

## 주요 기능

- 옷장: 옷 등록·수정, 카테고리 관리, 검색·정렬, 그리드/테이블 보기
- 상품 URL 자동 입력: 무신사·유니클로 및 표준 상품 메타데이터를 제공하는 페이지에서 상품명, 브랜드, 색상, 이미지 등을 가져온 뒤 사용자가 확인·수정
- 아웃핏: 날짜별 아이템 구성, 메모, 달력 탐색, 룩북 저장
- 룩북: 자주 입는 조합 저장, 검색·정렬, 원하는 날짜에 적용
- 위시리스트: 구매 후보 상품 저장, 검색·정렬
- 데이터 이동: 옷장 Excel 가져오기/내보내기 및 전체 데이터 JSON 백업/복원
- 기기 간 동기화: Microsoft 계정의 OneDrive 앱 전용 폴더를 이용한 수동 업로드·다운로드와 충돌 감지
- PC 단축키: 검색(`/`), 새 항목(`N`), 아웃핏 날짜 이동(`←`/`→`), 오늘로 이동(`T`), 달력(`C`), 닫기(`Esc`), 도움말(`?`)
- 반응형 UI: PC, 태블릿, 모바일과 라이트/다크 모드 지원

## 기술 스택

- React 19
- Vite 6
- SheetJS (`xlsx`)
- Phosphor Icons
- Fetch API 기반 서버/Worker 상품 메타데이터 API

## 로컬 실행

Node.js 20 이상을 권장합니다.

```bash
npm install
npm run dev
```

개발 서버가 안내하는 로컬 주소로 접속하면 됩니다.

## 테스트 및 빌드

```bash
npm test
npm run build
```

빌드 결과는 `dist/client`과 `dist/server`에 생성됩니다. 상품 URL 자동 입력 기능을 배포 환경에서도 사용하려면 정적 파일과 함께 `worker/index.js`의 `/api/product-metadata` 엔드포인트가 동작해야 합니다.

## 데이터 저장과 백업

MyCloset은 별도 데이터베이스 없이 브라우저 로컬 저장소에 데이터를 보관합니다. 다른 브라우저나 컴퓨터로 옮길 때는 설정의 전체 데이터 백업/복원 또는 OneDrive 동기화를 사용하세요.

- Excel 관리는 옷장 데이터만 대상으로 합니다.
- 전체 JSON 백업에는 옷장, 업로드 이미지, 아웃핏, 메모, 룩북, 스크랩, 위시리스트와 현재 탭·정렬·보기 방식 등의 화면 설정이 포함됩니다.
- 백업 파일에는 개인 데이터가 포함될 수 있으므로 안전하게 보관하세요.

## OneDrive 동기화 설정

동기화는 Microsoft Graph의 앱 전용 폴더 권한(`Files.ReadWrite.AppFolder`)만 사용합니다. MyCloset은 Microsoft 비밀번호를 받거나 저장하지 않습니다.

1. [Microsoft Entra 관리 센터](https://entra.microsoft.com/)에서 새 앱 등록을 만듭니다.
2. `인증 > 플랫폼 추가 > 단일 페이지 애플리케이션(SPA)`에서 개발 주소 `http://localhost:5173`과 배포 주소 `https://mycloset-mocha.vercel.app`을 Redirect URI로 등록합니다.
3. 지원 계정 유형은 개인 Microsoft 계정을 포함하도록 선택합니다.
4. `.env.example`을 참고해 로컬 `.env.local`에 애플리케이션 Client ID를 넣습니다.
5. Vercel 프로젝트의 Environment Variables에도 `VITE_MS_CLIENT_ID`와 `VITE_MS_TENANT_ID=consumers`를 추가한 뒤 다시 배포합니다.

```bash
VITE_MS_CLIENT_ID=00000000-0000-0000-0000-000000000000
VITE_MS_TENANT_ID=consumers
```

연결 후 자동 동기화가 기본으로 켜지며 다음 규칙이 적용됩니다.

- OneDrive 파일이 없으면 현재 기기 데이터를 최초 업로드합니다.
- 한쪽만 변경된 경우 변경된 쪽을 기준으로 자동 업로드 또는 다운로드합니다.
- 양쪽 모두 변경된 경우 선택 모달에서 현재 기기 또는 OneDrive 중 유지할 데이터를 고릅니다.
- 데이터 변경 후 3초, 앱 실행·화면 복귀·온라인 전환 시 최신 상태를 확인합니다.
- 앱이 보이는 동안에는 60초마다 변경을 확인하며, 닫힌 브라우저에서는 실행되지 않습니다.
- 설정에서 자동 동기화를 끄거나 `지금 동기화`로 수동 확인할 수 있습니다.

## 프로젝트 구조

```text
src/       React UI와 도메인 로직
worker/    상품 메타데이터 API와 배포 엔트리
tests/     단위·통합·Worker 테스트
public/    파비콘과 Excel 등록 템플릿
scripts/   배포용 빌드 준비 스크립트
```
