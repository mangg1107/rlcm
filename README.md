# rlcm
룰렛

## Vercel

- 정적 파일은 `public/`에 둡니다.
- API 함수는 `api/`에 둡니다.
- 브라우저 코드는 `/api/...` 경로로 API를 호출합니다.
- Google Sheets 인증은 Vercel 환경변수로 넣습니다.
- 로그는 `로그` 시트에 저장합니다. 없으면 API가 자동으로 만들고, 헤더는 `id, type, text, publicText, createdAt`을 사용합니다.

환경변수 방식 중 하나를 사용하세요.

1. `GOOGLE_SERVICE_ACCOUNT_JSON`: 서비스 계정 JSON 전체 문자열
2. `GOOGLE_CLIENT_EMAIL`, `GOOGLE_PRIVATE_KEY`: 서비스 계정 이메일과 private key

로컬 실행:

```bash
npm install
npm start
```
