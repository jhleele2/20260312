# 재고·발주 자동화 웹 시스템

재고를 파악하고, 부족 시 담당 기업에 **발주서 이메일**을 보내는 웹 시스템입니다. **팀 비밀번호** 입력 후 사용 가능.

- **데이터**: `domino_inventory_training.xlsx` 구조의 엑셀 파일 (시트: Suppliers, Inventory, EmailTemplate)
- **발신 이메일**: `jhleele2@gmail.com` (환경변수로 변경 가능)
- **동작**: 엑셀 데이터 분석 → 발주 필요 품목·공급업체별 요약 → 발주서 이메일 일괄 발송
- **팀 비밀번호**: 기본 `1234` (환경변수 `TEAM_PASSWORD`로 변경 가능)

## 설치 및 실행

```bash
cd c:\Users\SD2-20\Desktop\temp2
pip install -r requirements.txt
```

이메일 발송을 쓰려면 `.env` 파일을 만들고 Gmail 앱 비밀번호를 넣습니다.

```bash
copy .env.example .env
# .env 편집: SMTP_PASSWORD=실제_앱_비밀번호
```

Gmail 앱 비밀번호: [Google 앱 비밀번호](https://myaccount.google.com/apppasswords)에서 생성.

```bash
python app.py
```

브라우저에서 **http://localhost:5000** 접속.

## 기능

1. **엑셀 업로드**  
   같은 구조의 엑셀을 업로드하면 해당 파일 기준으로 재고·발주 요약을 봅니다.

2. **재고 현황**  
   품목별 현재고, 안전재고, MOQ, **발주수량**(= MAX(MOQ, 안전재고−현재고)), 상태를 표로 표시.

3. **공급업체별 발주 요약**  
   발주가 필요한 품목을 공급업체별로 묶어서 표시.

4. **발주서 이메일 일괄 발송**  
   각 공급업체 이메일로 엑셀 `EmailTemplate` 시트의 제목/본문 형식({{STORE_NAME}}, {{SUPPLIER_NAME}}, {{ORDER_DATE}}, {{ITEM_LIST}}, {{INTERNAL_OWNER}})을 채워 발송합니다. 발신은 `jhleele2@gmail.com`입니다.

## Vercel 배포 (메일 발송하려면 필수)

1. [Vercel](https://vercel.com)에 로그인 후 **Import Git Repository**에서 이 저장소 연결.
2. **Settings → Environment Variables**에서 아래 변수 추가:
   - **SMTP_PASSWORD** (필수): Gmail 앱 비밀번호. [앱 비밀번호 생성](https://myaccount.google.com/apppasswords)
   - **SMTP_USER** (선택): 발신 이메일, 기본 `jhleele2@gmail.com`
   - **TEAM_PASSWORD** (선택): 팀 로그인 비밀번호, 기본 `1234`
   - **SECRET_KEY** (선택): 세션 암호화용 랜덤 문자열
3. 변수 추가 후 **Redeploy** 한 번 실행해야 적용됩니다.
4. 배포 URL 접속 → 팀 비밀번호 입력 후 사용.

## 환경 변수 (.env)

| 변수 | 설명 | 기본값 |
|------|------|--------|
| SMTP_USER | 발신 이메일 | jhleele2@gmail.com |
| SMTP_PASSWORD | Gmail 앱 비밀번호 | (필수) |
| SMTP_HOST | SMTP 서버 | smtp.gmail.com |
| SMTP_PORT | SMTP 포트 | 587 |
| TEAM_PASSWORD | 팀 로그인 비밀번호 | 1234 |
| SECRET_KEY | 세션 암호화용 (배포 시 설정 권장) | (기본값 있음) |

## 엑셀 파일 구조

- **Suppliers**: 공급업체명, 담당자, 이메일, 리드타임(일), 품목
- **Inventory**: 품목코드, 이름, 규격, 단위, 현재고, 안전재고, MOQ, 공급업체, 담당자, 공급업체이메일, 리드타임, (발주수량·상태·발주메시지 등)
- **EmailTemplate**: 제목 형식, 본문 형식 (플레이스홀더: {{STORE_NAME}}, {{SUPPLIER_NAME}}, {{ORDER_DATE}}, {{ITEM_LIST}}, {{INTERNAL_OWNER}})
