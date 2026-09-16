# ☀️ SUNNY PHARM 365 — GitHub Preview

SUNNY 약사님용 Streamlit 업무 앱 프리뷰입니다. 이 저장소 버전은 **사용자가 제공한 SUNNY Excel의 안전한 스냅샷**을 `data/SUNNY.xlsx`에서 직접 읽습니다.

## 포함된 기능
- 오늘의 SUNNY 대시보드
- 입고관리
- 주문센터: 주문해야 하는 약 / OWM / 쿠팡 / 직거래
- 상품 · 거래처 빠른 검색
- 인플루언서 일정
- 세트상품
- 대체조제 확인
- SUNNY NOTE
- 원본 시트 탐색

## 데이터 보안
- 원본의 `홈페이지 아이디, 비번` 시트는 **이 저장소용 Excel에서 제거했습니다.**
- `❤️ HUB❤️` 안의 비밀번호/접속코드 관련 행도 저장소용 Excel에서 제거했습니다.
- 그래도 업무 데이터가 들어 있으므로 **GitHub Repository는 Private 권장**입니다.
- 이후 Google Sheets 권한을 받으면 `data/SUNNY.xlsx` 대신 Google Sheets를 데이터 소스로 바꾸면 됩니다.

## 로컬 실행
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Community Cloud 배포
1. 이 폴더 안의 파일을 GitHub 저장소 루트에 업로드합니다.
2. Streamlit Community Cloud에서 `New app`을 선택합니다.
3. Repository: 방금 만든 저장소 / Branch: `main` / Main file: `app.py`
4. Deploy 합니다.

### 선택: 앱에 비밀번호 걸기
Streamlit Cloud → App settings → Secrets에 아래처럼 넣으면 프리뷰 진입 비밀번호가 생깁니다.

```toml
APP_PASSWORD = "원하는비밀번호"
```

이 값은 GitHub 파일에 적지 마세요.

## 폴더 구조
```text
app.py
requirements.txt
.gitignore
.streamlit/config.toml
assets/sunny_profile.png
data/SUNNY.xlsx          # 민감 시트 제거된 GitHub 프리뷰용 스냅샷
services/excel_source.py
```
