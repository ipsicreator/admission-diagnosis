# 나의 입시 위치 진단 서비스 (MVP)

Streamlit 기반 입시 진단 앱입니다.

- 사용자 정보 입력
- 수시 6개 지원 입력
- 학생부 PDF 분석(5단계)
- 50%/70% 컷 기반 지원별 4단계 진단
- 최종 보고서(Markdown/PDF) 생성
- 진단 결과 SQLite DB 누적 저장
- CSV 업로드로 데이터 교체/수정 가능

## 1) 로컬 실행

```powershell
cd "C:\Users\chris\Desktop\입시위치진단서비스"
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

## 2) GitHub 연동

```powershell
cd "C:\Users\chris\Desktop\입시위치진단서비스"
git init
git add .
git commit -m "feat: initial admissions diagnosis streamlit app"
```

원격 저장소 생성 후:

```powershell
git remote add origin <YOUR_GITHUB_REPO_URL>
git branch -M main
git push -u origin main
```

## 3) 재부팅/PC 전원과 무관하게 계속 운영 (추천)

### Streamlit Community Cloud

1. GitHub에 코드 push
2. Streamlit Cloud 로그인
3. New app -> 리포지토리 선택
4. Main file path: `app.py`
5. Deploy

이 방식은 로컬 PC가 꺼져도 서비스가 계속 동작합니다.

## 4) 로컬 PC에서 자동 재시작 운영 (재부팅 후 자동복구)

### NSSM 기반 Windows 서비스

1. NSSM 다운로드 후 `C:\tools\nssm\nssm.exe` 배치
2. 아래 실행:

```powershell
cd "C:\Users\chris\Desktop\입시위치진단서비스"
.\scripts\install_service.ps1
```

서비스 제거:

```powershell
.\scripts\remove_service.ps1
```

주의: 이 방식은 PC 전원이 꺼지면 서비스도 꺼집니다.

## 5) 데이터 교체 방식

앱 상단 `데이터 관리`에서 아래 CSV를 업로드하면 즉시 반영됩니다.

- `data/susi_explorer.csv`
- `data/admission_cutoffs.csv`
- `data/holistic_criteria.csv`

원천데이터(구글드라이브)는 CSV로 내보낸 뒤 업로드하면 됩니다.

## 6) 파일 구조

- `app.py`: 메인 앱
- `requirements.txt`: 패키지 목록
- `data/`: 교체형 CSV 데이터
- `scripts/install_service.ps1`: 서비스 설치
- `scripts/remove_service.ps1`: 서비스 제거
