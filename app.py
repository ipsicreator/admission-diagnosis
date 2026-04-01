import io
import json
import sqlite3
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st
from docx import Document
from pypdf import PdfReader


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "output"
DB_PATH = BASE_DIR / "admission_diag.db"

SUSI_2026 = DATA_DIR / "susi_explorer.csv"
SUSI_2027 = DATA_DIR / "susi_explorer_2027.csv"
CUTOFFS = DATA_DIR / "admission_cutoffs.csv"
CRITERIA = DATA_DIR / "holistic_criteria.csv"

TOP15 = [
    "서울대",
    "연세대",
    "고려대",
    "서강대",
    "성균관대",
    "한양대",
    "중앙대",
    "경희대",
    "한국외대",
    "서울시립대",
    "이화여대",
    "건국대",
    "동국대",
    "홍익대",
    "숙명여대",
]

GRADE_OPTIONS = ["고1", "고2", "고3", "N수이상"]
ADMISSION_TYPES = ["학생부교과", "학생부종합", "논술", "특기자/실기", "기타"]


@dataclass
class SupportChoice:
    support_no: int
    university: str
    department: str
    admission_type: str
    track_name: str
    diag_level: str
    diag_reason: str
    cutoff50: float
    cutoff70: float


def inject_css() -> None:
    st.markdown(
        """
        <style>
        .service-title {
          text-align: center;
          font-size: 33px; /* 25pt */
          font-weight: 800;
          margin-bottom: 2px;
        }
        .service-subtitle {
          text-align: center;
          font-size: 20px; /* 15pt */
          font-weight: 600;
          margin-bottom: 14px;
          opacity: 0.95;
        }
        .step-box {
          text-align: center;
          font-size: 31px;
          font-weight: 700;
          padding: 12px 4px;
          border-radius: 10px;
        }
        .step-active {
          background: rgba(30, 91, 255, 0.18);
          border: 1px solid rgba(84, 139, 255, 0.55);
        }
        .footer-note {
          text-align: center;
          margin-top: 22px;
          opacity: 0.85;
          font-size: 14px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def ensure_dirs() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def init_db() -> None:
    ensure_dirs()
    conn = sqlite3.connect(DB_PATH)
    try:
        cur = conn.cursor()
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS diagnosis_sessions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at TEXT NOT NULL,
                consultant_name TEXT,
                student_name TEXT,
                school_name TEXT,
                grade TEXT,
                student_phone TEXT,
                email TEXT,
                parent_phone TEXT,
                student_index REAL,
                pdf_summary TEXT,
                holistic_score REAL,
                outcome_json TEXT
            )
            """
        )
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS support_choices (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                session_id INTEGER NOT NULL,
                support_no INTEGER NOT NULL,
                university TEXT,
                department TEXT,
                admission_type TEXT,
                track_name TEXT,
                diag_level TEXT,
                diag_reason TEXT,
                cutoff50 REAL,
                cutoff70 REAL,
                FOREIGN KEY(session_id) REFERENCES diagnosis_sessions(id)
            )
            """
        )
        conn.commit()
    finally:
        conn.close()


def _normalize_columns(df: pd.DataFrame, kind: str) -> pd.DataFrame:
    alias_map = {
        "year": ["year", "연도", "학년도"],
        "university": ["university", "대학교", "대학명", "대학"],
        "department": ["department", "모집단위", "학과", "학부"],
        "admission_type": ["admission_type", "전형유형", "전형유형명"],
        "track_name": ["track_name", "전형명"],
        "percentile_type": ["percentile_type", "컷구분", "percentile"],
        "cutoff_score": ["cutoff_score", "컷점수", "점수", "score"],
    }
    rename_map: Dict[str, str] = {}
    cols = set(df.columns)
    for target, aliases in alias_map.items():
        for alias in aliases:
            if alias in cols:
                rename_map[alias] = target
                break
    df = df.rename(columns=rename_map)

    required = ["year", "university", "department", "admission_type", "track_name"]
    if kind == "cutoff":
        required += ["percentile_type", "cutoff_score"]

    for c in required:
        if c not in df.columns:
            df[c] = None

    return df[required].copy()


def load_csv(path: Path) -> pd.DataFrame:
    return pd.read_csv(path, encoding="utf-8-sig")


def save_uploaded_csv(uploaded_file, target_path: Path) -> None:
    target_path.write_bytes(uploaded_file.getvalue())


def _to_int_year(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(-1).astype(int)


def _clean_text(s: pd.Series) -> pd.Series:
    return s.fillna("").astype(str).str.strip()


def _univ_key(name: str) -> str:
    s = str(name or "").strip().replace(" ", "")
    for suffix in ["대학교", "대학", "대"]:
        if s.endswith(suffix):
            s = s[: -len(suffix)]
    return s


def _canonical_university_map(base_2026: pd.DataFrame) -> Dict[str, str]:
    names = sorted(base_2026["university"].dropna().astype(str).str.strip().unique().tolist())
    key_map: Dict[str, str] = {}
    for n in names:
        k = _univ_key(n)
        if not k:
            continue
        if k not in key_map:
            key_map[k] = n
        else:
            prev = key_map[k]
            # 2026 명칭 중 더 공식형(길이 긴 값) 우선
            key_map[k] = n if len(n) > len(prev) else prev
    return key_map


def _remove_excluded_type(admission_type: str, track_name: str) -> bool:
    joined = (str(admission_type) + " " + str(track_name)).replace(" ", "")
    return "교과기회" in joined


def load_data() -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    if not SUSI_2026.exists():
        pd.DataFrame(columns=["year", "university", "department", "admission_type", "track_name"]).to_csv(
            SUSI_2026, index=False, encoding="utf-8-sig"
        )
    if not SUSI_2027.exists():
        pd.DataFrame(columns=["year", "university", "department", "admission_type", "track_name"]).to_csv(
            SUSI_2027, index=False, encoding="utf-8-sig"
        )
    if not CUTOFFS.exists():
        pd.DataFrame(
            columns=["year", "university", "department", "admission_type", "track_name", "percentile_type", "cutoff_score"]
        ).to_csv(CUTOFFS, index=False, encoding="utf-8-sig")
    if not CRITERIA.exists():
        pd.DataFrame(columns=["university", "criterion", "weight", "description"]).to_csv(CRITERIA, index=False, encoding="utf-8-sig")

    s26 = _normalize_columns(load_csv(SUSI_2026), "susi")
    s27 = _normalize_columns(load_csv(SUSI_2027), "susi")
    s26["year"] = 2026
    s27["year"] = 2027

    for col in ["university", "department", "admission_type", "track_name"]:
        s26[col] = _clean_text(s26[col])
        s27[col] = _clean_text(s27[col])

    # 2027 대학명은 2026 기준 명칭으로 통일 (예: 강남대 -> 강남대학교)
    canon_map = _canonical_university_map(s26)
    s27["university"] = s27["university"].apply(
        lambda x: canon_map.get(_univ_key(x), str(x).strip())
    )

    merged = pd.concat([s26, s27], ignore_index=True)
    key = ["university", "department", "admission_type", "track_name"]
    merged = merged.sort_values("year")
    merged = merged.drop_duplicates(subset=key, keep="last")  # 충돌 시 2027 우선
    merged = merged[merged["university"] != ""].copy()
    merged = merged[~merged.apply(lambda r: _remove_excluded_type(r["admission_type"], r["track_name"]), axis=1)].copy()

    cutoff = _normalize_columns(load_csv(CUTOFFS), "cutoff")
    for col in ["university", "department", "admission_type", "track_name"]:
        cutoff[col] = _clean_text(cutoff[col])
    cutoff["university"] = cutoff["university"].apply(lambda x: canon_map.get(_univ_key(x), str(x).strip()))
    cutoff["year"] = _to_int_year(cutoff["year"])
    cutoff["percentile_type"] = pd.to_numeric(cutoff["percentile_type"], errors="coerce")
    cutoff["cutoff_score"] = pd.to_numeric(cutoff["cutoff_score"], errors="coerce")
    cutoff = cutoff[cutoff["year"].between(2023, 2025)].copy()  # 2023~2025 기준

    criteria = load_csv(CRITERIA)
    return merged, cutoff, criteria


def extract_pdf_text(uploaded_file) -> str:
    if not uploaded_file:
        return ""
    reader = PdfReader(io.BytesIO(uploaded_file.getvalue()))
    return "\n".join((page.extract_text() or "") for page in reader.pages)


def analyze_holistic_5level(pdf_text: str) -> Tuple[int, Dict[str, int], str]:
    text = (pdf_text or "").strip()
    if not text:
        detail = {
            "학업역량": 45,
            "전공적합성": 45,
            "자기주도성": 45,
            "공동체역량": 45,
            "발전가능성": 45,
        }
    else:
        base = min(100, len(text) // 60)
        keywords = {
            "학업역량": ["수업", "탐구", "학습", "과제", "교과"],
            "전공적합성": ["진로", "전공", "심화", "프로젝트", "연계"],
            "자기주도성": ["주도", "기획", "문제해결", "탐색", "개선"],
            "공동체역량": ["협업", "소통", "리더", "봉사", "배려"],
            "발전가능성": ["성장", "피드백", "도전", "확장", "변화"],
        }
        detail = {}
        for k, words in keywords.items():
            cnt = sum(text.count(w) for w in words)
            detail[k] = min(100, int(base * 0.5 + cnt * 9))

    avg = int(sum(detail.values()) / len(detail))
    if avg >= 85:
        level = 5
    elif avg >= 70:
        level = 4
    elif avg >= 55:
        level = 3
    elif avg >= 40:
        level = 2
    else:
        level = 1
    summary = f"학생부 분석 결과 5단계 중 {level}단계로 추정됩니다."
    return level, detail, summary


def get_cutoff_23_25(cutoffs: pd.DataFrame, uni: str, dept: str, atype: str, track: str) -> Tuple[float, float]:
    hit = cutoffs[
        (cutoffs["university"] == uni)
        & (cutoffs["department"] == dept)
        & (cutoffs["admission_type"] == atype)
        & (cutoffs["track_name"] == track)
    ]
    if hit.empty:
        return 85.0, 80.0

    c50 = hit.loc[hit["percentile_type"] == 50, "cutoff_score"].dropna()
    c70 = hit.loc[hit["percentile_type"] == 70, "cutoff_score"].dropna()

    cutoff50 = float(c50.median()) if not c50.empty else 85.0
    cutoff70 = float(c70.median()) if not c70.empty else 80.0
    return cutoff50, cutoff70


def rating_4level(student_index: float, cutoff50: float, cutoff70: float) -> Tuple[str, str]:
    if student_index >= cutoff50:
        return "상향 가능", f"지표({student_index:.1f})가 50% 컷({cutoff50:.1f}) 이상입니다."
    if student_index >= cutoff70:
        return "적정", f"지표({student_index:.1f})가 70% 컷({cutoff70:.1f}) 이상입니다."
    if student_index >= cutoff70 - 5:
        return "안전", f"지표({student_index:.1f})가 70% 컷({cutoff70:.1f})에 근접합니다."
    return "하향 권장", f"지표({student_index:.1f})가 70% 컷({cutoff70:.1f})보다 낮습니다."


def _basis_university(uni: str) -> str:
    return uni if uni in TOP15 else "경희대"


def build_report_text(payload: Dict, choices: List[SupportChoice], holistic_detail: Dict[str, int]) -> str:
    lines = [
        "# 나의 입시 위치 진단 종합보고서",
        f"- 생성시각: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        "",
        "## 1) 사용자 정보",
        f"- 컨설턴트: {payload.get('consultant_name', '')}",
        f"- 학생명: {payload.get('student_name', '')}",
        f"- 학교명: {payload.get('school_name', '')}",
        f"- 학년: {payload.get('grade', '')}",
        f"- 학생 연락처: {payload.get('student_phone', '')}",
        f"- 이메일: {payload.get('email', '')}",
        f"- 학부모 연락처: {payload.get('parent_phone', '')}",
        "",
        "## 2) 학생부 분석 (5단계)",
        f"- 종합단계: {payload.get('holistic_level', 0)}단계",
    ]
    for k, v in holistic_detail.items():
        lines.append(f"- {k}: {v}점")

    lines.extend(["", "## 3) 지원희망대학 평가 (최대 6개)"])
    for c in choices:
        basis = _basis_university(c.university)
        lines.append(
            f"- {c.support_no}. {c.university} / {c.department} / {c.admission_type} / {c.track_name}"
            f" -> {c.diag_level} (50컷 {c.cutoff50:.1f}, 70컷 {c.cutoff70:.1f}, 기준대학 {basis})"
        )
        lines.append(f"  - 판단근거: {c.diag_reason}")

    lines.extend(
        [
            "",
            "## 4) 적용 기준",
            "- 대학 리스트: 2026 수시검색기 + 2027 업로드 병합",
            "- 충돌 규칙: 동일 조합(대학교/모집단위/전형유형/전형명) 충돌 시 2027 우선",
            "- 입결 판단: 2023~2025학년도 50%/70% 컷 기준",
            "- 학생부종합 평가기준: 서울 상위 15개 대학 기준, 미포함 대학은 경희대 기준 적용",
        ]
    )
    return "\n".join(lines)


def build_docx_bytes(payload: Dict, choices: List[SupportChoice], holistic_detail: Dict[str, int]) -> bytes:
    doc = Document()
    doc.add_heading("나의 입시 위치 진단 종합보고서", level=1)
    doc.add_paragraph(f"생성시각: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    doc.add_heading("1) 사용자 정보", level=2)
    info_rows = [
        ("컨설턴트", payload.get("consultant_name", "")),
        ("학생명", payload.get("student_name", "")),
        ("학교명", payload.get("school_name", "")),
        ("학년", payload.get("grade", "")),
        ("학생 연락처", payload.get("student_phone", "")),
        ("이메일", payload.get("email", "")),
        ("학부모 연락처", payload.get("parent_phone", "")),
    ]
    for k, v in info_rows:
        doc.add_paragraph(f"- {k}: {v}")

    doc.add_heading("2) 학생부 분석 (5단계)", level=2)
    doc.add_paragraph(f"- 종합단계: {payload.get('holistic_level', 0)}단계")
    for k, v in holistic_detail.items():
        doc.add_paragraph(f"- {k}: {v}점")

    doc.add_heading("3) 지원희망대학 평가", level=2)
    table = doc.add_table(rows=1, cols=7)
    hdr = table.rows[0].cells
    hdr[0].text = "번호"
    hdr[1].text = "대학/모집단위"
    hdr[2].text = "전형"
    hdr[3].text = "평가"
    hdr[4].text = "50컷"
    hdr[5].text = "70컷"
    hdr[6].text = "판단근거"

    for c in choices:
        row = table.add_row().cells
        row[0].text = str(c.support_no)
        row[1].text = f"{c.university} / {c.department}"
        row[2].text = f"{c.admission_type} / {c.track_name}"
        row[3].text = c.diag_level
        row[4].text = f"{c.cutoff50:.1f}"
        row[5].text = f"{c.cutoff70:.1f}"
        row[6].text = c.diag_reason

    doc.add_heading("4) 적용 기준", level=2)
    doc.add_paragraph("- 대학 리스트: 2026 수시검색기 + 2027 업로드 병합 (충돌 시 2027 우선)")
    doc.add_paragraph("- 입결 판단: 2023~2025학년도 50%/70% 컷 기준")

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


def save_session(payload: Dict, choices: List[SupportChoice]) -> None:
    conn = sqlite3.connect(DB_PATH)
    try:
        cur = conn.cursor()
        outcome = {"payload": payload, "supports": [c.__dict__ for c in choices]}
        cur.execute(
            """
            INSERT INTO diagnosis_sessions (
                created_at, consultant_name, student_name, school_name, grade,
                student_phone, email, parent_phone, student_index, pdf_summary,
                holistic_score, outcome_json
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                datetime.now().isoformat(timespec="seconds"),
                payload.get("consultant_name", ""),
                payload.get("student_name", ""),
                payload.get("school_name", ""),
                payload.get("grade", ""),
                payload.get("student_phone", ""),
                payload.get("email", ""),
                payload.get("parent_phone", ""),
                float(payload.get("student_index", 0)),
                payload.get("pdf_summary", ""),
                float(payload.get("holistic_score", 0)),
                json.dumps(outcome, ensure_ascii=False),
            ),
        )
        session_id = int(cur.lastrowid)
        for c in choices:
            cur.execute(
                """
                INSERT INTO support_choices (
                    session_id, support_no, university, department, admission_type,
                    track_name, diag_level, diag_reason, cutoff50, cutoff70
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    session_id,
                    c.support_no,
                    c.university,
                    c.department,
                    c.admission_type,
                    c.track_name,
                    c.diag_level,
                    c.diag_reason,
                    c.cutoff50,
                    c.cutoff70,
                ),
            )
        conn.commit()
    finally:
        conn.close()


def _step_header(current_step: int) -> None:
    labels = ["1. 사용자정보", "2. 학생부 분석", "3. 희망대학 지원", "4. 종합 보고서"]
    cols = st.columns(4)
    for i, (col, label) in enumerate(zip(cols, labels), start=1):
        with col:
            cls = "step-box step-active" if i == current_step else "step-box"
            st.markdown(f"<div class='{cls}'>{label}</div>", unsafe_allow_html=True)


def _consultant_panel() -> str:
    if "consultant_name_fixed" not in st.session_state:
        st.session_state.consultant_name_fixed = ""

    if st.session_state.consultant_name_fixed:
        st.text_input("컨설턴트 이름", value=st.session_state.consultant_name_fixed, disabled=True, key="consultant_fixed_view")
        return st.session_state.consultant_name_fixed

    c1, c2 = st.columns([3, 1])
    with c1:
        name = st.text_input("컨설턴트 이름", key="consultant_name_once")
    with c2:
        save = st.button("저장", use_container_width=True)
    if save and name.strip():
        st.session_state.consultant_name_fixed = name.strip()
        st.rerun()
    return name.strip()


def _choice_input_block(susi_df: pd.DataFrame, no: int) -> Tuple[str, str, str, str]:
    st.markdown(f"지원 {no}")
    a, b, c, d = st.columns(4)
    uni_options = sorted(susi_df["university"].dropna().unique().tolist())

    with a:
        uni = st.selectbox(f"대학교 {no}", ["직접입력"] + uni_options, key=f"uni_{no}")
        uni_text = st.text_input(f"대학교 직접입력 {no}", key=f"uni_txt_{no}") if uni == "직접입력" else uni

    filtered_uni = susi_df[susi_df["university"] == uni_text] if uni_text else susi_df
    dept_options = sorted(filtered_uni["department"].dropna().unique().tolist()) or [""]

    with b:
        dept = st.selectbox(f"모집단위 {no}", ["직접입력"] + dept_options, key=f"dept_{no}")
        dept_text = st.text_input(f"모집단위 직접입력 {no}", key=f"dept_txt_{no}") if dept == "직접입력" else dept

    filtered_dept = filtered_uni[filtered_uni["department"] == dept_text] if dept_text else filtered_uni
    atype_options = sorted(filtered_dept["admission_type"].dropna().unique().tolist())
    track_options = sorted(filtered_dept["track_name"].dropna().unique().tolist()) or [""]

    with c:
        atype = st.selectbox(f"전형유형 {no}", atype_options or ADMISSION_TYPES, key=f"atype_{no}")
    with d:
        track = st.selectbox(f"전형명 {no}", ["직접입력"] + track_options, key=f"track_{no}")
        track_text = st.text_input(f"전형명 직접입력 {no}", key=f"track_txt_{no}") if track == "직접입력" else track

    return uni_text.strip(), dept_text.strip(), atype.strip(), track_text.strip()


def main() -> None:
    st.set_page_config(page_title="나의 입시 위치 진단 서비스", layout="wide")
    init_db()
    inject_css()

    st.markdown("<div class='service-title'>나의 입시 위치 진단서비스</div>", unsafe_allow_html=True)
    st.markdown("<div class='service-subtitle'>by 대치수프리마</div>", unsafe_allow_html=True)
    st.caption("기준: 2026+2027 수시 병합(충돌 시 2027 우선), 2023~2025 50/70컷")

    # 데이터 업로드 관리
    st.subheader("데이터 관리")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.download_button("2026 수시 템플릿", SUSI_2026.read_bytes() if SUSI_2026.exists() else b"", file_name="susi_explorer.csv")
        up1 = st.file_uploader("2026 수시 CSV 업로드", type=["csv"], key="up_s26")
        if up1:
            save_uploaded_csv(up1, SUSI_2026)
            st.success("2026 수시 데이터 반영 완료")
    with c2:
        st.download_button("2027 수시 템플릿", SUSI_2027.read_bytes() if SUSI_2027.exists() else b"", file_name="susi_explorer_2027.csv")
        up2 = st.file_uploader("2027 수시 CSV 업로드", type=["csv"], key="up_s27")
        if up2:
            save_uploaded_csv(up2, SUSI_2027)
            st.success("2027 수시 데이터 반영 완료")
    with c3:
        st.download_button("컷 템플릿", CUTOFFS.read_bytes() if CUTOFFS.exists() else b"", file_name="admission_cutoffs.csv")
        up3 = st.file_uploader("컷 CSV 업로드(2023~2025)", type=["csv"], key="up_cut")
        if up3:
            save_uploaded_csv(up3, CUTOFFS)
            st.success("컷 데이터 반영 완료")
    with c4:
        st.download_button("학생부 평가기준 템플릿", CRITERIA.read_bytes() if CRITERIA.exists() else b"", file_name="holistic_criteria.csv")
        up4 = st.file_uploader("평가기준 CSV 업로드", type=["csv"], key="up_criteria")
        if up4:
            save_uploaded_csv(up4, CRITERIA)
            st.success("평가기준 데이터 반영 완료")

    susi_df, cutoff_df, _criteria_df = load_data()
    st.caption(f"대학 목록 데이터: {len(susi_df)}건 | 컷 데이터(2023~2025): {len(cutoff_df)}건")
    if cutoff_df.empty:
        st.warning("컷 데이터에 2023~2025 연도가 없습니다. 기본 컷(50%=85, 70%=80)으로 판단됩니다.")

    if "step" not in st.session_state:
        st.session_state.step = 1
    if "profile" not in st.session_state:
        st.session_state.profile = {}
    if "holistic" not in st.session_state:
        st.session_state.holistic = {}
    if "supports" not in st.session_state:
        st.session_state.supports = []
    if "saved" not in st.session_state:
        st.session_state.saved = False

    head_left, head_right = st.columns([6, 2])
    with head_left:
        _step_header(st.session_state.step)
    with head_right:
        consultant_global = _consultant_panel()
    st.markdown("---")

    # Step 1
    if st.session_state.step == 1:
        st.subheader("1단계 - 사용자정보")
        with st.form("step1_form"):
            c1, c2 = st.columns(2)
            with c1:
                student_name = st.text_input("학생명 *")
                school_name = st.text_input("학교명 *")
                grade = st.selectbox("학년 *", GRADE_OPTIONS)
            with c2:
                student_phone = st.text_input("학생 전화번호 *")
                email = st.text_input("메일주소(필수) *")
                parent_phone = st.text_input("학부모 연락처 *")
                student_index = st.slider("학생 지표(내신/활동 종합 지표)", min_value=50.0, max_value=100.0, value=82.0, step=0.5)

            ok = st.form_submit_button("다음 단계")

        if ok:
            required = {
                "컨설턴트 이름": consultant_global,
                "학생명": student_name,
                "학교명": school_name,
                "학생 전화번호": student_phone,
                "메일주소": email,
                "학부모 연락처": parent_phone,
            }
            missing = [k for k, v in required.items() if not str(v).strip()]
            if missing:
                st.error("필수값 누락: " + ", ".join(missing))
            else:
                st.session_state.profile = {
                    "consultant_name": consultant_global.strip(),
                    "student_name": student_name.strip(),
                    "school_name": school_name.strip(),
                    "grade": grade.strip(),
                    "student_phone": student_phone.strip(),
                    "email": email.strip(),
                    "parent_phone": parent_phone.strip(),
                    "student_index": float(student_index),
                }
                st.session_state.step = 2
                st.rerun()

    # Step 2
    elif st.session_state.step == 2:
        st.subheader("2단계 - 학생부 분석")
        pdf_file = st.file_uploader("학생부 PDF 업로드", type=["pdf"], key="pdf_step2")
        c1, c2, c3 = st.columns(3)
        with c1:
            if st.button("이전 단계"):
                st.session_state.step = 1
                st.rerun()
        with c2:
            if st.button("분석 실행", use_container_width=True):
                text = extract_pdf_text(pdf_file) if pdf_file else ""
                level, detail, summary = analyze_holistic_5level(text)
                st.session_state.holistic = {
                    "pdf_summary": summary,
                    "holistic_level": level,
                    "holistic_detail": detail,
                    "holistic_score": sum(detail.values()) / len(detail),
                }
        with c3:
            if st.button("다음 단계", use_container_width=True):
                if not st.session_state.holistic:
                    st.warning("먼저 분석 실행을 눌러 주세요.")
                else:
                    st.session_state.step = 3
                    st.rerun()

        if st.session_state.holistic:
            st.info(st.session_state.holistic.get("pdf_summary", ""))
            detail = st.session_state.holistic.get("holistic_detail", {})
            m1, m2 = st.columns(2)
            with m1:
                st.metric("종합 단계", f"{st.session_state.holistic.get('holistic_level', 0)}단계")
            with m2:
                st.metric("평균 점수", f"{st.session_state.holistic.get('holistic_score', 0):.1f}")
            for k, v in detail.items():
                st.write(f"- {k}: {v}점")
                st.progress(min(max(int(v), 0), 100))

    # Step 3
    elif st.session_state.step == 3:
        st.subheader("3단계 - 희망대학 지원 (최대 6개)")
        supports: List[Tuple[str, str, str, str]] = []

        with st.form("step3_form"):
            for no in range(1, 7):
                with st.expander(f"지원 {no}", expanded=(no <= 2)):
                    supports.append(_choice_input_block(susi_df, no))

            c1, c2 = st.columns(2)
            with c1:
                back = st.form_submit_button("이전 단계")
            with c2:
                nxt = st.form_submit_button("평가 후 보고서 보기", use_container_width=True)

        if back:
            st.session_state.step = 2
            st.rerun()

        if nxt:
            rows: List[SupportChoice] = []
            student_index = float(st.session_state.profile.get("student_index", 0))
            for i, (uni, dept, atype, track) in enumerate(supports, start=1):
                if not (uni and dept and atype and track):
                    continue
                c50, c70 = get_cutoff_23_25(cutoff_df, uni, dept, atype, track)
                level, reason = rating_4level(student_index, c50, c70)
                rows.append(SupportChoice(i, uni, dept, atype, track, level, reason, c50, c70))

            if not rows:
                st.warning("최소 1개 이상의 지원 대학/학과/전형 정보를 입력해 주세요.")
            else:
                st.session_state.supports = [r.__dict__ for r in rows]
                st.session_state.step = 4
                st.rerun()

    # Step 4
    else:
        st.subheader("4단계 - 종합 보고서")

        payload = {
            **st.session_state.profile,
            **st.session_state.holistic,
            "top15_reference": TOP15,
        }
        choices = [SupportChoice(**x) for x in st.session_state.supports]
        holistic_detail = st.session_state.holistic.get("holistic_detail", {})

        if not st.session_state.saved:
            save_session(payload, choices)
            st.session_state.saved = True

        report_text = build_report_text(payload, choices, holistic_detail)
        st.text_area("종합 보고서 미리보기", value=report_text, height=420)

        docx_bytes = build_docx_bytes(payload, choices, holistic_detail)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        st.download_button(
            "DOC 보고서 다운로드",
            data=docx_bytes,
            file_name=f"admission_report_{stamp}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

        c1, c2 = st.columns(2)
        with c1:
            if st.button("이전 단계"):
                st.session_state.step = 3
                st.rerun()
        with c2:
            if st.button("새 진단 시작", use_container_width=True):
                st.session_state.step = 1
                st.session_state.profile = {}
                st.session_state.holistic = {}
                st.session_state.supports = []
                st.session_state.saved = False
                st.rerun()

        st.info(f"진단 결과가 DB에 저장되었습니다: {DB_PATH}")

    st.markdown("<div class='footer-note'>자료: 대학어디가  운영: 대치 수프리마</div>", unsafe_allow_html=True)


if __name__ == "__main__":
    main()
