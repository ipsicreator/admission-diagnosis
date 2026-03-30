import io
import json
import sqlite3
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st
from pypdf import PdfReader
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "output"
DB_PATH = BASE_DIR / "admission_diag.db"

SUSI_TEMPLATE = DATA_DIR / "susi_explorer.csv"
CUTOFF_TEMPLATE = DATA_DIR / "admission_cutoffs.csv"
CRITERIA_TEMPLATE = DATA_DIR / "holistic_criteria.csv"
SUSI_2027 = DATA_DIR / "susi_explorer_2027.csv"
CUTOFF_2027 = DATA_DIR / "admission_cutoffs_2027.csv"

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
ADMISSION_TYPES = ["학생부교과", "학생부종합", "논술", "실기/실적", "기타"]


@dataclass
class SupportChoice:
    university: str
    department: str
    admission_type: str
    track_name: str
    diag_level: str
    diag_reason: str


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
                FOREIGN KEY(session_id) REFERENCES diagnosis_sessions(id)
            )
            """
        )
        cur.execute("CREATE INDEX IF NOT EXISTS idx_sessions_created_at ON diagnosis_sessions(created_at)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_choices_session_id ON support_choices(session_id)")
        conn.commit()
    finally:
        conn.close()


def seed_templates_if_missing() -> None:
    if not SUSI_TEMPLATE.exists():
        pd.DataFrame(
            [
                {"year": 2026, "university": "서울대", "department": "컴퓨터공학부", "admission_type": "학생부종합", "track_name": "일반전형"},
                {"year": 2026, "university": "연세대", "department": "컴퓨터과학과", "admission_type": "학생부종합", "track_name": "활동우수형"},
                {"year": 2026, "university": "고려대", "department": "컴퓨터학과", "admission_type": "학생부교과", "track_name": "학교추천"},
            ]
        ).to_csv(SUSI_TEMPLATE, index=False, encoding="utf-8-sig")
    if not CUTOFF_TEMPLATE.exists():
        pd.DataFrame(
            [
                {"year": 2026, "university": "서울대", "department": "컴퓨터공학부", "admission_type": "학생부종합", "track_name": "일반전형", "percentile_type": 50, "cutoff_score": 87},
                {"year": 2026, "university": "서울대", "department": "컴퓨터공학부", "admission_type": "학생부종합", "track_name": "일반전형", "percentile_type": 70, "cutoff_score": 83},
                {"year": 2026, "university": "연세대", "department": "컴퓨터과학과", "admission_type": "학생부종합", "track_name": "활동우수형", "percentile_type": 50, "cutoff_score": 85},
                {"year": 2026, "university": "연세대", "department": "컴퓨터과학과", "admission_type": "학생부종합", "track_name": "활동우수형", "percentile_type": 70, "cutoff_score": 80},
            ]
        ).to_csv(CUTOFF_TEMPLATE, index=False, encoding="utf-8-sig")
    if not CRITERIA_TEMPLATE.exists():
        rows = []
        for uni in ["서울대", "연세대", "고려대"]:
            rows.extend(
                [
                    {"university": uni, "criterion": "학업역량", "weight": 0.3, "description": "교과 성취와 심화학습의 일관성"},
                    {"university": uni, "criterion": "전공적합성", "weight": 0.25, "description": "전공 연계 선택과목 및 활동"},
                    {"university": uni, "criterion": "자기주도성", "weight": 0.2, "description": "탐구의 주도성과 지속성"},
                    {"university": uni, "criterion": "공동체역량", "weight": 0.15, "description": "협업 및 소통 경험"},
                    {"university": uni, "criterion": "발전가능성", "weight": 0.1, "description": "학년별 성장 맥락"},
                ]
            )
        pd.DataFrame(rows).to_csv(CRITERIA_TEMPLATE, index=False, encoding="utf-8-sig")


def load_csv(path: Path) -> pd.DataFrame:
    return pd.read_csv(path, encoding="utf-8-sig")


def _to_int_year(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(-1).astype(int)


def _collect_2027_universities(base_susi: pd.DataFrame, base_cutoff: pd.DataFrame) -> List[str]:
    universities = set()
    if "year" in base_susi.columns and "university" in base_susi.columns:
        y = _to_int_year(base_susi["year"])
        universities.update(base_susi.loc[y == 2027, "university"].dropna().astype(str).tolist())
    if "year" in base_cutoff.columns and "university" in base_cutoff.columns:
        y = _to_int_year(base_cutoff["year"])
        universities.update(base_cutoff.loc[y == 2027, "university"].dropna().astype(str).tolist())

    for path in [SUSI_2027, CUTOFF_2027]:
        if path.exists():
            df = load_csv(path)
            if "university" in df.columns:
                universities.update(df["university"].dropna().astype(str).tolist())
    return sorted(u.strip() for u in universities if str(u).strip())


def build_2026_only_frames(base_susi: pd.DataFrame, base_cutoff: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame, List[str]]:
    susi_df = base_susi.copy()
    cutoff_df = base_cutoff.copy()

    if "year" in susi_df.columns:
        susi_df = susi_df.loc[_to_int_year(susi_df["year"]) == 2026].copy()
    if "year" in cutoff_df.columns:
        cutoff_df = cutoff_df.loc[_to_int_year(cutoff_df["year"]) == 2026].copy()

    excluded = _collect_2027_universities(base_susi, base_cutoff)
    if excluded and "university" in susi_df.columns:
        susi_df = susi_df[~susi_df["university"].astype(str).isin(excluded)].copy()
    if excluded and "university" in cutoff_df.columns:
        cutoff_df = cutoff_df[~cutoff_df["university"].astype(str).isin(excluded)].copy()

    return susi_df, cutoff_df, excluded


def extract_pdf_text(uploaded_file) -> str:
    if not uploaded_file:
        return ""
    reader = PdfReader(io.BytesIO(uploaded_file.getvalue()))
    texts = []
    for page in reader.pages:
        texts.append(page.extract_text() or "")
    return "\n".join(texts)


def analyze_holistic_5level(pdf_text: str) -> Tuple[int, Dict[str, int], str]:
    text = pdf_text.strip()
    length_score = min(100, len(text) // 50) if text else 0
    keywords = {
        "학업역량": ["세부능력", "교과", "성취", "탐구"],
        "전공적합성": ["진로", "전공", "심화", "프로젝트"],
        "자기주도성": ["주도", "기획", "문제해결", "탐구"],
        "공동체역량": ["협력", "소통", "리더", "봉사"],
        "발전가능성": ["성장", "피드백", "개선", "도전"],
    }
    per_item: Dict[str, int] = {}
    for k, words in keywords.items():
        cnt = sum(text.count(w) for w in words)
        per_item[k] = min(100, cnt * 8 + int(length_score * 0.3))

    avg = int(sum(per_item.values()) / len(per_item)) if per_item else 0
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

    summary = f"학생부 정성평가(서울 상위권 15개 대학 학생부종합 기준) 5단계 중 {level}단계로 추정됩니다."
    return level, per_item, summary


def rating_4level(student_index: float, cutoff50: float, cutoff70: float) -> Tuple[str, str]:
    if student_index >= cutoff50:
        return "상향 가능", f"지원자 지표({student_index:.1f})가 50% 컷({cutoff50:.1f}) 이상입니다."
    if student_index >= cutoff70:
        return "적정", f"지원자 지표({student_index:.1f})가 70% 컷({cutoff70:.1f}) 이상입니다."
    if student_index >= cutoff70 - 5:
        return "도전", f"지원자 지표({student_index:.1f})가 70% 컷({cutoff70:.1f})에 근접합니다."
    return "하향 권장", f"지원자 지표({student_index:.1f})가 70% 컷({cutoff70:.1f}) 대비 낮습니다."


def get_cutoff(cutoffs: pd.DataFrame, uni: str, dept: str, atype: str, track: str) -> Tuple[float, float]:
    hit = cutoffs[
        (cutoffs["university"] == uni)
        & (cutoffs["department"] == dept)
        & (cutoffs["admission_type"] == atype)
        & (cutoffs["track_name"] == track)
    ]
    if hit.empty:
        return 85.0, 80.0
    c50 = hit.loc[hit["percentile_type"] == 50, "cutoff_score"]
    c70 = hit.loc[hit["percentile_type"] == 70, "cutoff_score"]
    return float(c50.iloc[0] if not c50.empty else 85.0), float(c70.iloc[0] if not c70.empty else 80.0)


def save_session(payload: Dict, choices: List[SupportChoice]) -> None:
    conn = sqlite3.connect(DB_PATH)
    try:
        cur = conn.cursor()
        payload_out = dict(payload)
        payload_out["supports"] = [c.__dict__ for c in choices]
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
                json.dumps(payload_out, ensure_ascii=False),
            ),
        )
        session_id = int(cur.lastrowid)
        for i, row in enumerate(choices, start=1):
            cur.execute(
                """
                INSERT INTO support_choices (
                    session_id, support_no, university, department, admission_type,
                    track_name, diag_level, diag_reason
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    session_id,
                    i,
                    row.university,
                    row.department,
                    row.admission_type,
                    row.track_name,
                    row.diag_level,
                    row.diag_reason,
                ),
            )
        conn.commit()
    finally:
        conn.close()


def report_markdown(payload: Dict, choices: List[SupportChoice], holistic_detail: Dict[str, int]) -> str:
    lines = [
        "# 나의 입시 위치 진단 보고서",
        f"- 생성시각: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        "",
        "## 1) 사용자 정보",
        f"- 컨설턴트: {payload['consultant_name']}",
        f"- 학생명: {payload['student_name']}",
        f"- 학교명: {payload['school_name']}",
        f"- 학년: {payload['grade']}",
        f"- 학생 연락처: {payload['student_phone']}",
        f"- 이메일: {payload['email']}",
        f"- 학부모 연락처: {payload['parent_phone']}",
        "",
        "## 2) 학생부 분석(5단계)",
        f"- 종합 단계: {payload['holistic_level']}단계",
    ]
    for k, v in holistic_detail.items():
        lines.append(f"- {k}: {v}점")
    lines.extend(
        [
            "",
            "## 3) 수시 6개 지원 진단(4단계)",
        ]
    )
    for i, c in enumerate(choices, start=1):
        lines.append(
            f"- {i}. {c.university} / {c.department} / {c.admission_type} / {c.track_name} -> {c.diag_level} ({c.diag_reason})"
        )
    lines.extend(
        [
            "",
            "## 4) 데이터 기준",
            "- 대학어디가 50%, 70% 컷 기반(보유 CSV 기준)",
            "- 서울 상위권 15개 대학 학생부종합 평가기준 준용",
            "- 2023~2026 데이터는 CSV 교체 업로드로 반영 가능",
        ]
    )
    return "\n".join(lines)


def markdown_to_pdf_bytes(text: str) -> bytes:
    pdfmetrics.registerFont(UnicodeCIDFont("HYSMyeongJo-Medium"))
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    title = ParagraphStyle("title", parent=styles["Title"], fontName="HYSMyeongJo-Medium", fontSize=18)
    body = ParagraphStyle("body", parent=styles["BodyText"], fontName="HYSMyeongJo-Medium", fontSize=10, leading=14)

    story = []
    for line in text.splitlines():
        if line.startswith("# "):
            story.append(Paragraph(line[2:], title))
        else:
            safe = line.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            story.append(Paragraph(safe if safe else " ", body))
        story.append(Spacer(1, 4))
    doc.build(story)
    buf.seek(0)
    return buf.read()


def main() -> None:
    st.set_page_config(page_title="나의 입시 위치 진단 서비스", layout="wide")
    init_db()
    seed_templates_if_missing()

    st.title("나의 입시 위치 진단 서비스")
    st.caption("Streamlit + GitHub 기반 MVP | CSV 교체형 데이터 운영")

    st.subheader("데이터 관리")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.download_button("수시탐색기 템플릿 다운로드", SUSI_TEMPLATE.read_bytes(), file_name="susi_explorer.csv")
        up = st.file_uploader("수시탐색기 CSV 업로드", type=["csv"], key="u1")
        if up:
            SUSI_TEMPLATE.write_bytes(up.getvalue())
            st.success("수시탐색기 CSV 반영 완료")
    with c2:
        st.download_button("50/70 컷 템플릿 다운로드", CUTOFF_TEMPLATE.read_bytes(), file_name="admission_cutoffs.csv")
        up2 = st.file_uploader("컷 CSV 업로드", type=["csv"], key="u2")
        if up2:
            CUTOFF_TEMPLATE.write_bytes(up2.getvalue())
            st.success("컷 CSV 반영 완료")
    with c3:
        st.download_button("학생부 평가기준 템플릿 다운로드", CRITERIA_TEMPLATE.read_bytes(), file_name="holistic_criteria.csv")
        up3 = st.file_uploader("평가기준 CSV 업로드", type=["csv"], key="u3")
        if up3:
            CRITERIA_TEMPLATE.write_bytes(up3.getvalue())
            st.success("평가기준 CSV 반영 완료")

    raw_susi_df = load_csv(SUSI_TEMPLATE)
    raw_cutoff_df = load_csv(CUTOFF_TEMPLATE)
    susi_df, cutoff_df, excluded_2027 = build_2026_only_frames(raw_susi_df, raw_cutoff_df)
    _ = load_csv(CRITERIA_TEMPLATE)

    st.caption(
        f"현재 운영데이터: 2026 기준 {len(susi_df)}건 / "
        f"2027 대학 제외 {len(excluded_2027)}개"
    )
    if excluded_2027:
        st.caption("제외 대학: " + ", ".join(excluded_2027))

    st.markdown("---")
    st.subheader("1) 진단검사 사용자 정보")
    with st.form("diag_form"):
        col1, col2 = st.columns(2)
        with col1:
            consultant_name = st.text_input("컨설턴트 이름")
            student_name = st.text_input("학생명")
            school_name = st.text_input("학교명")
            grade = st.selectbox("학년", GRADE_OPTIONS)
        with col2:
            student_phone = st.text_input("학생 전화번호")
            email = st.text_input("메일주소(필수)")
            parent_phone = st.text_input("학부모 연락처")
            student_index = st.slider("학생 지표(내신/활동 종합 환산 예시)", min_value=50.0, max_value=100.0, value=82.0, step=0.5)

        st.subheader("2) 수시 6개 지원 입력")
        uni_options = sorted(susi_df["university"].dropna().unique().tolist())
        support_inputs = []
        for i in range(1, 7):
            st.markdown(f"지원 {i}")
            a, b, c, d = st.columns(4)
            with a:
                uni = st.selectbox(f"학교선택 {i}", ["직접입력"] + uni_options, key=f"uni_{i}")
                uni_text = st.text_input(f"학교 직접입력 {i}", key=f"uni_txt_{i}") if uni == "직접입력" else uni
            filtered = susi_df[susi_df["university"] == uni_text] if uni_text else susi_df
            dept_options = sorted(filtered["department"].dropna().unique().tolist()) or [""]
            track_options = sorted(filtered["track_name"].dropna().unique().tolist()) or [""]
            with b:
                dept = st.selectbox(f"학과 {i}", ["직접입력"] + dept_options, key=f"dept_{i}")
                dept_text = st.text_input(f"학과 직접입력 {i}", key=f"dept_txt_{i}") if dept == "직접입력" else dept
            with c:
                atype = st.selectbox(f"전형유형 {i}", ADMISSION_TYPES, key=f"atype_{i}")
            with d:
                track = st.selectbox(f"전형명 {i}", ["직접입력"] + track_options, key=f"track_{i}")
                track_text = st.text_input(f"전형명 직접입력 {i}", key=f"track_txt_{i}") if track == "직접입력" else track

            support_inputs.append((uni_text.strip(), dept_text.strip(), atype.strip(), track_text.strip()))

        st.subheader("3) 학생부 분석")
        pdf_file = st.file_uploader("학생부 PDF 업로드", type=["pdf"])
        submit = st.form_submit_button("진단 실행")

    if not submit:
        st.stop()

    required = {
        "컨설턴트 이름": consultant_name,
        "학생명": student_name,
        "학교명": school_name,
        "학생 전화번호": student_phone,
        "메일주소": email,
        "학부모 연락처": parent_phone,
    }
    missing = [k for k, v in required.items() if not str(v).strip()]
    if missing:
        st.error("필수값 누락: " + ", ".join(missing))
        st.stop()

    pdf_text = extract_pdf_text(pdf_file) if pdf_file else ""
    holistic_level, holistic_detail, holistic_summary = analyze_holistic_5level(pdf_text)
    st.success(holistic_summary)

    choices: List[SupportChoice] = []
    st.subheader("4) 50%, 70% 기준 진단 결과(4단계)")
    for i, (uni, dept, atype, track) in enumerate(support_inputs, start=1):
        if not (uni and dept and atype and track):
            continue
        c50, c70 = get_cutoff(cutoff_df, uni, dept, atype, track)
        level, reason = rating_4level(student_index, c50, c70)
        choices.append(SupportChoice(uni, dept, atype, track, level, reason))
        st.write(f"{i}. {uni} / {dept} / {atype} / {track} -> {level}")
        st.caption(reason)

    if not choices:
        st.warning("유효한 지원 입력이 없어 진단 결과를 생성하지 못했습니다.")
        st.stop()

    payload = {
        "consultant_name": consultant_name,
        "student_name": student_name,
        "school_name": school_name,
        "grade": grade,
        "student_phone": student_phone,
        "email": email,
        "parent_phone": parent_phone,
        "student_index": student_index,
        "pdf_summary": holistic_summary,
        "holistic_score": sum(holistic_detail.values()) / len(holistic_detail) if holistic_detail else 0,
        "holistic_level": holistic_level,
        "top15_reference": TOP15,
    }
    save_session(payload, choices)

    st.subheader("5) 최종 분석 보고서")
    md = report_markdown(payload, choices, holistic_detail)
    pdf_bytes = markdown_to_pdf_bytes(md)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    st.download_button("보고서 Markdown 다운로드", md.encode("utf-8"), file_name=f"admission_report_{stamp}.md")
    st.download_button("보고서 PDF 다운로드", pdf_bytes, file_name=f"admission_report_{stamp}.pdf")

    st.subheader("6) DB 누적 저장")
    st.info(f"진단 결과가 DB에 저장되었습니다: {DB_PATH}")
    st.caption("구글드라이브 원천데이터는 CSV로 내보내 본 서비스 업로드 기능으로 교체 반영하세요.")


if __name__ == "__main__":
    main()
