import csv
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent.parent
RAW_2026 = BASE_DIR / "2026학년도 수시전형 및 입결 검색기.csv"
RAW_2027 = BASE_DIR / "2027 수시검색기(경기대까지.csv"
OUT_2026 = BASE_DIR / "data" / "susi_explorer.csv"
OUT_2027 = BASE_DIR / "data" / "susi_explorer_2027.csv"


def _clean(x: str) -> str:
    return (x or "").replace("\n", " ").replace("\r", " ").strip()


def normalize_2026() -> int:
    if not RAW_2026.exists():
        return 0
    rows = []
    with RAW_2026.open("r", encoding="utf-8-sig", newline="") as f:
        r = csv.reader(f)
        header = None
        for row in r:
            if not header:
                if row and len(row) > 6 and _clean(row[2]) in ("대학교", "대학", "대학교명"):
                    header = row
                continue
            if not row or len(row) < 7:
                continue
            university = _clean(row[2])
            department = _clean(row[4])
            admission_type = _clean(row[5])
            track_name = _clean(row[6])
            if not (university and department and admission_type and track_name):
                continue
            rows.append(
                {
                    "year": "2026",
                    "university": university,
                    "department": department,
                    "admission_type": admission_type,
                    "track_name": track_name,
                }
            )
    return _write_unique(OUT_2026, rows)


def normalize_2027() -> int:
    if not RAW_2027.exists():
        return 0
    rows = []
    with RAW_2027.open("r", encoding="utf-8-sig", newline="") as f:
        r = csv.reader(f)
        for i, row in enumerate(r, start=1):
            if i <= 3:
                continue
            if not row or len(row) < 8:
                continue
            university = _clean(row[2])
            raw_type = _clean(row[4])
            track_name = _clean(row[5])
            department = _clean(row[6])
            if not (university and department and track_name):
                continue
            if raw_type == "종합":
                admission_type = "학생부종합"
            elif raw_type == "교과":
                admission_type = "학생부교과"
            elif raw_type == "논술":
                admission_type = "논술"
            else:
                admission_type = raw_type or "기타"
            rows.append(
                {
                    "year": "2027",
                    "university": university,
                    "department": department,
                    "admission_type": admission_type,
                    "track_name": track_name,
                }
            )
    return _write_unique(OUT_2027, rows)


def _write_unique(path: Path, rows: list[dict]) -> int:
    path.parent.mkdir(parents=True, exist_ok=True)
    seen = set()
    uniq = []
    for r in rows:
        key = (r["year"], r["university"], r["department"], r["admission_type"], r["track_name"])
        if key in seen:
            continue
        seen.add(key)
        uniq.append(r)
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["year", "university", "department", "admission_type", "track_name"])
        w.writeheader()
        w.writerows(uniq)
    return len(uniq)


if __name__ == "__main__":
    n2026 = normalize_2026()
    n2027 = normalize_2027()
    print(f"normalized_2026={n2026}")
    print(f"normalized_2027={n2027}")
