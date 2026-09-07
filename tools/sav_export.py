# -*- coding: utf-8 -*-
"""SPSS(.sav) → CSV/Excel 변환 + 코드북 생성.

SPSS 프로그램 없이 원자료를 열람하기 위한 도구.
- 라벨 적용본(사람이 읽는 값)과 코드본(분석용 숫자)을 모두 생성
- 파일별 코드북(변수명·라벨·척도·선지·결측·기초통계)을 한 엑셀로 통합
"""
import os, unicodedata
import pyreadstat, pandas as pd

N = lambda s: unicodedata.normalize("NFC", str(s)) if s is not None else ""
ROOT = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
OUT  = os.path.join(ROOT, "data", "exports")

WAVES = {
    "wave1_2021": {
        "2021YM1_student total(공개용).sav": "학생",
        "2021YM1_PAR(공개용).sav":           "학부모",
        "2021YM1_TEA(공개용).sav":           "교사",
        "2021YM1_SCH(공개용).sav":           "학교",
        "2021YM1_SCH DB(공개용).sav":        "학교DB",
    },
    "wave2_2025": {
        "1. 학생.sav":   "학생",
        "2. 학부모.sav": "학부모",
        "3. 교사.sav":   "교사",
        "4. 학교.sav":   "학교",
        "5. SCH DB.sav": "학교DB",
    },
}

def resolve(folder, name):
    """macOS NFD 파일명 대응."""
    for f in sorted(os.listdir(folder)):
        if N(f) == name:
            return os.path.join(folder, f)
    raise FileNotFoundError(f"{name} in {folder}")

codebooks = {}
for wave, files in WAVES.items():
    src = os.path.join(ROOT, "data", "raw", wave, "spss")
    dst_lab  = os.path.join(OUT, wave, "라벨적용")
    dst_code = os.path.join(OUT, wave, "코드값")
    os.makedirs(dst_lab, exist_ok=True); os.makedirs(dst_code, exist_ok=True)

    for fname, resp in files.items():
        df, meta = pyreadstat.read_sav(resolve(src, fname))
        # 정수만 담긴 열은 Int64로 — ID·코드가 "1001.0"으로 보이지 않게
        for c in df.columns:
            if pd.api.types.is_float_dtype(df[c]):
                nn = df[c].dropna()
                if len(nn) and (nn % 1 == 0).all():
                    df[c] = df[c].astype("Int64")
        labels = {c: {k: N(v) for k, v in d.items()}
                  for c, d in meta.variable_value_labels.items()}
        varlab = {c: N(meta.column_names_to_labels.get(c)) for c in meta.column_names}

        # 코드값본
        df.to_csv(os.path.join(dst_code, f"{resp}.csv"),
                  index=False, encoding="utf-8-sig")
        # 라벨적용본
        lab = df.copy()
        for c, m in labels.items():
            if c in lab.columns:
                lab[c] = lab[c].map(m).fillna(lab[c])
        lab.to_csv(os.path.join(dst_lab, f"{resp}.csv"),
                   index=False, encoding="utf-8-sig")

        # 코드북
        rows = []
        for c in meta.column_names:
            s = df[c]
            vl = labels.get(c, {})
            num = pd.to_numeric(s, errors="coerce")
            rows.append({
                "변수명": c,
                "변수설명": varlab[c],
                "척도": "명목/순서(선지 있음)" if vl else ("연속형" if num.notna().any() else "문자"),
                "선지수": len(vl) if vl else "",
                "선지": " / ".join(f"{k:g}){v}" for k, v in sorted(vl.items())) if vl else "",
                "응답수": int(s.notna().sum()),
                "결측수": int(s.isna().sum()),
                "최소": num.min() if num.notna().any() else "",
                "최대": num.max() if num.notna().any() else "",
                "평균": round(float(num.mean()), 3) if num.notna().any() else "",
            })
        codebooks[f"{wave[-4:]}_{resp}"] = pd.DataFrame(rows)
        print(f"  {wave}/{resp}: {len(df):,}행 × {len(df.columns)}변수")

xl = os.path.join(OUT, "코드북_전체.xlsx")
with pd.ExcelWriter(xl, engine="openpyxl") as xw:
    idx = pd.DataFrame(
        [{"시트": k, "변수 수": len(v)} for k, v in codebooks.items()])
    idx.to_excel(xw, sheet_name="0_목차", index=False)
    for name, cb in codebooks.items():
        cb.to_excel(xw, sheet_name=name[:31], index=False)
print(f"\n코드북: {xl}")
print(f"내보내기 위치: {OUT}")
