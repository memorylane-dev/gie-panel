# -*- coding: utf-8 -*-
"""집계 검증 스크립트

02_집계사양서.md 3~7장의 집계 절차를 구현하고, 그 결과를 함께 제공된
fact_*.csv 와 대조합니다. 업체 시스템에 집계 로직을 구현하신 뒤
동일한 결과가 나오는지 확인하는 용도입니다.

이 스크립트 자체가 사양서의 참조 구현이므로, 로직이 모호할 때 참고하십시오.

실행:  python 03_집계검증.py
필요:  Python 3.9+, pandas   (pip install pandas)
"""
import glob, sys
import pandas as pd

HERE = __file__.rsplit("/", 1)[0] if "/" in __file__ else "."
R = lambda n: pd.read_csv(f"{HERE}/{n}", encoding="utf-8-sig", low_memory=False)

# ── 입력 ────────────────────────────────────────────────────
micro = pd.concat([pd.read_csv(f, encoding="utf-8-sig", low_memory=False)
                   for f in sorted(glob.glob(f"{HERE}/micro_*.csv"))], ignore_index=True)
dim   = R("dim_indicator.csv")
print(f"마이크로데이터 {len(micro):,}행 / 지표 {len(dim)}개 로드")

P = dim.set_index("indicator_id")[
    ["통계유형", "척도최소", "척도최대", "긍정코드", "역문항",
     "시계열비교가능", "대분류코드", "대분류", "소주제코드", "소주제", "응답주체"]]

# ── 3~5장: 지표 단위 집계 ───────────────────────────────────
def aggregate(df, region_code, region_name):
    out = []
    for (cid, year), g in df.groupby(["indicator_id", "조사연도"]):
        if cid not in P.index:
            continue
        p = P.loc[cid]
        v = g["값"].dropna()
        rec = dict(indicator_id=cid, 조사연도=year,
                   지역규모코드=region_code, 지역규모=region_name,
                   응답수=len(v), 결측수=len(g) - len(v))
        if len(v) == 0:
            out.append(rec); continue
        rec["평균"] = round(float(v.mean()), 4)

        st = p["통계유형"]
        if st == "likert":                                    # 4-1
            lo, hi = p["척도최소"], p["척도최대"]
            rec["환산100"] = round((v.mean() - lo) / (hi - lo) * 100, 2)
            rec["상위2선지비율"] = round(float(v.isin([hi, hi - 1]).mean() * 100), 2)
            rec["대표값"] = rec["환산100"]
        elif st == "binary":                                  # 4-2
            pos = p["긍정코드"]
            rec["긍정응답률"] = round(float((v == pos).mean() * 100), 2) if pd.notna(pos) else None
            rec["대표값"] = rec["긍정응답률"]
        elif st == "continuous":                              # 4-3
            pay = v[v > 0]
            rec["지출자수"] = len(pay)
            rec["지출참여율"] = round(len(pay) / len(g) * 100, 2)
            rec["지출자중앙값"] = float(pay.median()) if len(pay) else None
            rec["대표값"] = rec["지출자중앙값"]
        out.append(rec)
    return pd.DataFrame(out)

fw = pd.concat([
    aggregate(micro, 0, "전체"),
    *[aggregate(g, int(rc), nm) for rc, nm in [(1, "중소도시"), (2, "읍면지역")]
      for g in [micro[micro.지역규모코드 == rc]]],
], ignore_index=True)

# ── 7-1장: 소주제 집계 ──────────────────────────────────────
core = fw[fw.지역규모 == "전체"].merge(P.reset_index(), on="indicator_id")
core = core[(core.통계유형 == "likert")
            & (core.시계열비교가능.isin(["Y", "조건부"]))
            & (core.역문항 == "N")]                            # 역문항 제외 (사양서 7-1장 조건 3)
ft = (core.groupby(["대분류코드", "대분류", "소주제코드", "소주제", "응답주체", "조사연도"])
          .agg(지표수=("indicator_id", "nunique"),
               평균환산100=("환산100", "mean")).reset_index())
ft["평균환산100"] = ft["평균환산100"].round(2)

# ── 7-2장: 헤드라인 ─────────────────────────────────────────
years = sorted(ft.조사연도.unique())
prev, curr = years[-2], years[-1]
p2 = ft.pivot_table(index=["대분류코드", "대분류", "소주제코드", "소주제", "응답주체"],
                    columns="조사연도", values="평균환산100").reset_index().dropna(subset=[prev, curr])
p2["증감"] = (p2[curr] - p2[prev]).round(2)
p2["절대증감순위"] = p2["증감"].abs().rank(ascending=False, method="min").astype(int)

# ── 8장: 검증 ───────────────────────────────────────────────
print("\n" + "=" * 62)
ok = True
def check(name, got, exp):
    global ok
    good = got == exp
    ok &= good
    print(f"  {'✔' if good else '✘'} {name:34s} {got:>7} (기준 {exp})")

ref = R("fact_indicator_wave.csv")
check("fact_indicator_wave 행수", len(fw), len(ref))
check("fact_topic_wave 소주제 수", ft.소주제코드.nunique(),
      R("fact_topic_wave.csv").소주제코드.nunique())
check("fact_headline 행수", len(p2), len(R("fact_headline.csv")))

m = fw.merge(ref[["indicator_id", "조사연도", "지역규모코드", "응답수", "대표값"]],
             on=["indicator_id", "조사연도", "지역규모코드"], suffixes=("_계산", "_기준"))
check("응답수 불일치 건수", int((m.응답수_계산 != m.응답수_기준).sum()), 0)
diff = (m.대표값_계산 - m.대표값_기준).abs()
worst = round(float(diff.max()), 4) if diff.notna().any() else 0.0
print(f"  {'✔' if worst <= 0.01 else '✘'} {'대표값 최대 오차':34s} {worst:>7} (기준 0.01 이하)")
ok &= worst <= 0.01

print("=" * 62)
print("결과:", "모든 검증 통과 — 구현이 정확합니다." if ok
      else "불일치 발견 — 02_집계사양서.md를 다시 확인하십시오.")
sys.exit(0 if ok else 1)
