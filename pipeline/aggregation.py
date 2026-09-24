"""유형별 통계량. 역채점은 구인 평균용 점수에만 적용한다."""
import numpy as np

OPTIONAL_STATS = ["평균", "표준편차", "중앙값", "백분위25", "백분위75", "절사평균",
                  "환산100", "구인환산100", "상위2선지비율", "긍정응답률", "지출자수",
                  "지출참여율", "지출자평균", "지출자중앙값", "발생건수합계",
                  "분모학생수합계", "발생률산출학교수", "대표값"]


def aggregate(group, indicator):
    v = group["값"].dropna()
    st = indicator["통계유형"]
    out = dict.fromkeys(OPTIONAL_STATS, np.nan)
    units = {"likert": "환산100점", "binary": "긍정응답률(%)", "continuous": "지출자 중앙값(만원)",
             "score100": "역량 환산점수 평균(점)", "school_rate": "학생 100명당 발생 건수",
             "school_count": "학교당 평균 실시 횟수"}
    out.update(응답수=len(v), 결측수=int(group["값"].isna().sum()), 대표값유형=units.get(st, ""))
    if not len(v):
        return out
    mean = float(v.mean())
    out.update(평균=round(mean, 4), 표준편차=round(float(v.std(ddof=1)), 4) if len(v) > 1 else np.nan,
               중앙값=float(v.median()))
    out.update(백분위25=float(v.quantile(.25)), 백분위75=float(v.quantile(.75)))
    if st == "likert":
        lo, hi = indicator["척도최소"], indicator["척도최대"]
        scaled = (mean - lo) / (hi - lo) * 100
        out["환산100"] = round(scaled, 2)
        out["구인환산100"] = round(100 - scaled if indicator["역문항"] == "Y" else scaled, 2)
        out["상위2선지비율"] = round(float(v.isin([hi, hi - 1]).mean()) * 100, 2)
        out["대표값"] = out["환산100"]
    elif st == "binary":
        out["긍정응답률"] = round(float(v.eq(indicator["긍정코드"]).mean()) * 100, 2)
        out["대표값"] = out["긍정응답률"]
    elif st == "continuous":
        pay = v[v > 0]
        out.update(지출자수=len(pay), 지출참여율=round(len(pay) / len(group) * 100, 2),
                   지출자평균=round(float(pay.mean()), 4) if len(pay) else np.nan,
                   지출자중앙값=float(pay.median()) if len(pay) else np.nan)
        out["대표값"] = out["지출자중앙값"] if len(pay) else out["중앙값"]
        if len(v) >= 20:
            lo, hi = v.quantile(.01), v.quantile(.99)
            out["절사평균"] = round(float(v[v.between(lo, hi)].mean()), 4)
    elif st == "score100":
        out["환산100"] = round(mean, 2)
        out["대표값"] = out["환산100"]
    elif st == "school_count":
        out["대표값"] = out["평균"]
    elif st == "school_rate":
        paired = group[group["값"].notna() & group["분모학생수"].gt(0)]
        out["발생률산출학교수"] = len(paired)
        if len(paired):
            out["발생건수합계"] = float(paired["값"].sum())
            out["분모학생수합계"] = float(paired["분모학생수"].sum())
            out["대표값"] = round(out["발생건수합계"] / out["분모학생수합계"] * 100, 4)
    return out
