"""전달 패키지만으로 독립 재계산한다. 실행: python 03_집계검증.py [패키지 경로]"""
from pathlib import Path
import sys
import numpy as np
import pandas as pd

HERE = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).resolve().parent


def read(name):
    return pd.read_csv(HERE / name, encoding="utf-8-sig", low_memory=False)


def compare(name, calculated, expected, keys, columns, tolerance=0.00011):
    assert not calculated.duplicated(keys).any(), f"{name}: 재계산 키 중복"
    assert not expected.duplicated(keys).any(), f"{name}: 산출물 키 중복"
    joined = calculated.merge(expected, on=keys, how="outer", suffixes=("_calc", "_ref"), indicator=True)
    assert joined._merge.eq("both").all(), f"{name}: 키 누락 또는 추가"
    for col in columns:
        a, b = joined[f"{col}_calc"], joined[f"{col}_ref"]
        assert a.isna().equals(b.isna()), f"{name}: {col} 결측 불일치"
        if pd.api.types.is_numeric_dtype(a) and pd.api.types.is_numeric_dtype(b):
            assert np.allclose(a, b, atol=tolerance, rtol=0, equal_nan=True), f"{name}: {col} 수치 불일치"
        else:
            assert a.fillna("").equals(b.fillna("")), f"{name}: {col} 문자열 불일치"
    print(f"통과: {name} {len(expected):,}행 / {len(columns)}개 비교 열")


def main():
    dim = read("dim_indicator.csv").set_index("indicator_id")
    micro = pd.concat([pd.read_csv(p, low_memory=False) for p in sorted(HERE.glob("micro_*.csv"))], ignore_index=True)
    assert micro.indicator_id.isin(dim.index).all()
    assert not micro.duplicated(["indicator_id", "조사연도", "응답자ID"]).any()
    stats = ["응답수", "결측수", "평균", "표준편차", "중앙값", "백분위25", "백분위75",
             "환산100", "구인환산100", "상위2선지비율", "긍정응답률", "절사평균",
             "지출자수", "지출참여율", "지출자평균", "지출자중앙값", "대표값",
             "발생건수합계", "분모학생수합계", "발생률산출학교수"]
    records = []
    for region, sample in [(0, micro), (1, micro[micro.지역규모코드 == 1]), (2, micro[micro.지역규모코드 == 2])]:
        for (cid, year), group in sample.groupby(["indicator_id", "조사연도"]):
            definition = dim.loc[cid]
            values = group.값.dropna()
            n, st = len(values), definition.통계유형
            r = dict.fromkeys(stats, np.nan)
            r.update(indicator_id=cid, 조사연도=year, 지역규모코드=region, 응답수=n, 결측수=len(group)-n)
            if n:
                avg = values.mean()
                r.update(평균=round(avg, 4), 표준편차=round(values.std(ddof=1), 4) if n > 1 else np.nan,
                         중앙값=values.median(), 백분위25=values.quantile(.25), 백분위75=values.quantile(.75))
                if st == "likert":
                    lo, hi = definition.척도최소, definition.척도최대
                    r["환산100"] = round((avg-lo)/(hi-lo)*100, 2)
                    scored = lo + hi - values if definition.역문항 == "Y" else values
                    r["구인환산100"] = round((scored.mean()-lo)/(hi-lo)*100, 2)
                    r["상위2선지비율"] = round(values.isin([hi, hi-1]).mean()*100, 2)
                    r["대표값"] = r["환산100"]
                elif st == "binary":
                    r["긍정응답률"] = round(values.eq(definition.긍정코드).mean()*100, 2)
                    r["대표값"] = r["긍정응답률"]
                elif st == "continuous":
                    payers = values[values > 0]
                    r.update(지출자수=len(payers), 지출참여율=round(len(payers)/len(group)*100, 2),
                             지출자평균=round(payers.mean(), 4), 지출자중앙값=payers.median())
                    r["대표값"] = payers.median() if len(payers) else values.median()
                    if n >= 20:
                        low, high = values.quantile([.01, .99])
                        r["절사평균"] = round(values[values.between(low, high)].mean(), 4)
                elif st == "score100":
                    r["환산100"] = r["대표값"] = round(avg, 2)
                elif st == "school_count":
                    r["대표값"] = round(avg, 4)
                elif st == "school_rate":
                    pair = group[group.값.notna() & group.분모학생수.gt(0)]
                    r["발생률산출학교수"] = len(pair)
                    if len(pair):
                        r["발생건수합계"], r["분모학생수합계"] = pair.값.sum(), pair.분모학생수.sum()
                        r["대표값"] = round(pair.값.sum()/pair.분모학생수.sum()*100, 4)
            records.append(r)
    fw = pd.DataFrame(records)
    keys = ["indicator_id", "조사연도", "지역규모코드"]
    compare("지표 전체·지역 집계", fw, read("fact_indicator_wave.csv"), keys, stats)
    compare("지역 분리본", fw[fw.지역규모코드 != 0], read("fact_indicator_wave_region.csv"), keys, stats)

    eligible = dim.index[dim.소주제평균포함 == "Y"]
    core = fw[(fw.지역규모코드 == 0) & fw.indicator_id.isin(eligible)].merge(dim.reset_index(), on="indicator_id")
    topic_keys = ["소주제코드", "응답주체", "조사연도"]
    ft = core.groupby(topic_keys).agg(지표수=("indicator_id", "nunique"), 평균환산100=("구인환산100", "mean"),
                                    평균응답수=("응답수", "mean")).reset_index()
    ft["평균환산100"] = ft["평균환산100"].round(2)
    ft["평균응답수"] = ft["평균응답수"].round(0)
    compare("소주제 평균", ft, read("fact_topic_wave.csv"), topic_keys, ["지표수", "평균환산100", "평균응답수"])

    previous, latest = sorted(read("dim_wave.csv").조사연도.unique())[-2:]
    # 실제로 양 주기에 시계열 비교 가능한 지표들만 후보로 선정
    comparable_topics = core.loc[core.시계열비교가능.isin(["Y", "조건부"]), ["소주제코드", "응답주체"]].drop_duplicates()
    fh = ft.merge(comparable_topics).pivot(index=["소주제코드", "응답주체"], columns="조사연도", values="평균환산100")
    fh = fh.dropna(subset=[previous, latest]).reset_index().rename(columns={previous:f"값_{previous}", latest:f"값_{latest}"})
    fh["증감"] = (fh[f"값_{latest}"] - fh[f"값_{previous}"]).round(2)
    fh["증감률"] = (fh.증감/fh[f"값_{previous}"].replace(0, np.nan)*100).round(2)
    fh["절대증감순위"] = fh.증감.abs().rank(method="min", ascending=False).astype(int)
    fh["변화방향"] = np.where(fh.증감 > 0, "증가", np.where(fh.증감 < 0, "감소", "유지"))
    compare("헤드라인 후보", fh, read("fact_headline.csv"), ["소주제코드", "응답주체"],
            [f"값_{previous}", f"값_{latest}", "증감", "증감률", "절대증감순위", "변화방향"])
    assert read("fact_headline.csv").선정상태.eq("후보(선정 미확정)").all()

    distribution = micro[micro.indicator_id.isin(dim.index[dim.통계유형.isin(["likert", "binary"])]) & micro.값.notna()]
    fd = distribution.groupby(["indicator_id", "조사연도", "값"]).size().reset_index(name="응답수").rename(columns={"값":"코드"})
    fd["비율"] = (fd.응답수/fd.groupby(["indicator_id", "조사연도"]).응답수.transform("sum")*100).round(2)
    compare("문항별 원척도 분포", fd, read("fact_indicator_dist.csv"), ["indicator_id", "조사연도", "코드"], ["응답수", "비율"])
    print("모든 검증 통과")


if __name__ == "__main__":
    main()
