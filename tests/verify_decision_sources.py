"""SPSS 원자료와 v2 산출물 직접 대조. 빌드·집계 함수에 의존하지 않는다."""
from pathlib import Path
import sys
import numpy as np
import pandas as pd
import pyreadstat

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "pipeline"))
from common import DELIVERY, RAW_W1, RAW_W2, find_file

out = Path(DELIVERY)
dim = pd.read_csv(out / "dim_indicator.csv").set_index("indicator_id")
fact = pd.read_csv(out / "fact_indicator_wave.csv")
overall = fact[fact.지역규모코드 == 0].set_index(["indicator_id", "조사연도"])
topics = pd.read_csv(out / "fact_topic_wave.csv")
headline = pd.read_csv(out / "fact_headline.csv")
assert len(dim) == 246
assert not dim.index.str.endswith("_TOT").any()
assert dim.loc[dim.index.str.startswith("ST_CULTURE_"), "화면표시"].eq("숨김").all()
assert dim.loc[["PR_TALK_06", "PR_TALK_07"], "화면표시"].eq("숨김").all()
assert "S0704" not in set(topics.소주제코드) | set(headline.소주제코드)
assert "S0704" not in set(pd.read_csv(out / "dim_topic.csv").소주제코드)
assert not dim.loc[dim.응답주체 == "학생", "var_2025"].fillna("").str.startswith("STQ29_").any()
assert set(fact.학교급) == {"중학교"}
assert headline.선정상태.eq("후보(선정 미확정)").all()
assert pd.read_csv(out / "dim_category.csv").확정상태.eq("잠정(분류 확정 대기)").all()

for wave, year, folder, student, parent, school, dbname in [
    (1, 2021, RAW_W1, "student total", "PAR", "YM1_SCH(", "SCH DB"),
    (2, 2025, RAW_W2, "1. 학생", "2. 학부모.sav", "4. 학교", "5. SCH DB"),
]:
    student_data, _ = pyreadstat.read_sav(find_file(folder, student))
    for key in ("CC", "SC", "AC"):
        source = student_data[f"{key}_conversion"].where(student_data[f"{key}_conversion"].between(0, 100))
        row = overall.loc[(f"ST_COMP_{key}_CONV", year)]
        assert row.대표값 == round(source.mean(), 2)
        assert row.응답수 == source.count()
        assert pd.isna(row.지출자수)
        if wave == 2:
            assert row.결측수 == 484
    for resp, filename in [("학생", student), ("학부모", parent)]:
        raw, metadata = pyreadstat.read_sav(find_file(folder, filename))
        subset = dim[(dim.응답주체 == resp) & (dim.소주제코드 == "S0401")]
        scores = []
        for cid, item in subset.iterrows():
            var = item[f"var_{year}"]
            labels = metadata.variable_value_labels[var]
            values = raw[var].where(raw[var].isin(labels))
            lo, hi = min(labels), max(labels)
            assert abs(overall.loc[(cid, year), "평균"] - values.mean()) < .0001
            if item.소주제평균포함 == "Y":
                if cid in ("ST_RELATION_06", "PR_RELATION_02", "PR_RELATION_05"):
                    values = lo + hi - values
                scores.append(round((values.mean()-lo)/(hi-lo)*100, 2))
        actual = topics[(topics.소주제코드 == "S0401") & (topics.응답주체 == resp) & (topics.조사연도 == year)].iloc[0]
        assert actual.지표수 == len(scores)
        assert abs(actual.평균환산100 - round(float(np.mean(scores)), 2)) < .010001

    raw_school, _ = pyreadstat.read_sav(find_file(folder, school))
    raw_db, _ = pyreadstat.read_sav(find_file(folder, dbname))
    cols = [f"YM{wave}_DB2_{g}{sex}" for g in [1, 2, 3] for sex in ["M", "W"]]
    enrollment = raw_db.set_index("SCHID")[cols].sum(axis=1, min_count=6)
    school_micro = pd.read_csv(out / f"micro_school_{year}.csv")
    conflicts = 0
    for cid, item in dim[dim.응답주체 == "학교"].iterrows():
        var = item[f"var_{year}"]
        grouped = raw_school.groupby("SCHID")[var]
        expected = {}
        for sid, values in grouped:
            valid = values[values.ge(0) & values.mod(1).eq(0)].dropna()
            expected[sid] = valid.min() if valid.nunique() == 1 else np.nan
            conflicts += valid.nunique() > 1
        expected = pd.Series(expected).sort_index()
        actual_values = school_micro[school_micro.indicator_id == cid].set_index("학교ID").값.sort_index()
        assert actual_values.index.equals(expected.index)
        assert np.allclose(actual_values, expected, equal_nan=True)
        row = overall.loc[(cid, year)]
        assert row.응답수 == expected.count()
        if item.통계유형 == "school_rate":
            pairs = pd.DataFrame({"v": expected, "n": enrollment}).dropna()
            pairs = pairs[pairs.n > 0]
            expected_rate = round(pairs.v.sum()/pairs.n.sum()*100, 4)
            assert abs(row.대표값-expected_rate) < .000001
            assert row.분모학생수합계 == pairs.n.sum()
            assert row.발생률산출학교수 == len(pairs)
        else:
            assert row.대표값 == round(expected.mean(), 4)
    print(f"원자료 대조 통과: {year}, 학교 {raw_school.SCHID.nunique()}개, 충돌 학교·문항 {conflicts}건")

teacher, _ = pyreadstat.read_sav(find_file(RAW_W2, "3. 교사.sav"))
digital = dim[dim.index.str.startswith("TR_DIGITAL_")]
assert len(digital) == 4 and digital.화면표시.eq("단년도").all()
assert not ((fact.indicator_id.str.startswith("TR_DIGITAL_")) & (fact.조사연도 == 2021)).any()
assert "S0303" not in set(headline.소주제코드)
for cid, item in digital.iterrows():
    values = teacher[item.var_2025]
    assert overall.loc[(cid, 2025), "대표값"] == round((values.mean()-1)/4*100, 2)
assert dim.loc["ST_RELATION_02", "소주제평균포함"] == "N"
assert topics.loc[(topics.소주제코드 == "S0401") & (topics.응답주체 == "학생"), "주의사항"].str.contains("ST_RELATION_02").all()
print("확정 답변 10건의 원자료·산출물 대조 통과")
