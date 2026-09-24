"""학교 단위 중복 응답 처리. 원자료의 행 순서로 대표 응답을 선택하지 않는다."""
import pandas as pd


def collapse_schools(frame, variables):
    if frame.SCHID.isna().any():
        raise ValueError("학교 식별자 결측은 학교 단위 집계에 사용할 수 없습니다")
    grouped = frame.groupby("SCHID", sort=True)
    rows, notes = [], {}
    for school_id, group in grouped:
        row = {"SCHID": school_id}
        for var in variables:
            values = group[var].where((group[var] >= 0) & (group[var] % 1 == 0)).dropna().unique()
            row[var] = values[0] if len(values) == 1 else float("nan")
            if len(group) > 1:
                notes[(school_id, var)] = "학교중복충돌→결측" if len(values) > 1 else "학교중복응답통합"
        rows.append(row)
    return pd.DataFrame(rows), notes
