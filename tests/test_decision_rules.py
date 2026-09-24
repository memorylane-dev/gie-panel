"""2026-09-24 결정에 대한 경계 조건: 0점·0건, 중복 충돌, 분자/분모 정합성."""
import sys
from pathlib import Path
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "pipeline"))
import pandas as pd
from aggregation import aggregate
from cleaning import collapse_schools


class DecisionRulesTest(unittest.TestCase):
    def test_score_keeps_zero_and_is_not_money(self):
        result = aggregate(pd.DataFrame({"값": [0, 50, 100, None]}), {"통계유형": "score100"})
        self.assertEqual(result["대표값"], 50)
        self.assertEqual(result["응답수"], 3)
        self.assertEqual(result["결측수"], 1)
        self.assertTrue(pd.isna(result["지출자수"]))

    def test_reverse_only_changes_construct_score(self):
        group = pd.DataFrame({"값": [1, 1, 3, None]})
        result = aggregate(group, {"통계유형": "likert", "척도최소": 1, "척도최대": 5, "역문항": "Y"})
        self.assertEqual(result["환산100"], 16.67)
        self.assertEqual(result["구인환산100"], 83.33)
        self.assertEqual(result["대표값"], 16.67)
        self.assertEqual(group.값.iloc[0], 1)

    def test_school_rate_pairs_denominator_and_keeps_zero(self):
        group = pd.DataFrame({"값": [10, 0, None, 999], "분모학생수": [100, 300, 1000, None]})
        result = aggregate(group, {"통계유형": "school_rate"})
        self.assertEqual(result["대표값"], 2.5)
        self.assertEqual(result["분모학생수합계"], 400)
        self.assertEqual(result["발생건수합계"], 10)
        self.assertEqual(result["발생률산출학교수"], 2)
        self.assertEqual(result["응답수"], 3)

    def test_school_rate_does_not_fill_missing_denominator(self):
        result = aggregate(pd.DataFrame({"값": [1, 0], "분모학생수": [0, None]}), {"통계유형": "school_rate"})
        self.assertTrue(pd.isna(result["대표값"]))
        self.assertEqual(result["발생률산출학교수"], 0)

    def test_school_duplicate_conflict_does_not_depend_on_row_order(self):
        raw = pd.DataFrame({"SCHID": [1, 1, 2, 2, 3], "q": [0, 0, 2, 5, -1], "r": [None, 4, 3, 3, 2]})
        first, notes = collapse_schools(raw, ["q", "r"])
        second, _ = collapse_schools(raw.iloc[::-1], ["q", "r"])
        pd.testing.assert_frame_equal(first, second)
        self.assertEqual(first.q.iloc[0], 0)
        self.assertTrue(pd.isna(first.q.iloc[1]))
        self.assertTrue(pd.isna(first.q.iloc[2]))
        self.assertEqual(first.r.iloc[0], 4)
        self.assertEqual(notes[(2, "q")], "학교중복충돌→결측")

    def test_school_count_mean_includes_zero(self):
        result = aggregate(pd.DataFrame({"값": [0, 10, None]}), {"통계유형": "school_count"})
        self.assertEqual(result["대표값"], 5)
        self.assertTrue(pd.isna(result["지출자수"]))


if __name__ == "__main__":
    unittest.main()
