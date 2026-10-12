import unittest

import pandas as pd

from columns_manager import ColumnsManagerFrame


class CalculatedColumnFormulaTests(unittest.TestCase):
    def setUp(self):
        self.frame = object.__new__(ColumnsManagerFrame)
        self.df = pd.DataFrame(
            {
                "A21": [10, 2, 8],
                "B32": [5, 4, 8],
                "B3": [3, 7, 2],
                "Sell Qty": [0, 1, 0],
            }
        )

    def evaluate(self, formula):
        return self.frame._evaluate_formula_expr(self.df, formula).tolist()

    def test_if_formula_with_arithmetic(self):
        self.assertEqual(self.evaluate("IF(A21 > B32, A21 * b3, 0)"), [30, 0, 0])

    def test_if_without_false_value_returns_blank(self):
        result = self.evaluate("IF A21 > B32, A21 * B3")
        self.assertEqual(result[0], 30)
        self.assertTrue(pd.isna(result[1]))
        self.assertTrue(pd.isna(result[2]))

    def test_nested_and_condition(self):
        self.assertEqual(
            self.evaluate("IF(AND(A21 >= B32, `Sell Qty` = 0), 'Include', 'Exclude')"),
            ["Include", "Exclude", "Include"],
        )

    def test_nested_or_condition(self):
        self.assertEqual(
            self.evaluate("IF(OR(A21 < B32, B3 = 2), 1, 0)"),
            [0, 1, 1],
        )

    def test_excel_leading_equals_is_supported(self):
        self.assertEqual(self.evaluate("=IF(A21 > B32, A21, B32)"), [10, 4, 8])


if __name__ == "__main__":
    unittest.main()
