import unittest

from main import sum_two_numbers


class TestMain(unittest.TestCase):
    def test_sum_two_numbers_returns_total(self) -> None:
        """
        Checks that two numbers are summed.
        """
        total = sum_two_numbers(2, 3)

        self.assertEqual(5, total)


if __name__ == '__main__':
    unittest.main()
