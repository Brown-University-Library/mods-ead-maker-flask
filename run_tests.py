"""
Runs tests for this webap.

Usage examples:
    (all) uv run ./run_tests.py
    (file) uv run ./run_tests.py tests.test
    (class) uv run ./run_tests.py tests.test.TestMain
    (method) uv run ./run_tests.py tests.test.TestMain.test_sum_two_numbers_returns_total

    Also takes a -v or --verbose flag to increase verbosity to level 2, which yields:
    name of test being run
    the test's docstring ("Checks...) ... ...and the result (ok)
    etc...

"""

import argparse
import sys
import unittest


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        'test_path',
        nargs='?',
        default='tests',
        help='Optional test module/class/method, or a directory to discover tests from.',
    )
    parser.add_argument(
        '-v',
        '--verbose',
        action='store_true',
        help='Increase verbosity to level 2.',
    )
    args = parser.parse_args()

    if args.test_path.endswith('.py'):
        args.test_path = args.test_path[:-3]

    if args.test_path in {'tests', 'test', '.'}:
        suite = unittest.defaultTestLoader.discover('tests')
    else:
        suite = unittest.defaultTestLoader.loadTestsFromName(args.test_path)

    verbosity = 2 if args.verbose else 1
    runner = unittest.TextTestRunner(verbosity=verbosity)
    result = runner.run(suite)
    sys.exit(0 if result.wasSuccessful() else 1)


if __name__ == '__main__':
    main()
