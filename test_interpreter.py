import csv
import importlib
import os
from pathlib import Path
import tempfile
import unittest


class InterpreterTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp = tempfile.TemporaryDirectory()
        os.environ['QFTINTERPRETER_DATA_DIR'] = cls.temp.name
        cls.app = importlib.import_module('qft_interpreter')

    @classmethod
    def tearDownClass(cls):
        for handler in cls.app.qft_logger.handlers[:]:
            handler.close()
            cls.app.qft_logger.removeHandler(handler)
        os.environ.pop('QFTINTERPRETER_DATA_DIR', None)
        cls.temp.cleanup()

    def test_existing_interpretation_examples(self):
        for label, nil, tb1, tb2, mit, expected in self.app.SELF_TEST_CASES:
            with self.subTest(label=label):
                self.assertEqual(self.app.interpret_qft(nil, tb1, tb2, mit)['result'], expected)

    def test_zero_values_are_not_missing_in_log(self):
        self.app.log_event('INTERPRET', sample_id='zero-test', details={'nil': 0, 'tb1': 0.0, 'tb2': 0, 'mit': 0})
        with open(self.app.LOG_FILENAME, newline='', encoding='utf-8') as stream:
            row = list(csv.reader(stream))[-1]
        self.assertEqual(row[4:8], ['0.000'] * 4)
        self.assertEqual(Path(self.app.LOG_FILENAME).parent, Path(self.temp.name))


if __name__ == '__main__':
    unittest.main()
