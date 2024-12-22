import unittest
from datetime import datetime
from location_data_v1 import convert_timestamp, validate_time_format

class TestLocationData(unittest.TestCase):

    def test_convert_timestamp_valid(self):
        ts = 1000000000  # Example timestamp
        date_str, time_str, time_zone = convert_timestamp(ts)
        self.assertEqual(date_str, '09/09/2032')
        self.assertEqual(time_str, '01:46:40')
        self.assertEqual(time_zone, 'AEST (UTC+10)')

    def test_convert_timestamp_invalid(self):
        ts = 'invalid_timestamp'
        date_str, time_str, time_zone = convert_timestamp(ts)
        self.assertEqual(date_str, 'invalid_timestamp')
        self.assertEqual(time_str, 'invalid_timestamp')
        self.assertEqual(time_zone, 'Unknown')

    def test_validate_time_format_valid(self):
        self.assertTrue(validate_time_format('12:34'))
        self.assertTrue(validate_time_format('00:00'))
        self.assertTrue(validate_time_format('23:59'))

    def test_validate_time_format_invalid(self):
        self.assertFalse(validate_time_format('24:00'))
        self.assertFalse(validate_time_format('12:60'))
        self.assertFalse(validate_time_format('invalid_time'))

if __name__ == '__main__':
    unittest.main()
