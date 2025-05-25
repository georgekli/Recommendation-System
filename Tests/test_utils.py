""" Bring the packages here to test them """
import sys, os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

import unittest
from common import utils

class TestCalc(unittest.TestCase):
    
    def test_create_random_group(self):
        self.assertEqual(utils.create_random_group(5, 10), [2, 6, 8, 5, 9])
        with self.assertRaises(ValueError):
            utils.create_random_group(5, 2)

if __name__ == '__main__':
    unittest.main()