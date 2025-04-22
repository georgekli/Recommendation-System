import unittest
import numpy as np
import pandas as pd
from pathlib import Path
import utils 

class TestCalc(unittest.TestCase):
    
    def test_create_random_group(self):
        self.assertEqual(utils.create_random_group(5, 10), [2, 6, 9, 4, 3])
        with self.assertRaises(ValueError):
            utils.create_random_group(5, 2)

if __name__ == '__main__':
    unittest.main()