import os 

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
DATASET_DIR = f"{ROOT_DIR}/Datasets"

'''
A really small number
'''
EPSILON = 0.0001

KENDALL_TAU_GROUP_SIZE = 10