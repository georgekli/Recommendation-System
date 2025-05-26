import os 

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
DATASET_DIR = f"{ROOT_DIR}/Datasets"

'''Smallest possible rating someone can have for an item'''
EPSILON = 0.0001

KENDALL_TAU_GROUP_SIZE = 10