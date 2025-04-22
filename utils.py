import random
import numpy as np
import pandas as pd
import xlsxwriter

from pathlib import Path
    
random.seed(21)

################## Functions for storing and retrieving Matrices used in the code below ################################

""" Import an excel matrix with name equal to name
:param filename: The file of the matrix
:returns: Imported matrix
:rtype: np.array
"""
def import_excel_matrix(filename):
    file_path = f"{Path(__file__).parent}/Datasets/{filename}"
    WS = pd.read_excel(file_path)
    return np.array(WS)


""" Store a matrix in an excel sheet
:param matrix: The matrix to be stored
:param filename: Where to store the matrix
"""
def store_excel_matrix(matrix, filename):
    matrixWb = xlsxwriter.Workbook(f"{Path(__file__).parent}/Datasets/{filename}")
    matrixWs = matrixWb.add_worksheet()
    try:
        for j in range(0, matrix.shape[0]):
            matrixWs.write_row(j + 1, 0, matrix[j])
    except:
        for j in range(0, matrix.shape[0]):
            matrixWs.write_number(j + 1, 0, matrix[j])
    matrixWb.close()

# Helping Functions used for showing and generating data ###############################################################
""" Create a random group of users
:param groupSize: The desired size of the group
:param usersSize: The maximum user index
:returns: random list of user 
""" 
def create_random_group(groupSize, usersSize):
    return random.sample(range(0, usersSize), groupSize)


# Calculate and show the average score of an algorithm that suggests an item(1) by calculating the average rating of the
# item by users in each group (used for 100 groups)
def calculate_avg_algo_score(preferedItems, prefList, groups, groupSize):
    itemScore = [0] * 100
    for k in range(0, 100):
        for j in range(0, groupSize):
            itemScore[k] += prefList[groups[k][j]][preferedItems[k]]
        itemScore[k] = itemScore[k] / groupSize
    avgItemScore = np.sum(itemScore) / 100
    print(f"For group size = {groupSize} average score is {avgItemScore}")


# Functions used for the Task's algorithms #############################################################################
# Print top 10 items for the first 50 users
def print_top10(r):
    sortedItemsIndexes = np.zeros(r.shape[1])
    top10Items = np.zeros((r.shape[0], 10))
    top10Indexes = np.zeros((r.shape[0], 10))
    print("Top 10 items for every user(descending)")
    for users in range(0, 50):
        # Take indexes of sorted items
        sortedItemsIndexes = np.argsort(r[users])
        top10Indexes[users] = np.take(sortedItemsIndexes, range(r.shape[1] - 10, r.shape[1]))
        # Take top 10 indexes and make them items
        top10Items[users] = np.take(r[users], np.take(sortedItemsIndexes, range(r.shape[1] - 10, r.shape[1])))
        print("User: ", users, np.flip(top10Indexes[users]))

