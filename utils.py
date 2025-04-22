import random
import numpy as np
import pandas as pd
import xlsxwriter

from pathlib import Path
    
random.seed(21)
# Functions for storing and retrieving Matrices used in the code below #################################################
# Write into preferenceList.xlsx the preference list of the users for each item
def store_r(r):
    prefListWb = xlsxwriter.Workbook(Path(__file__).parent / "Datasets/preferenceList.xlsx")
    prefListWs = prefListWb.add_worksheet()

    for j in range(0, r.shape[0]):
        prefListWs.write_row(j + 1, 0, r[j])
    prefListWb.close()

# Write into itemsCost.xlsx the itemsCost Matrix
def store_items_cost(itemsCost):
    itemsCostWb = xlsxwriter.Workbook(Path(__file__).parent / "Datasets/itemsCost.xlsx")
    itemsCostWs = itemsCostWb.add_worksheet()

    for j in range(0, itemsCost.shape[0]):
        itemsCostWs.write_number(j + 1, 0, itemsCost[j])
    itemsCostWb.close()


# Write into usersBudget.xlsx the usersBudget matrix
def store_users_budget(usersBudget):
    usersBudgetWb = xlsxwriter.Workbook(Path(__file__).parent / "Datasets/usersBudget.xlsx")
    usersBudgetWs = usersBudgetWb.add_worksheet()

    for j in range(0, usersBudget.shape[0]):
        usersBudgetWs.write_number(j + 1, 0, usersBudget[j])
    usersBudgetWb.close()

### Import some excel matrix
def import_excel_matrix(name):
    file_path = f"{Path(__file__).parent}/Datasets/{name}.xlsx"
    WS = pd.read_excel(file_path)
    return np.array(WS)

# Helping Functions used for showing and generating data ###############################################################
# Create a random group of groupSize users and return their index in prefList. usersSize is the max index possible
def create_random_group(groupSize, usersSize):
    return random.sample(range(0, usersSize + 1), groupSize)


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

