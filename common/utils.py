import random
import numpy as np
import pandas as pd
import xlsxwriter

from pathlib import Path
from definitions import *

random.seed(21)

################## Functions for storing and retrieving Matrices used in the code below ################################

""" Import an excel matrix with name equal to name
:param filename: The file of the matrix
:returns: Imported matrix
:rtype: np.array
"""
def import_excel_matrix(filename):
    worksheet = pd.read_excel(filename)
    return np.array(worksheet)

""" Store a matrix in an excel sheet
:param matrix: The matrix to be stored
:param filename: Where to store the matrix
"""
def store_excel_matrix(matrix, filename):
    matrixWb = xlsxwriter.Workbook(f"{DATASET_DIR}/{filename}")
    matrixWs = matrixWb.add_worksheet()
    try:
        for j in range(0, matrix.shape[0]):
            matrixWs.write_row(j + 1, 0, matrix[j])
    except:
        for j in range(0, matrix.shape[0]):
            matrixWs.write_number(j + 1, 0, matrix[j])

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
# Print top k items for the first 50 users
def print_top_k(r, k=10):
    sortedItemsIndexes = np.zeros(r.shape[1])
    top_k_items = np.zeros((r.shape[0], k))
    top_k_indexes = np.zeros((r.shape[0], k))
    print(f"Top {k} items for every user(descending)")
    for users in range(0, 50):
        # Take indexes of sorted items
        sortedItemsIndexes = np.argsort(r[users])
        top_k_indexes[users] = np.take(sortedItemsIndexes, range(r.shape[1] - k, r.shape[1]))
        # Take top 10 indexes and make them items
        top_k_items[users] = np.take(r[users], np.take(sortedItemsIndexes, range(r.shape[1] - k, r.shape[1])))
        print("User: ", users, np.flip(top_k_indexes[users]))


# Function that returns the items that the users of group can acquire with their budget
def items_feasible(group, items, users):
    # Find budget for group
    budget = 0
    for userId in group:
        budget = budget + users[userId]
    # search every item for feasibility for lowest budget
    feasibleItemsIndexes = items < budget
    feasibleItems = np.where(feasibleItemsIndexes)[0]
    return feasibleItems


# Function that calculates and shows the payment vector (showing only if boolean var show is True) for a group that
# gets an item (=selectedItem)
def calculate_payments(group, selectedItem, prefList, itemsCost, usersBudget, show):
    # Cost Distribution Mechanism
    i = 0
    # Initialize the user satisfaction
    userSatisfaction = [0]*len(group)
    # For every user in the group
    for userId in group:
        # User satisfaction comes from relevance metric
        userSatisfaction[i] = prefList[userId, selectedItem]
        i += 1
    # Calculate the overall similarity of the user satisfaction
    overallSimilarity = sum(userSatisfaction)
    # Initialize the user payments
    userpayments = [0]*len(group)
    richUsers = []
    sharedCost = 0
    i = 0
    # Calculate payment for each user
    for userId in group:
        # Calculate how much each user pays based on its preference, satisfaction
        userpayments[i] = (userSatisfaction[i]/overallSimilarity)*itemsCost[selectedItem]
        richUsers.append([userId, i])
        if userpayments[i] > usersBudget[userId]:
            # Accumulate debt for the rich
            sharedCost += userpayments[i] - usersBudget[userId]
            userpayments[i] = usersBudget[userId]
            # Exclude poor user for future distribution
            richUsers.pop()
        i += 1
    # Well now the rich should pay for the poor recursively (ancient Athens theatre)
    while sharedCost != 0:
        newSharedCost = 0
        newRichUsers = []
        richUsersIdx = [index[1] for index in richUsers]
        overallSimilarity = 0
        for idTmp in richUsersIdx:
            overallSimilarity += userSatisfaction[idTmp]
        for user in richUsers:
            i = user[1]
            userId = user[0]
            userpayments[i] += (userSatisfaction[i] / overallSimilarity) * sharedCost  # Simple distribution metric
            newRichUsers.append([userId, i])
            if userpayments[i] > usersBudget[userId]:
                newSharedCost += userpayments[i] - usersBudget[userId]  # Accumulate debt for the rich
                userpayments[i] = usersBudget[userId]
                newRichUsers.pop()  # Exclude poor user for future distribution
        sharedCost = newSharedCost
        richUsers = newRichUsers
    if show:
        for i in range(0, len(group)):
            print("User:", group[i], " pays:", userpayments[i], "for similarity", userSatisfaction[i], " with budget",
                  usersBudget[group[i]])
        print("Cost of movie:", itemsCost[selectedItem])
        if sum(userpayments)-itemsCost[selectedItem] > 1e-10:
            print("Failed Distribution Test with", sum(userpayments)-itemsCost[selectedItem], "$ Difference")
    return userpayments


# Function that calculates the satisfaction of a user when itemId is purchased by the group
def calculate_sat(prefList, userBudget, userId, itemId, payment, a=8, b=2):
    first_part = a ** ((-(max(prefList[userId])) - prefList[userId][itemId]) / max(prefList[userId]))
    second_part = b ** ((userBudget-payment)/userBudget)
    return first_part * second_part