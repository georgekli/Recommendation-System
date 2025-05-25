import random
import numpy as np
import pandas as pd
import xlsxwriter

from pathlib import Path
from definitions import *

random.seed(21)

################## Functions for storing and retrieving Matrices used in the code below ################################

""" Import an excel matrix with name equal to name. Calamine is faster than default
:param filename: The file of the matrix
:returns: Imported matrix
:rtype: np.array
"""
def import_excel_matrix(filename):
    return np.array(pd.read_excel(filename, engine="calamine"))

""" Store a matrix in an excel sheet
:param matrix: The matrix to be stored
:param filename: Where to store the matrix
"""
def store_excel_matrix(matrix, filename):
    workbook = xlsxwriter.Workbook(f"{DATASET_DIR}/{filename}")
    worksheet = workbook.add_worksheet()
    try:
        for j in range(0, matrix.shape[0]):
            worksheet.write_row(j + 1, 0, matrix[j])
    except:
        for j in range(0, matrix.shape[0]):
            worksheet.write_number(j + 1, 0, matrix[j])

# Helping Functions used for showing and generating data ###############################################################
""" Create a random group of users
:param groupSize: The desired size of the group
:param usersSize: The maximum user index
:returns: random list of user 
""" 
def create_random_group(groupSize, usersSize):
    return random.sample(range(0, usersSize), groupSize)

# Calculate and show the average score of an algorithm that suggests an 
# item(1) by calculating the average rating of the item by users in each 
# group (used for 100 groups)
def calculate_average_score(preferedItems, prefList, groups, groupSize, number_of_groups=100):
    item_score = [0] * number_of_groups
    for k in range(number_of_groups):
        for j in range(0, groupSize):
            item_score[k] += prefList[groups[k][j]][preferedItems[k]]
        item_score[k] = item_score[k] / groupSize
    item_score_avg = np.sum(item_score) / number_of_groups
    print(f"Average score : {item_score_avg} | Group size : {groupSize} ")


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
def calculate_payments(group, selectedItem, prefList, itemsCost, usersBudget, show=True):
    # Cost Distribution Mechanism
    # Initialize the user satisfaction
    userSatisfaction = [0]*len(group)
    # For every user in the group
    for i, userId in enumerate(group):
        # User satisfaction comes from relevance metric
        userSatisfaction[i] = prefList[userId, selectedItem]
    # Calculate the overall similarity of the user satisfaction
    overallSimilarity = sum(userSatisfaction)
    # Initialize the user payments
    userpayments = [0]*len(group)
    richUsers = []
    sharedCost = 0
    # Calculate payment for each user
    for i, userId in enumerate(group):
        # Calculate how much each user pays based on its preference, satisfaction
        userpayments[i] = (userSatisfaction[i]/overallSimilarity)*itemsCost[selectedItem]
        richUsers.append([userId, i])
        if userpayments[i] > usersBudget[userId]:
            # Accumulate debt for the rich
            sharedCost += userpayments[i] - usersBudget[userId]
            userpayments[i] = usersBudget[userId]
            # Exclude poor user for future distribution
            richUsers.pop()
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
            print(f"User: {group[i]}, pays: {userpayments[i]}, for similarity {userSatisfaction[i]} with budget {usersBudget[group[i]]}")
        print(f"Movie Cost: {itemsCost[selectedItem]}")
        if sum(userpayments)-itemsCost[selectedItem] > 1e-10:
            print(f"Failed Distribution Test with {sum(userpayments)-itemsCost[selectedItem]} $ Difference")
    return userpayments


# Function that calculates the satisfaction of a user when itemId is purchased by the group
def calculate_satisfaction(prefList, userBudget, userId, itemId, payment, a=8, b=2):
    first_part = a ** ((-(max(prefList[userId])) - prefList[userId][itemId]) / max(prefList[userId]))
    second_part = b ** ((userBudget-payment)/userBudget)
    return first_part * second_part