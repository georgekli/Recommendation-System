import random
import numpy as np
from numpy.linalg import inv
import time
import threading
from multiprocessing import Process, Array

import xlrd
import corankco as crc

from definitions import *
from common.utils import *
from common.voting_algorithms import *

'''These are for code profiling to make it faster'''
import pstats
import cProfile

'''Control Sequence variables'''
# Array correspond to Tasks [ Task1, Task2a, Task2b, Task3, Task4, Task5]
TOTAL_GROUPS = 100
IMPORT_MATRICES = False
STORE_MATRICES = False
SPEND_TIME_WAITING = [True, True, True, True, True, True]
    
random.seed(21)

# Function that spawns threads calculating copeland_method (through group_set_copeland) for different group
# sizes (= userNum). Specifically for group sizes = 5, 10, 15, 20
class MyBigThread (threading.Thread):
    def __init__(self, thread_id, userNum):
        threading.Thread.__init__(self)
        self.thread_id = thread_id
        self.userNum = userNum

    def run(self):
        print(f"Starting {self.thread_id}")
        group_set_copeland(self.userNum)
        print(f"Exiting {self.thread_id}")

# Function that spawns threads calculating copeland_method (through group_set_copeland) for
# group size = KENDALL_TAU_GROUP_SIZE
class MyBigThread2 (threading.Thread):
    def __init__(self, thread_id, userNum, groups):
        threading.Thread.__init__(self)
        self.thread_id = thread_id
        self.userNum = userNum
        self.groups = groups

    def run(self):
        print("Starting ", self.thread_id)
        for i in range(TOTAL_GROUPS):
            winner = copeland_method(self.groups[i], r)
            winnersArray[i] = winner
        calculate_average_score(winnersArray, r, groups, self.userNum)
        print("Exiting ", self.thread_id)

# Function that spawn a process for each group to be created
def group_them(firstUserPrefernce, firstUser, r, simGroup, divGroup, groupSize, pid):
    condition = True
    simUsers, divUsers = [], []
    simUsers.append(firstUser)
    divUsers.append(firstUser)
    divThreshold = simThreshold = 0.6

    while condition:
        while True:
            tmp = random.randint(0, r.shape[0] - 1)
            if tmp not in divUsers:
                if tmp not in simUsers:
                    compareUser = tmp
                    break
        counter = 0
        compareUserPreference = r[compareUser]
        for i in range(0, r.shape[1] - 1):
            for j in range(i + 1, r.shape[1]):
                if (((firstUserPrefernce[i] > firstUserPrefernce[j]) and
                     (compareUserPreference[i] < compareUserPreference[j]))
                    or ((firstUserPrefernce[i] <= firstUserPrefernce[j]) and
                        (compareUserPreference[i] >= compareUserPreference[j]))):
                    counter = counter + 1
        per = counter / ((r.shape[1] * (r.shape[1] - 1)) / 2)
        if (per < 1 - simThreshold) and (len(simUsers) < groupSize):
            simUsers.append(compareUser)

        if (per >= divThreshold) and (len(divUsers) < groupSize):
            divUsers.append(compareUser)

        if (len(simUsers) == groupSize) and (len(divUsers) == groupSize):
            print(pid)
            print("Divergent Users: ", divUsers)
            print("Similar Users: ", simUsers)
            condition = False

    for k in range(0, len(divUsers)):
        simGroup[k] = simUsers[k]
        divGroup[k] = divUsers[k]

    
def main():# Control Sequence Variables 
    spendTimeWaiting = SPEND_TIME_WAITING
    if IMPORT_MATRICES:
        r = import_excel_matrix(f"{DATASET_DIR}/preferenceList.xlsx")
        itemsCost = import_excel_matrix(f"{DATASET_DIR}/itemsCost.xlsx")
        usersBudget = import_excel_matrix(f"{DATASET_DIR}/usersBudget.xlsx")
    else:
        # Open Workbook for the users
        usersWb = xlrd.open_workbook(f"{DATASET_DIR}/users.xls")
        usersSheet = usersWb.sheet_by_index(0)

        # To open Workbook for the items
        itemsWb = xlrd.open_workbook(f"{DATASET_DIR}/items.xls")
        itemsSheet = itemsWb.sheet_by_index(0)
        # Define the dimensions of the features
        D = 8
        rmax = 10
        # Create the Variance matrices of the users and the items
        itemsSigma = np.identity(D)
        usersSigma = 2 * np.identity(D)
        # Create the inverse of the Variance matrices of the users and the items
        itemsSigmaInv = inv(itemsSigma)
        usersSigmaInv = inv(usersSigma)
        # Initialize the array that the preference list will be stored
        r = np.zeros((usersSheet.nrows - 1, itemsSheet.nrows - 1), dtype=float)
        # Initialize array dor itemsCost
        itemsCost = np.zeros((itemsSheet.nrows - 1), dtype=int)
        # Initialize array dor usersBudget
        usersBudget = np.zeros((usersSheet.nrows - 1), dtype=int)
        # Produce non-variable value of the KL value
        conVal = 0.5 * np.log(np.linalg.det(np.matmul(usersSigmaInv, itemsSigma))) + 0.5 * np.trace(
            inv(np.matmul(usersSigmaInv, itemsSigma))) - D / 2
        for i in range(1, itemsSheet.nrows):
            itemsCost[i - 1] = itemsSheet.cell_value(i, 10)
            for u in range(1, usersSheet.nrows):
                # Initialize the vectors of each user and item
                mI = np.zeros((D, 1))
                mU = np.zeros((D, 1))
                count = 0
                # Get the features from a user and an item
                for j in range(2, 10):
                    mI[count] = itemsSheet.cell_value(i, j)
                    mU[count] = usersSheet.cell_value(u, j)
                    count = count + 1

                klTmp = conVal + 0.5 * np.matmul(np.matmul((mU - mI).transpose(), itemsSigmaInv), (mU - mI))
                sc = rmax - (klTmp / rmax)
                if sc <= 0:
                    sc = EPSILON
                r[u - 1, i - 1] = sc
                if i == 1:
                    usersBudget[u - 1] = usersSheet.cell_value(u, 10)
        # Store matrices?
        if STORE_MATRICES:
            store_excel_matrix(r, f"{DATASET_DIR}/preferenceList.xlsx")
            store_excel_matrix(itemsCost, f"{DATASET_DIR}/itemsCost.xlsx")
            store_excel_matrix(usersBudget, f"{DATASET_DIR}/usersBudget.xlsx")
    
    # Initialize the sorted Preference list
    prefList = r
    sortedPref = np.zeros((r.shape[0], r.shape[1]), dtype=tuple)
    for i in range(0, sortedPref.shape[0]):
        for j in range(0, sortedPref.shape[1]):
            sortedPref[i][j] = (prefList[i][j], j)
    # Sort the list
    sortedPref = np.flip(np.sort(sortedPref), axis=1)

    # Task 1
    if spendTimeWaiting[0]:
        print_top_k(r, k=5)
    # Task 2a 
    if spendTimeWaiting[1]:
        groupSizes = [5, 10, 15, 20]
        # Create Random Groups and execute Borda count
        preferedItems = [[0] * TOTAL_GROUPS] * len(groupSizes)
        # Iterate through different sets of groups
        for i in range(0, len(groupSizes)):
            groups = [[0] * groupSizes[i]] * TOTAL_GROUPS
            # Iterate through groups
            for numOfGroups in range(TOTAL_GROUPS):
                groups[numOfGroups] = (create_random_group(groupSizes[i], r.shape[0]))
            # Run Borda Algo
            preferedItems[i] = borda_count(groups, sortedPref)
            # Calculate average score
            print("Borda Count Results:")
            calculate_average_score(preferedItems[i], prefList, groups, groupSizes[i], TOTAL_GROUPS)
        winnersArray = [[0] * TOTAL_GROUPS] * 4
        # Execute copeland method
        print("Copeland Method Results:")
        threads = list()
        for i, group_size in enumerate(groupSizes):
            threads.append(MyBigThread(i, group_size))

        for i, thread in enumerate(threads):
            winnersArray[i] = thread.start()

    # Task 2b############
    if spendTimeWaiting[2]:
        threshold, k = 6, 10
        groupSizes = [5, 10, 15, 20]

        # S = [[[] for i in range(0, k)]*100]*len(groupSizes)
        for sizes in range(0, len(groupSizes)):
            S = [[-1]*k] * TOTAL_GROUPS
            groups = [0] * TOTAL_GROUPS
            for i in range(TOTAL_GROUPS):
                groups[i] = create_random_group(groupSizes[sizes], prefList.shape[0])
                S[i] = reweighted_approval_voting(groups[i], prefList, k, threshold)
                print(f"For group size: {groupSizes[sizes]} recommended items are {S[i]}")
        print(f"\nFor example for the last group of group size 20 recommended items are {S[99]}")
    # Task 3###########
    if spendTimeWaiting[3]:
        # Define the number of items to recommend
        groupSize = KENDALL_TAU_GROUP_SIZE
        groupSizes = [groupSize]
        # Get the users size from the excel
        userSize = r.shape[0]

        groups = [[[0 for k in range(groupSize)] for j in range(0, 2)] for i in range(TOTAL_GROUPS)]
        p = []
        simGroup = [0] * TOTAL_GROUPS
        divGroup = [0] * TOTAL_GROUPS
        for groupsIdx in range(TOTAL_GROUPS):
            # Index of the first user.(Random)
            firstUser = random.randint(0, userSize - 1)
            firstUserPrefernce = r[firstUser]
            simGroup[groupsIdx] = Array('i', range(groupSize))
            divGroup[groupsIdx] = Array('i', range(groupSize))
            p.append(Process(target=group_them, args=(firstUserPrefernce, firstUser, r, simGroup[groupsIdx],
                                                      divGroup[groupsIdx], groupSize, groupsIdx,)))
            p[groupsIdx].start()
        for pr in p:
            pr.join()
        newDivGroups = [[0] * KENDALL_TAU_GROUP_SIZE] * TOTAL_GROUPS
        newSimGroups = [[0] * KENDALL_TAU_GROUP_SIZE] * TOTAL_GROUPS
        for s in range(TOTAL_GROUPS):
            newSimGroups[s] = simGroup[s][:]
            newDivGroups[s] = divGroup[s][:]
        simPreferedItems = [[0] * TOTAL_GROUPS] * len(groupSizes)
        divPreferedItems = [[0] * TOTAL_GROUPS] * len(groupSizes)

        print("Borda Count Results:")
        # Iterate through different sets of groups
        for i in range(len(groupSizes)):
            # Run Borda Algo
            simPreferedItems[i] = borda_count(newSimGroups, sortedPref)
            divPreferedItems[i] = borda_count(newDivGroups, sortedPref)
            # Calculate average score
            print("Borda Outcome (Similar groups):")
            calculate_average_score(simPreferedItems[i], prefList, newSimGroups, groupSizes[i])
            print("Borda Outcome (Divergent groups):")
            calculate_average_score(divPreferedItems[i], prefList, newDivGroups, groupSizes[i])

        winnersArray = [0] * TOTAL_GROUPS
        # Execute copeland method
        print("Copeland Outcome (Similar users): ")
        for i in range(TOTAL_GROUPS):
            winner = copeland_method(newSimGroups[i], r)
            winnersArray[i] = winner
        calculate_average_score(winnersArray, r, newSimGroups, KENDALL_TAU_GROUP_SIZE)

        print("Copeland Outcome (Divergent users): ")
        for i in range(TOTAL_GROUPS):
            winner = copeland_method(newDivGroups[i], r)
            winnersArray[i] = winner
        calculate_average_score(winnersArray, r, newDivGroups, KENDALL_TAU_GROUP_SIZE)
    # Task 4#######
    # Provide user's budgets with user ID
    if spendTimeWaiting[4]:
        groupSize = 7
        # Create the random groups
        group = create_random_group(groupSize, prefList.shape[0])
        # Find the feasible items for the particular groups
        feasibleItems = items_feasible(group, itemsCost, usersBudget)
        if len(feasibleItems) == 0:
            print("Budget not enough")
            print("Exiting .")
            time.sleep(1)
            print("Exiting ..")
            time.sleep(1)
            print("Exiting ...")
            time.sleep(1)
            exit()
        # Choose a random item from the feasible items
        selectedItemIdx = random.randint(0, feasibleItems.shape[0]-1)
        selectedItem = feasibleItems[selectedItemIdx]
        payments = calculate_payments(group, selectedItem, prefList, itemsCost, usersBudget, True)

    # Task 5######
    if spendTimeWaiting[5]:
        groupSizes = [4, 6, 8, 10, 12]
        group_satisfaction = [0] * len(groupSizes)
        for k in range(0, len(groupSizes)):
            tmpAvgSat = [0] * TOTAL_GROUPS
            for j in range(TOTAL_GROUPS):
                groupSize = groupSizes[k]
                final = 0
                satItem = 0
                feasibleItems = []
                # Create the random groups
                while len(feasibleItems) == 0:
                    group = create_random_group(groupSize, prefList.shape[0])
                    # Find the feasible items for the particular groups
                    feasibleItems = items_feasible(group, itemsCost, usersBudget)
                for selectedItem in feasibleItems:
                    payments = calculate_payments(group, selectedItem, prefList, itemsCost, usersBudget, False)
                    total_satisfaction = 0
                    for i, userId in enumerate(group):
                        total_satisfaction += calculate_satisfaction(prefList, usersBudget[userId], userId, selectedItem, payments[i])

                    if final < total_satisfaction:
                        final = total_satisfaction
                        satItem = selectedItem
                # print("Item: ", satItem)
                # print("SAT: ", final)
                tmpAvgSat[j] = final
                # print("Fisibles: ", feasibleItems)
            group_satisfaction[k] = sum(tmpAvgSat)/ TOTAL_GROUPS
            print(f"For group size: {groupSizes[k]}, average satisfaction is:{group_satisfaction[k]}")

# Main code for running Tasks in the assignment 
if __name__ == '__main__':
    cProfile.run('main()', "main_statistics")

    p = pstats.Stats("main_statistics")
    p.sort_stats("cumulative").reverse_order().print_stats()