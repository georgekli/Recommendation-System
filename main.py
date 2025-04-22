import random
import numpy as np
from numpy.linalg import inv
import threading
from multiprocessing import Process, Array
import time

import xlrd
import corankco as crc

from definitions import *
from package.utils import *
from package.voting_algorithms import *

# Control Sequence Variables############################################################################################
# Array correspond to Tasks [ Task1, Task2a, Task2b, Task3, Task4, Task5]
EPSILON = 0.0001
IMPORT_MATRICES = False
SPEND_TIME_WAITING = [True, False, False, False, False, False]
    
random.seed(21)

# Function that spawns threads calculating copeland_method (through group_set_copeland) for different group
# sizes (= userNum). Specifically for group sizes = 5, 10, 15, 20
class MyBigThread (threading.Thread):
    def __init__(self, threadID, userNum):
        threading.Thread.__init__(self)
        self.threadID = threadID
        self.userNum = userNum

    def run(self):
        print("Starting ", self.threadID)
        group_set_copeland(self.userNum)
        print("Exiting ", self.threadID)

# Function that spawns threads calculating copeland_method (through group_set_copeland) for
# group size = KENDALL_TAU_GROUP_SIZE
class MyBigThread2 (threading.Thread):
    def __init__(self, threadID, userNum, groups):
        threading.Thread.__init__(self)
        self.threadID = threadID
        self.userNum = userNum
        self.groups = groups

    def run(self):
        print("Starting ", self.threadID)
        for i in range(0, 100):
            winner = copeland_method(self.groups[i], r)
            winnersArray[i] = winner
        calculate_avg_algo_score(winnersArray, r, groups, self.userNum)
        print("Exiting ", self.threadID)

# Function that spawn a process for each group to be created
def group_them(firstUserPrefernce, firstUser, r, simGroup, divGroup, groupSize, pid):
    cond = 1
    simUsers = []
    simUsers.append(firstUser)
    divUsers = []
    divUsers.append(firstUser)
    divThreshold = 0.6
    simThreshold = 0.6
    while cond != 0:
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
        elif (per >= divThreshold) and (len(divUsers) < groupSize):
            divUsers.append(compareUser)
        elif (len(simUsers) == groupSize) and (len(divUsers) == groupSize):
            print(pid)
            print("divUsers: ", divUsers)
            print("simUsers: ", simUsers)
            cond = 0
    for k in range(0, len(divUsers)):
        simGroup[k] = simUsers[k]
        divGroup[k] = divUsers[k]

# Main code for running Tasks in the assignment ########################################################################
if __name__ == '__main__':
    # Control Sequence Variables 
    storeMatrices = not IMPORT_MATRICES
    spendTimeWaiting = SPEND_TIME_WAITING
    if IMPORT_MATRICES:
        r = import_excel_matrix(f"{DATASET_DIR}/preferenceList.xlsx")
        itemsCost = import_excel_matrix(f"{DATASET_DIR}/itemsCost.xlsx")
        usersBudget = import_excel_matrix(f"{DATASET_DIR}/usersBudget.xlsx")
    else:
        # Open Workbook for the users
        usersWb = xlrd.open_workbook(f"{DATASET_DIR}/items.xls")
        usersSheet = usersWb.sheet_by_index(0)

        # To open Workbook for the items
        itemsWb = xlrd.open_workbook(f"{DATASET_DIR}/users.xls")
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
        if storeMatrices:
            store_excel_matrix(r, f"{DATASET_DIR}/preferenceList.xlsx")
            store_excel_matrix(itemsCost, f"{DATASET_DIR}/itemsCost.xlsx")
            store_excel_matrix(usersBudget, f"{DATASET_DIR}/usersBudget.xlsx")
    
    # Initialize the sorted Preference list
    start1 = time.time()
    prefList = r
    sortedPref = np.zeros((r.shape[0], r.shape[1]), dtype=tuple)
    for i in range(0, r.shape[0]):
        for j in range(0, r.shape[1]):
            sortedPref[i][j] = (prefList[i][j], j)
    # Sort the list
    sortedPref = np.flip(np.sort(sortedPref), axis=1)
    end1 = time.time()
    print(end1 - start1)
    

    # Task 1############################################################################################################
    if spendTimeWaiting[0]:
        print_top_k(r, k=5)
    # Task 2a###########################################################################################################
    if spendTimeWaiting[1]:
        groupSizes = [5, 10, 15, 20]
        # Create Random Groups and execute Borda count
        preferedItems = [[0] * 100] * len(groupSizes)
        # Iterate through different sets of groups
        for i in range(0, len(groupSizes)):
            groups = [[0] * groupSizes[i]] * 100
            # Iterate through groups
            for numOfGroups in range(0, 100):
                groups[numOfGroups] = (create_random_group(groupSizes[i], r.shape[0]))
            # Run Borda Algo
            preferedItems[i] = borda_count(groups, sortedPref)
            # Calculate average score
            print("Borda Count Results:")
            calculate_avg_algo_score(preferedItems[i], prefList, groups, groupSizes[i])
        winnersArray = [[0] * 100] * 4
        # Execute copeland method
        print("Copeland Method Results:")
        thread1 = MyBigThread(1, 5)
        thread2 = MyBigThread(2, 10)
        thread3 = MyBigThread(3, 15)
        thread4 = MyBigThread(4, 20)
        winnersArray[0] = thread1.start()
        winnersArray[1] = thread2.start()
        winnersArray[2] = thread3.start()
        winnersArray[3] = thread4.start()
    # Task 2b###########################################################################################################
    if spendTimeWaiting[2]:
        k = 10
        threshold = 6
        groupSizes = [5, 10, 15, 20]
        # S = [[[] for i in range(0, k)]*100]*len(groupSizes)
        for sizes in range(0, len(groupSizes)):
            S = [[-1]*k]*100
            groups = [0]*100
            for i in range(0, 100):
                groups[i] = create_random_group(groupSizes[sizes], prefList.shape[0])
                S[i] = rav(groups[i], prefList, k, threshold)
                print("For group size:", groupSizes[sizes], " recommended items are ", S[i])
        print("\nFor example for the last group of group size 20 recommended items are", S[99])
    # Task 3############################################################################################################
    if spendTimeWaiting[3]:
        # Define the number of items to recommend
        groupSize = KENDALL_TAU_GROUP_SIZE
        groupSizes = [groupSize]
        # Get the users size from the excel
        userSize = r.shape[0]
        numberOfGroups = 100
        groups = [[[0 for k in range(0, groupSize)] for j in range(0, 2)] for i in range(0, numberOfGroups)]
        p = []
        simGroup = [0] * numberOfGroups
        divGroup = [0] * numberOfGroups
        for groupsIdx in range(0, numberOfGroups):
            # Index of the first user.(Random)
            firstUser = random.randint(0, userSize - 1)
            firstUserPrefernce = r[firstUser]
            simGroup[groupsIdx] = Array('i', range(0, groupSize))
            divGroup[groupsIdx] = Array('i', range(0, groupSize))
            p.append(Process(target=group_them, args=(firstUserPrefernce, firstUser, r, simGroup[groupsIdx],
                                                      divGroup[groupsIdx], groupSize, groupsIdx,)))
            p[groupsIdx].start()
        for pr in p:
            pr.join()
        newDivGroups = [[0]*KENDALL_TAU_GROUP_SIZE]*numberOfGroups
        newSimGroups = [[0]*KENDALL_TAU_GROUP_SIZE]*numberOfGroups
        for s in range(0, numberOfGroups):
            newSimGroups[s] = simGroup[s][:]
            newDivGroups[s] = divGroup[s][:]
        simPreferedItems = [[0] * 100] * len(groupSizes)
        divPreferedItems = [[0] * 100] * len(groupSizes)
        print("Borda Count Results:")
        # Iterate through different sets of groups
        for i in range(0, len(groupSizes)):
            # Iterate through groups
            # Run Borda Algo
            simPreferedItems[i] = borda_count(newSimGroups, sortedPref)
            divPreferedItems[i] = borda_count(newDivGroups, sortedPref)
            # Calculate average score
            print("Borda Count Results for similar groups:")
            calculate_avg_algo_score(simPreferedItems[i], prefList, newSimGroups, groupSizes[i])
            print("Borda Count Results for divergent groups:")
            calculate_avg_algo_score(divPreferedItems[i], prefList, newDivGroups, groupSizes[i])
        winnersArray = [0] * 100
        # Execute copeland method
        print("Copeland Method Results for similar users:")
        for i in range(0, 100):
            winner = copeland_method(newSimGroups[i], r)
            winnersArray[i] = winner
        calculate_avg_algo_score(winnersArray, r, newSimGroups, KENDALL_TAU_GROUP_SIZE)
        print("Copeland Method Results for divergent users:")
        for i in range(0, 100):
            winner = copeland_method(newDivGroups[i], r)
            winnersArray[i] = winner
        calculate_avg_algo_score(winnersArray, r, newDivGroups, KENDALL_TAU_GROUP_SIZE)
    # Task 4############################################################################################################
    # Provide user's budgets with user ID
    if spendTimeWaiting[4]:
        groupSize = 7
        # Create the random groups
        group = create_random_group(groupSize, prefList.shape[0])
        # Find the feasible items for the particular groups
        feasibleItems = items_feasible(group, itemsCost, usersBudget)
        if len(feasibleItems) == 0:
            print("Budget not enough")
            exit()
        # Choose a random item from the feasible items
        selectedItemIdx = random.randint(0, feasibleItems.shape[0]-1)
        selectedItem = feasibleItems[selectedItemIdx]
        payments = calculate_payments(group, selectedItem, prefList, itemsCost, usersBudget, True)
    # Task 5############################################################################################################
    if spendTimeWaiting[5]:
        groupSizes = [4, 6, 8, 10, 12]
        groupsSat = [0]*len(groupSizes)
        for k in range(0, len(groupSizes)):
            tmpAvgSat = [0]*100
            for j in range(0, 100):
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
                    totalSAT = 0
                    i = 0
                    for userId in group:
                        totalSAT += calculate_sat(prefList, usersBudget[userId], userId, selectedItem, payments[i])
                        i += 1
                    if final < totalSAT:
                        final = totalSAT
                        satItem = selectedItem
                # print("Item: ", satItem)
                # print("SAT: ", final)
                tmpAvgSat[j] = final
                # print("Fisibles: ", feasibleItems)
            groupsSat[k] = sum(tmpAvgSat)/100
            print(f"For group size: {groupSizes[k]}, average satisfaction is:{groupsSat[k]}")
