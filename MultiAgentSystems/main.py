import random
import threading
from multiprocessing import Process, Array
from pathlib import Path
import xlrd
import xlsxwriter
import pandas
import numpy as np
from numpy.linalg import inv

# Control Sequence Variables############################################################################################
# Array correspond to Tasks [ Task1, Task2a, Task2b, Task3, Task4, Task5]
SPEND_TIME_WAITING = [True, False, False, False, False, False]
EPSILON = 0.0001
IMPORT_MATRICES = False
KENDALL_TAU_GROUP_SIZE = 10


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


# Import the preference list matrix
def import_r():
    WS = pandas.read_excel(Path(__file__).parent / "Datasets/preferenceList.xlsx")
    # WS = pandas.read_excel('D:\\TUC\\THL_311\\pythonProject\\Datasets\\smallPreferencedList.xlsx')
    r = np.array(WS)
    return r


# Import the itemsCost Matrix
def import_items_cost():
    WS = pandas.read_excel(Path(__file__).parent / "Datasets/itemsCost.xlsx")
    itemsCost = np.array(WS)
    return itemsCost


# Import the usersBudget matrix
def import_users_budget():
    WS = pandas.read_excel(Path(__file__).parent / "Datasets/usersBudget.xlsx")
    usersBudget = np.array(WS)
    return usersBudget


# Helping Functions used for showing and generating data ###############################################################
# Create a random group of groupSize users and return their index in prefList. usersSize is the max index possible
def create_random_group(groupSize, usersSize):
    random.seed()
    group = [-1] * groupSize
    for i in range(0, groupSize):
        userIndex = random.randint(0, usersSize - 1)
        while (userIndex in group):
            userIndex = random.randint(0, usersSize - 1)
        group[i] = userIndex
    return group


# Calculate and show the average score of an algorithm that suggests an item(1) by calculating the average rating of the
# item by users in each group (used for 100 groups)
def calculate_avg_algo_score(preferedItems, prefList, groups, groupSize):
    itemScore = [0] * 100
    for k in range(0, 100):
        for j in range(0, groupSize):
            itemScore[k] += prefList[groups[k][j]][preferedItems[k]]
        itemScore[k] = itemScore[k] / groupSize
    avgItemScore = np.sum(itemScore) / 100
    print("For group size = ", groupSize, " average score is ", avgItemScore)


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


# Borda Count Algorithm for 100 groups in groupsIndexes returning 100 recommended items
def borda_count(groupsIndexes, prefList):
    preferedItems = [0] * 100
    for i in range(0, 100):
        groupIndexes = groupsIndexes[i]
        group = [[0] * prefList.shape[1]] * len(groupIndexes)
        # Make groups from matrix
        for j in range(0, len(groupIndexes)):
            group[j] = prefList[groupIndexes[j]]
        # Initialize counters
        itemsRating = [0] * prefList.shape[1]
        # Count for every user prefered sequence
        for users in range(0, len(groupIndexes)):
            user = group[users]
            for items in range(0, prefList.shape[1]):
                item = user[items][1]
                # Borda count increment
                itemsRating[item] = itemsRating[item] + (prefList.shape[1] - items)
        preferedItems[i] = np.flip(np.argsort(itemsRating))[0]
    return preferedItems


# Copeland Method Algorithm for a group returning a recommended item
def copeland_method(groupIndexes, prefList):
    # Create the copeland matrix
    itemWins = [0]*prefList.shape[1]
    itemA = 0
    itemB = 1
    while itemA < (prefList.shape[1] - 1) and itemB < (prefList.shape[1]):
        roundWins = [0] * 2
        for i in range(0, len(groupIndexes)):
            if prefList[groupIndexes[i]][itemA] > prefList[groupIndexes[i]][itemB]:
                roundWins[0] += 1
            elif prefList[groupIndexes[i]][itemA] == prefList[groupIndexes[i]][itemB]:
                pass
            else:
                roundWins[1] += 1

        if roundWins[0] > roundWins[1]:
            itemWins[itemA] += 1
        elif roundWins[0] < roundWins[1]:
            itemWins[itemB] += 1
        else:
            itemWins[itemA] += 0.5
            itemWins[itemB] += 0.5
        itemB += 1
        if itemB == prefList.shape[1]-1:
            itemA = itemA + 1
            itemB = itemA + 1

    return itemWins.index(max(itemWins))


# Function calling copeland_method() for 100 groups of groupSize and showing the results
def group_set_copeland(groupSize):
    # Initialize the winners array and the groups
    groups = [0] * 100
    winnersArray = [0] * 100
    # For 100 groups calculate with the copeland method the winner items
    for i in range(0, 100):
        groups[i] = create_random_group(groupSize, r.shape[0])
        winner = copeland_method(groups[i], r)
        winnersArray[i] = winner
    calculate_avg_algo_score(winnersArray, r, groups, groupSize)


# Function that spawns threads calculating copeland_method (through group_set_copeland) for different group
# sizes (= userNum). Specifically for group sizes = 5, 10 , 15, 20
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


# Reweighed Approval Voting algorithm for a groups that returns k items
def rav(groupIndexes, prefList, k, threshold):
    # Create approval list for every user
    A = [[] for i in range(len(groupIndexes))]
    for i in range(0, len(groupIndexes)):
        # Get for the i user. Its preference list
        userItems = prefList[groupIndexes[i]]
        # For all the items j in its preference list
        for j in range(0, len(userItems)):
            # Check if the rating is greater than the threshold
            if userItems[j] > threshold:  # If yes insert in the approval list
                nest = A[i]
                nest.append(j)
    # print(A)
    S = []
    # Recommend k Items
    for kIters in range(0, k):
        # Item votes
        weightedItemVotes = [0]*prefList.shape[1]
        # For all users in the group
        for i in range(0, len(groupIndexes)):
            # Get for the i user. Its preference list
            userApprovedItems = A[i]
            electedGroup = 0
            # Calculate mumber of items of A[i] in S
            for item in userApprovedItems:
                if item in S:
                    electedGroup += 1
            # For all the items j in its prefence list
            for j in range(0, len(userApprovedItems)):
                if userApprovedItems[j] not in S:
                    # weighted vote
                    weightedItemVotes[userApprovedItems[j]] += (1/(electedGroup+1))
        # Find winner
        electedCanditate = weightedItemVotes.index(max(weightedItemVotes))
        S.append(electedCanditate)
        # print(weightedItemVotes)
        # print(electedCanditate)
    return S


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
            random.seed()
            tmp = random.randint(0, r.shape[0] - 1)
            if tmp not in divUsers:
                if tmp not in simUsers:
                    compareUser = tmp
                    break
        counter = 0;
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
        for i in range(0, groupSize):
            print("User:", group[i], " pays:", userpayments[i], "for similarity", userSatisfaction[i], " with budget",
                  usersBudget[group[i]])
        print("Cost of movie:", itemsCost[selectedItem])
        if sum(userpayments)-itemsCost[selectedItem] > 1e-10:
            print("Failed Distribution Test with", sum(userpayments)-itemsCost[selectedItem], "$ Difference")
    return userpayments


# Function that calculates the satisfaction of a user when itemId is purchased by the group
def calculate_sat(prefList, userBudget, userId, itemId, payment):
    a = 8
    b = 2
    fisrtPart = a**((-(max(prefList[userId]))-prefList[userId][itemId])/max(prefList[userId]))
    secondPart = b**((userBudget-payment)/userBudget)
    return fisrtPart*secondPart


# Main code for running Tasks in the assignment ########################################################################
if __name__ == '__main__':
    # Control Sequence Variables
    impMatrices = IMPORT_MATRICES
    storeMatrices = not impMatrices
    spendTimeWaiting = SPEND_TIME_WAITING
    if impMatrices:
        r = import_r()
        itemsCost = import_items_cost()
        usersBudget = import_users_budget()
    else:
        # Give the location of the file of the items
        #itemsLoc = 'D:\\TUC\\THL_311\\pythonProject\\Datasets\\items.xls'
        itemsLoc = Path(__file__).parent / "Datasets/items.xls"
        # Give the location of the file of the users
        # usersLoc = 'D:\\TUC\\THL_311\\pythonProject\\Datasets\\users.xls'
        usersLoc = Path(__file__).parent / "Datasets/users.xls"

        # To open Workbook for the users
        usersWb = xlrd.open_workbook(usersLoc)
        usersSheet = usersWb.sheet_by_index(0)

        # To open Workbook for the items
        itemsWb = xlrd.open_workbook(itemsLoc)
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
        # Helpful Storing
        if storeMatrices:
            store_r(r)
            store_items_cost(itemsCost)
            store_users_budget(usersBudget)
    # Initialize the sorted Preference list
    prefList = r
    sortedPref = np.zeros((r.shape[0], r.shape[1]), dtype=tuple)
    for i in range(0, r.shape[0]):
        for j in range(0, r.shape[1]):
            sortedPref[i][j] = (prefList[i][j], j)
    # Sort the list
    sortedPref = np.flip(np.sort(sortedPref), axis=1)
    # Task 1############################################################################################################
    if spendTimeWaiting[0]:
        print_top10(r)
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
        random.seed()
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
            print("For group size:", groupSizes[k], " average satisfaction is:", groupsSat[k])
