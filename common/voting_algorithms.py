import numpy as np
from common.utils import create_random_group, calculate_average_score

# Borda Count Algorithm for 100 groups in groupsIndexes returning 100 recommended items
def borda_count(groupsIndexes, prefList, number_of_groups=100):
    preferedItems = [0] * number_of_groups
    for i in range(number_of_groups):
        groupIndexes = groupsIndexes[i]
        group = [[0] * prefList.shape[1]] * len(groupIndexes)
        # Make groups from matrix
        for j in range(0, len(groupIndexes)):
            group[j] = prefList[groupIndexes[j]]
        # Initialize counters
        itemsRating = [0] * prefList.shape[1]
        # Count for every user prefered sequence
        for users in range(len(groupIndexes)):
            user = group[users]
            for items in range(prefList.shape[1]):
                item = user[items][1]
                # Borda count increment
                itemsRating[item] = itemsRating[item] + (prefList.shape[1] - items)
        preferedItems[i] = np.flip(np.argsort(itemsRating))[0]
    return preferedItems

# Copeland Method Algorithm for a group returning a recommended item
def copeland_method(groupIndexes, prefList):
    # Create the copeland matrix
    item_wins = [0] * prefList.shape[1]
    item_a, item_b = 0, 1

    while item_a < (prefList.shape[1] - 1) and item_b < (prefList.shape[1]):
        roundWins = [0] * 2
        for i in range(0, len(groupIndexes)):
            if prefList[groupIndexes[i]][item_a] > prefList[groupIndexes[i]][item_b]:
                roundWins[0] += 1
            elif prefList[groupIndexes[i]][item_a] == prefList[groupIndexes[i]][item_b]:
                pass
            else:
                roundWins[1] += 1

        if roundWins[0] > roundWins[1]:
            item_wins[item_a] += 1
        
        if roundWins[0] < roundWins[1]:
            item_wins[item_b] += 1
        else:
            item_wins[item_a] += 0.5
            item_wins[item_b] += 0.5

        item_b += 1
        if item_b == prefList.shape[1]-1:
            item_a = item_a + 1
            item_b = item_a + 1

    return item_wins.index(max(item_wins))

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
    calculate_average_score(winnersArray, r, groups, groupSize)


# Reweighed Approval Voting algorithm for a groups that returns k items
def reweighted_approval_voting(groupIndexes, prefList, k, threshold):
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
        weightedItemVotes = [0] * prefList.shape[1]
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