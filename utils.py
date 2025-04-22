import numpy as np
import pandas as pd
import xlsxwriter

from pathlib import Path
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
    WS = pd.read_excel(Path(__file__).parent / "Datasets/preferenceList.xlsx")
    r = np.array(WS)
    return r


# Import the itemsCost Matrix
def import_items_cost():
    WS = pd.read_excel(Path(__file__).parent / "Datasets/itemsCost.xlsx")
    itemsCost = np.array(WS)
    return itemsCost


# Import the usersBudget matrix
def import_users_budget():
    WS = pd.read_excel(Path(__file__).parent / "Datasets/usersBudget.xlsx")
    usersBudget = np.array(WS)
    return usersBudget
