"""Loads all the data from the datasource into a list of Workorders
Input methods:
1/ WODataFeed.xlsm - rich excel sheet to load and aggregate
2/ Assorted spreadsheets (FUTURE) - the raw data sources used to populate WODataFeed, so that can be done away with
3/ Database/Data Store (FUTURE) - processed schema ODS style to reduce all the mucking about with teh raw sources
"""
__author__ = 'mmcclure'


import xlrd
from workorder import Workorder
import metrics.config
datafeed = metrics.config.datafeed


def load_from_spreadsheet(datafeed):
    """Load the work order data from a spreadsheet in the WODataFeed format
    :param datafeed: WODataFeed format file
    :return: list of Workorder objects
    """
    wb = xlrd.open_workbook(datafeed)
    # return the summary sheet
    sheet = wb.sheet_by_name('WO Rollup Info')

    # process each row (work order) in the sheet
    wolist = []
    for x in range(1, sheet.nrows):
        # build the list of objects to return
        wolist.append(Workorder(sheet.row(x)))
    return wolist


# TODO: add in a set of default filtered lists of Workorders, such as "Open", "Estimated", "Completed", etc. this way people use the same data sets more readily


if __name__ == '__main__':
    """if this module is run, assume that the user just wants to load the data from a file
    """
    print 'loading the WODataFeed...'
    wolist = load_from_spreadsheet(metrics.config.test_datafeed)
    print 'Found %d records' % len(wolist)


