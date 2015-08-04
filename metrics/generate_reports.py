"""
This is where all the reports are generated

1/ Project Status Dashboard (AKA The EVM Dashboard) - Unfinished
2/ Open projects - Unfinished
3/ Estimates In / Estimates Out (AKA The Flow Report) - Unfinished
"""

__author__ = 'mmcclure'


import openpyxl
from load_data import load_from_spreadsheet
import metrics.config
datafeed = metrics.config.datafeed


def create_dashboard(wlist):
    """
    Create an excel dashboard from a list of objects
    :param wlist: list of Workorder objects to be processed into the dashboard
    """
    wb = openpyxl.Workbook()
    page = wb.active
    page.title = 'Dashboard'
    # TODO add in other tabs per project
    # TODO change WO # to be a URL
    # TODO formatting
    # TODO names ranges, so this can just feed in the raw data, and charts & pivots can be done off that?

    # populate dashboard tab with a header
    wb['Dashboard'].append(wlist[0].get_dashboard_header())
    # populate dashboard tab with each WO
    counter = 2 #offset to find the first cell with a WO in it "A2"
    for row in wlist:
        # add the meat
        wb['Dashboard'].append(row.get_dashboard_content())
        # convert the WO number into a link
        wb['Dashboard'].cell('A' + str(counter)).hyperlink = row.wo_link
        counter += 1

    # add in autofilter, automatically determining the last column to use
    wb['Dashboard'].auto_filter.ref = 'A1:' + chr(ord('A') -1 + len(wlist[0].get_dashboard_header())) + '1'
    print wb['Dashboard'].auto_filter.ref
    wb['Dashboard'].auto_filter.add_sort_condition('D1:D1')

    # commit sheet to file, beware - this silently overwrites!
    wb.save('EVM Dashboard.xlsx')


def build_flow_report(wlist):
    """
    Create an excel report from a list of objects:
        Flow Report - showing the estimates completed, and the projects completed, per month
        Designed to demonstrate ????
    :param wlist: list of Workorder objects to be processed into the report
    """
    print "starting flow report, unfiltered list contains %d work orders" % len(wlist)
    tmp_list = []
    print "finished flow report, final list contains %d work orders" % len(tmp_list)


if __name__ == '__main__':
    """if this module is run, assume that the user just wants to load the data from a file
    """
    print 'loading the WODataFeed...'
    print 'Found %d records' % len(wolist)

    # load the list of work orders
    wolist = load_from_spreadsheet(metrics.config.test_datafeed)

    # build the flow report
    build_flow_report(wolist)
