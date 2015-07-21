"""
Workorder class and utility functions
"""
__author__ = 'mmcclure'

# TODO: fix the leaky abstracton around the loading of the data
class Workorder:
    '''
    Class that hides all the complexity and data of the work order
    Abstracts out the loading interface to the work order data too, so we can change sources with minimal disruption
    Currently takes a list, raw out of the WODataFeed.xlsm and builds the object from that
    '''
    def __init__(self, raw):
        self.raw = raw
        self.wo = int(raw[0].value)
        self.wo_link = 'http://pons/wo/Pages/WorkOrder.aspx?WO=' + str(self.wo)
        self.name = raw[84].value
        self.client = raw[83].value
        self.type = raw[20].value
        self.PM = raw[97].value
        self.stage = raw[68].value
        self.is_active = bool(raw[100].value)
        # values related to the estimate
        self.est = {}
        self.est['Complexity'] = raw[99].value
        self.est['Total Hours'] = float(raw[6].value)
        # values related to the actuals from FF
        self.act = {}
        self.act['Total Hours'] = float(raw[13].value)
        self.act['Estimate Complete'] = cell_to_date(raw[23].value)
        self.act['WO Signed'] = cell_to_date(raw[26].value)
        self.act['Client KO'] = cell_to_date(raw[28].value)
        self.act['Dev KO'] = cell_to_date(raw[30].value)
        self.act['QA KO'] = cell_to_date(raw[31].value)
        self.act['QA Complete'] = cell_to_date(raw[32].value)
        self.act['Go Live'] = cell_to_date(raw[33].value)
        self.act['VCS Handoff'] = cell_to_date(raw[34].value)

        #### calculated/derived fields
        #calculate what stage the project is at based on dates that are filled in
        if isinstance(self.act['VCS Handoff'], datetime.date):
            self.status = "BAU"
        elif isinstance(self.act['Go Live'], datetime.date):
            self.status = "Live"
        elif isinstance(self.act['QA Complete'], datetime.date):
            self.status = "QA Complete"
        elif isinstance(self.act['QA KO'], datetime.date):
            self.status = "QA Kickoff"
        elif isinstance(self.act['Dev KO'], datetime.date):
            self.status = "Dev Kickoff"
        elif isinstance(self.act['Client KO'], datetime.date):
            self.status = "Client Kickoff"
        elif isinstance(self.act['WO Signed'], datetime.date):
            self.status = "Signed"
        elif isinstance(self.act['Estimate Complete'], datetime.date):
            self.status = "Estimated"
        else:
            self.status = "UNKNOWN"

    def get_dashboard_header(self):
        '''
        Build the header row for the dashboard
        :return: list of row headers
        '''
        headers = []
        headers.append('WO #')
        headers.append('Client')
        headers.append('WO Name')
        headers.append('PM')
        headers.append('Project Type')
        headers.append('Estimated Hours')
        headers.append('Actual Hours')
        headers.append('Estimated Date')
        headers.append('Signed Date')
        headers.append('Client Kickoff')
        headers.append('Dev Kickoff')
        headers.append('QA Kickoff')
        headers.append('QA Complete')
        headers.append('Go Live Date')
        headers.append('VCS Handoff Date')
        headers.append('WO Status')
        headers.append('Project Stage')
        return headers

    def get_dashboard_content(self):
        '''
        Build the WO details (content row) for the dashboard
        :return: list of row items
        '''
        content = []
        content.append(self.wo)
        content.append(self.client)
        content.append(self.name)
        content.append(self.PM)
        content.append(self.type)
        content.append(self.est['Total Hours'])
        content.append(self.act['Total Hours'])
        content.append(self.act['Estimate Complete'])
        content.append(self.act['WO Signed'])
        content.append(self.act['Client KO'])
        content.append(self.act['Dev KO'])
        content.append(self.act['QA KO'])
        content.append(self.act['QA Complete'])
        content.append(self.act['Go Live'])
        content.append(self.act['VCS Handoff'])
        content.append(self.status)
        content.append(self.stage)
        return content