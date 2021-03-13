import gspread

from options import Options
from utils import Utils

class GspreadWrapper():
    def __init__(self):
        self.options = Options()
        self.utils = Utils()
        self.gc = gspread.service_account(filename=self.options.gSheetAuthFile)
        self.assessmentsData = {}

    def loadAssessmentsFile(self):
        self.assessmentsDoc = self.gc.open_by_key(self.options.originalExportFromIdeascale)
        self.assessmentsSheet = self.assessmentsDoc.worksheet(self.options.assessmentsSheet)

    def getProposersData(self):
        self.proposersDoc = self.gc.open_by_key(self.options.proposersFile)
        self.proposersSheet = self.proposersDoc.worksheet(self.options.assessmentsSheet)
        self.proposersData = self.proposersSheet.get_all_records()
        for (i, data) in enumerate(self.proposersData):
            # Add id if not present. 2 Offset because Google Sheet starts from 1
            # and the first row is the table heading.
            if (self.options.assessmentsIdColumn not in data):
                data[self.options.assessmentsIdColumn] = i + 2
        return self.proposersData

    def getAssessmentsData(self):
        if (self.assessmentsSheet):
            cacheFn = self.options.assessmentsCacheFilename
            self.assessmentsData = self.utils.loadCache(cacheFn)
            if (self.assessmentsData is False):
                self.assessmentsData = self.assessmentsSheet.get_all_records()
                self.utils.saveCache(self.assessmentsData, cacheFn)
            # Loop over all records to assign a unique id
            for (i, data) in enumerate(self.assessmentsData):
                # Add id if not present. 2 Offset because Google Sheet starts from 1
                # and the first row is the table heading.
                if (self.options.assessmentsIdColumn not in data):
                    data[self.options.assessmentsIdColumn] = i + 2
            return self.assessmentsData
        return False
