import gspread
from gspread.models import Cell
from gspread_formatting import *
from itertools import groupby

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

    def countMarked(self, data, column):
        count = sum(map(lambda rec: rec[column] == 'x', data))
        return count

    def groupByAssessor(self, data):
        assessors = {}
        if (data):
            sort = sorted(data, key=lambda rec: rec[self.options.assessorColumn])
            for k, assessments in groupby(sort, lambda rec: rec[self.options.assessorColumn]):
                assessments = list(assessments)
                blankCount = self.countMarked(assessments, self.options.blankColumn)
                assessors[k] = {}
                assessors[k]['assessments'] = assessments
                assessors[k]['total'] = len(assessments)
                assessors[k]['profanity'] = self.countMarked(assessments, self.options.profanityColumn)
                assessors[k]['blank'] = blankCount
                assessors[k]['blankPercentage'] = (100 * blankCount) / len(assessments)
                assessors[k]['score'] = self.countMarked(assessments, self.options.scoreColumn)
                assessors[k]['copy'] = self.countMarked(assessments, self.options.copyColumn)
                assessors[k]['wrongChallenge'] = self.countMarked(assessments, self.options.wrongChallengeColumn)
                assessors[k]['wrongCriteria'] = self.countMarked(assessments, self.options.wrongCriteriaColumn)
                assessors[k]['other'] = self.countMarked(assessments, self.options.otherColumn)
        return assessors

    def createSheetFromGroup(self, spreadsheet, title, data, keysWhitelist, columnsBlacklist):
        # data is expected to be a dict of dict. A column for each dict key will
        # be created
        if (data):
            headings = list(list(data.items())[0][1].keys())
            # filtering out blacklisted keys
            headings = [x for x in headings if (x not in columnsBlacklist)]
            worksheet = spreadsheet.add_worksheet(title=title, rows=100, cols=len(headings) + 1)
            cellsToAdd = []

            set_column_widths(worksheet, [
                ('A', 110), ('B:J', 40)
            ])

            for i, value in enumerate(headings):
                cellsToAdd.append(
                    Cell(row=1, col=(i + 2), value=value)
                )

            prIndex = 2
            for index, mainKey in enumerate(data):
                if (mainKey in keysWhitelist):
                    cellsToAdd.append(
                        Cell(row=prIndex, col=1, value=mainKey)
                    )
                    for i, k in enumerate(data[mainKey]):
                        if (k not in columnsBlacklist):
                            value = data[mainKey][k]
                            cellsToAdd.append(
                                Cell(row=prIndex, col=(i + 1), value=value)
                            )
                    prIndex = prIndex + 1
            worksheet.update_cells(cellsToAdd, value_input_option='USER_ENTERED')
