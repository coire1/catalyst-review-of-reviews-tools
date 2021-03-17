import gspread
from gspread.models import Cell
from gspread_formatting import *
from gspread.utils import (
    finditem,
    fill_gaps
)
from itertools import groupby

from options import Options
from utils import Utils


'''
    Monkey patch to get merged cells values.
'''

def fix_merge_values(worksheet_metadata, values):
    """Assign the top-left value to all cells in a merged range."""
    for merge in worksheet_metadata.get("merges", []):
        start_row, end_row = merge["startRowIndex"], merge["endRowIndex"]
        start_col, end_col = (merge["startColumnIndex"], merge["endColumnIndex"])

        # ignore merge cells outside the data range
        if start_row < len(values) and start_col < len(values[0]):
            orig_val = values[start_row][start_col]
            for row in values[start_row:end_row]:
                row[start_col:end_col] = [
                    orig_val for i in range(start_col, end_col)
                ]

    return values

def _get_all_values(self, value_render_option='FORMATTED_VALUE'):
    """Returns a list of lists containing all cells' values as strings.
    :param value_render_option: (optional) Determines how values should be
                                rendered in the the output. See
                                `ValueRenderOption`_ in the Sheets API.
    :type value_render_option: str
    .. _ValueRenderOption: https://developers.google.com/sheets/api/reference/rest/v4/ValueRenderOption
    .. note::
        Empty trailing rows and columns will not be included.
    """
    title = self.title.replace("'", "''")
    data = self.spreadsheet.values_get(
        "'{}'".format(title),
        params={'valueRenderOption': value_render_option}
    )
    spreadsheet_meta = self.spreadsheet.fetch_sheet_metadata()

    # not catching StopIteration becuase Worksheet exists
    # potential issue if this Worksheet
    worksheet_meta = finditem(
        lambda x: x['properties']['title'] == self.title,
        spreadsheet_meta['sheets']
    )

    try:
        values = fill_gaps(data['values'])
        return fix_merge_values(worksheet_meta, values)
    except KeyError:
        return []
# Apply Monkey patch
gspread.models.Worksheet.get_all_values = _get_all_values

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
        proposalsWithIds = self.generateProposalsId(self.proposersData)
        for (i, data) in enumerate(self.proposersData):
            # Add id if not present. 2 Offset because Google Sheet starts from 1
            # and the first row is the table heading.
            if (self.options.assessmentsIdColumn not in data):
                data[self.options.assessmentsIdColumn] = i + 2
            if (self.options.tripletIdColumn not in data):
                assessorId = data[self.options.assessorColumn].replace('z_assessor_', '')
                proposalK = data[self.options.proposalKeyColumn]
                proposalId = proposalsWithIds[proposalK]
                data[self.options.tripletIdColumn] = "{}-{}".format(assessorId, proposalId)
        return self.proposersData

    def getAssessmentsData(self):
        if (self.assessmentsSheet):
            cacheFn = self.options.assessmentsCacheFilename
            self.assessmentsData = self.utils.loadCache(cacheFn)
            if (self.assessmentsData is False):
                self.assessmentsData = self.assessmentsSheet.get_all_records()
                self.utils.saveCache(self.assessmentsData, cacheFn)
            # Loop over all records to assign a unique id
            proposalsWithIds = self.generateProposalsId(self.assessmentsData)
            for (i, data) in enumerate(self.assessmentsData):
                # Add id if not present. 2 Offset because Google Sheet starts from 1
                # and the first row is the table heading.
                if (self.options.assessmentsIdColumn not in data):
                    data[self.options.assessmentsIdColumn] = i + 2
                if (self.options.tripletIdColumn not in data):
                    assessorId = data[self.options.assessorColumn].replace('z_assessor_', '')
                    proposalK = data[self.options.proposalKeyColumn]
                    proposalId = proposalsWithIds[proposalK]
                    data[self.options.tripletIdColumn] = "{}-{}".format(assessorId, proposalId)
            return self.assessmentsData
        return False

    def countMarked(self, data, column):
        count = sum(map(lambda rec: rec[column] == 'x', data))
        return count

    def generateProposalsId(self, data):
        proposals = {}
        if (data):
            sort = sorted(data, key=lambda rec: rec[self.options.proposalKeyColumn])
            proposalId = 1
            for k, v in groupby(sort, lambda rec: rec[self.options.proposalKeyColumn]):
                proposals[k] = proposalId
                proposalId = proposalId + 1

        return proposals

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
                assessors[k]['blankPercentage'] = blankCount / len(assessments)
                assessors[k]['score'] = self.countMarked(assessments, self.options.scoreColumn)
                assessors[k]['copy'] = self.countMarked(assessments, self.options.copyColumn)
                assessors[k]['wrongChallenge'] = self.countMarked(assessments, self.options.wrongChallengeColumn)
                assessors[k]['wrongCriteria'] = self.countMarked(assessments, self.options.wrongCriteriaColumn)
                assessors[k]['other'] = self.countMarked(assessments, self.options.otherColumn)
        return assessors

    def groupById(self, data):
        ids = {}
        if (data):
            for row in data:
                ids[row[self.options.assessmentsIdColumn]] = row
        return ids

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


    def createSheetFromList(self, spreadsheet, title, data, columnsBlacklist):
        # data is expected to be a list of dict. A column for each dict key will
        # be created
        if (data):
            headings = list(data[0].keys())
            # filtering out blacklisted keys
            headings = [x for x in headings if (x not in columnsBlacklist)]
            worksheet = spreadsheet.add_worksheet(title=title, rows=100, cols=len(headings) + 1)
            cellsToAdd = []
            for i, value in enumerate(headings):
                cellsToAdd.append(
                    Cell(row=1, col=(i + 1), value=value)
                )

            # Add aggregated counts to cells
            prIndex = 2
            for el in data:
                for j, k in enumerate(el):
                    v = el[k]
                    cellsToAdd.append(
                        Cell(row=prIndex, col=(j + 1), value=v)
                    )
                prIndex = prIndex + 1

            worksheet.update_cells(cellsToAdd, value_input_option='USER_ENTERED')
