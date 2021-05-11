import gspread
from gspread.models import Cell
from gspread_formatting import *
from gspread.utils import (
    finditem,
    fill_gaps
)

from options import Options
from utils import Utils
import pandas as pd


class GspreadWrapper():
    def __init__(self):
        self.opt = Options()
        self.utils = Utils()
        self.gc = gspread.service_account(filename=self.opt.gSheetAuthFile)

        self.infringementsColumns = [
            self.opt.profanityCol, self.opt.nonConstructiveCol,
            self.opt.scoreCol, self.opt.copyCol, self.opt.incompleteReadingCol,
            self.opt.notRelatedCol, self.opt.otherCol
        ]

    def loadAssessmentsFile(self):
        self.assessmentsDoc = self.gc.open_by_key(self.opt.originalExportFromIdeascale)
        self.assessmentsSheet = self.assessmentsDoc.worksheet(self.opt.assessmentsSheet)
        self.df = False

    def prepareDataFromExport(self):
        if (self.assessmentsSheet):
            self.df = pd.DataFrame(self.assessmentsSheet.get_all_records())
            # Assign ids in new column
            self.df.insert(0, self.opt.assessmentsIdCol, self.df.index + 1)
            # Assign proposal ids
            self.df[self.opt.proposalIdCol] = self.df.groupby(
                self.opt.proposalKeyCol
            ).ngroup()
            # Assign assessor_id
            self.df['assessor_id'] = self.df[self.opt.assessorCol].str.replace('z_assessor_', '')
            # Assign triplet_id
            self.df[self.opt.tripletIdCol] = self.df['assessor_id'] + '-' + self.df[self.opt.proposalIdCol].astype(str)
            return self.df
        return False

    def getProposersData(self):
        self.proposersDoc = self.gc.open_by_key(self.opt.proposersFile)
        self.proposersSheet = self.proposersDoc.worksheet(self.opt.assessmentsSheet)
        self.dfProposers = pd.DataFrame(self.proposersSheet.get_all_records())
        self.dfProposers[self.opt.proposerMarkCol] = self.dfProposers.apply(self.checkMarks, axis=1)

    def checkMarks(self, row):
        res = False
        for col in self.infringementsColumns:
            if (row[col].strip() != ''):
                # If General check also the rationale
                if (col == self.opt.otherCol):
                    if (row[self.opt.otherRationaleCol].strip() != ''):
                        res = True
                else:
                    res = True
        if (row[self.opt.topQualityCol].strip() != ''):
            res = (not res)
        return res

    def createDoc(self, name):
        print('Create new document...')
        spreadsheet = self.gc.create(name)
        spreadsheet.share(
            self.opt.accountEmail,
            perm_type='user',
            role='writer'
        )

        #blank_worksheet = spreadsheet.get_worksheet(0)
        #spreadsheet.del_worksheet(blank_worksheet)

        return spreadsheet

    def writeDf(self, worksheet, df, headings=False):
        # Extract the columns already present in dataframe
        existingHeadings = list(df.columns)
        if headings == False:
            headings = existingHeadings
        # Create new Dataframe only with needed columns
        newDf = pd.DataFrame()
        for col in headings:
            if (col in existingHeadings):
                newDf[col] = df[col]
            else:
                newDf[col] = ""

        worksheet.update(
            [newDf.columns.values.tolist()] + newDf.values.tolist()
        )

    def createSheetFromDf(self, spreadsheet, sheetName, df, headings=False, columnWidths=False, formats=False):
        print('Create sheet...')
        worksheet = spreadsheet.add_worksheet(title=sheetName, rows=1, cols=1)
        self.writeDf(worksheet, df, headings)
        if (columnWidths is not False):
            set_column_widths(worksheet, columnWidths)
        if (formats is not False):
            format_cell_ranges(worksheet, formats)

        worksheet.freeze(rows=1)
