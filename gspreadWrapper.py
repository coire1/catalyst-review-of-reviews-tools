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

    def createDoc(self, name, sheetName):
        print('Create new document...')
        spreadsheet = self.gc.create(name)
        spreadsheet.share(
            self.opt.accountEmail,
            perm_type='user',
            role='writer'
        )

        print('Create sheet...')
        worksheet = spreadsheet.get_worksheet(0)
        worksheet.update_title(sheetName)

        return spreadsheet, worksheet

    def writeDf(self, worksheet, df):
        worksheet.update(
            [df.columns.values.tolist()] + df.values.tolist()
        )
