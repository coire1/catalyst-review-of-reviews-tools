from gspreadWrapper import GspreadWrapper
from options import Options
from utils import Utils

from gspread_formatting import *
import pandas as pd

class CreateProposerDocument():
    def __init__(self):
        self.opt = Options()
        self.utils = Utils()
        self.gspreadWrapper = GspreadWrapper()

    def createDoc(self):
        pd.options.display.max_columns = 100
        print('Loading original...')
        self.gspreadWrapper.loadAssessmentsFile()
        proposerDf = self.gspreadWrapper.prepareDataFromExport()
        spreadsheet = self.gspreadWrapper.createDoc(
            self.opt.proposerDocumentName
        )
        # Define all the columns needed in the file
        headings = [
            self.opt.assessmentsIdCol,
            self.opt.proposalKeyCol, self.opt.ideaURLCol, self.opt.assessorCol,
            self.opt.tripletIdCol, self.opt.proposalIdCol,
            self.opt.q0Col, self.opt.q0Rating, self.opt.q1Col, self.opt.q1Rating,
            self.opt.q2Col, self.opt.q2Rating, self.opt.blankCol,
            self.opt.notValidCol, self.opt.notValidRationaleCol
        ]

        print('Assign blanks...')
        # Assign 'x' for blank assessments
        proposerDf[self.opt.blankCol] = proposerDf.apply(
            lambda r: 'x' if (
                (str(r[self.opt.q0Col]).strip() == "") or
                (str(r[self.opt.q1Col]).strip() == "") or
                (str(r[self.opt.q2Col]).strip() == "")
            ) else ''
        , axis=1)

        print('Format columns...')
        widths = [
            ('A', 30), ('B:C', 150), ('D', 100), ('E:F', 40), ('G', 300),
            ('H', 30), ('I', 300), ('J', 30), ('K', 300), ('L:N', 30),
            ('O', 300)
        ]

        formats = [
            ('A', self.utils.counterFormat),
            ('H', self.utils.counterFormat),
            ('J', self.utils.counterFormat),
            ('L', self.utils.counterFormat),
            ('M', self.utils.counterFormat),
            ('A1:O1', self.utils.headingFormat),
            ('M1:N1', self.utils.verticalHeadingFormat),
            ('H1', self.utils.verticalHeadingFormat),
            ('J1', self.utils.verticalHeadingFormat),
            ('L1', self.utils.verticalHeadingFormat),
            ('N2:N', self.utils.redFormat),
            ('G2:G', self.utils.textFormat),
            ('I2:I', self.utils.textFormat),
            ('K2:K', self.utils.textFormat),
        ]

        self.gspreadWrapper.createSheetFromDf(
            spreadsheet,
            "Assessments",
            proposerDf,
            headings,
            widths,
            formats
        )
        print('Master Document for proposers created')
        print('Link: {}'.format(spreadsheet.url))

c = CreateProposerDocument()
c.createDoc()
