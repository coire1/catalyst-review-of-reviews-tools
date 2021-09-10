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
            self.opt.proposalKeyCol, self.opt.ideaURLCol, self.opt.assessorCol,
            self.opt.tripletIdCol, self.opt.proposalIdCol,
            self.opt.q0Col, self.opt.q0Rating, self.opt.q1Col, self.opt.q1Rating,
            self.opt.q2Col, self.opt.q2Rating, self.opt.blankCol,
            self.opt.topQualityCol, self.opt.goodCol, self.opt.notValidCol,
            self.opt.otherRationaleCol
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
            ('A:B', 150), ('C', 100), ('D:E', 40), ('F', 300), ('G', 30), ('H', 300), ('I', 30),
            ('J', 300), ('K:O', 30), ('P', 300)
        ]

        formats = [
            ('G', self.utils.counterFormat),
            ('I', self.utils.counterFormat),
            ('K', self.utils.counterFormat),
            ('L', self.utils.counterFormat),
            ('A1:P1', self.utils.headingFormat),
            ('L1:O1', self.utils.verticalHeadingFormat),
            ('G1', self.utils.verticalHeadingFormat),
            ('I1', self.utils.verticalHeadingFormat),
            ('K1', self.utils.verticalHeadingFormat),
            ('M2:M', self.utils.greenFormat),
            ('N2:N', self.utils.greenFormat),
            ('O2:O', self.utils.redFormat),
            ('F2:F', self.utils.textFormat),
            ('H2:H', self.utils.textFormat),
            ('J2:J', self.utils.textFormat),
        ]

        self.gspreadWrapper.createSheetFromDf(
            spreadsheet,
            "Assessments",
            proposerDf,
            headings,
            widths,
            formats
        )
        print('Document for proposers created')
        print('Link: {}'.format(spreadsheet.url))

c = CreateProposerDocument()
c.createDoc()
