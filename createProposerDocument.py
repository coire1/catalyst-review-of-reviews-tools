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
        print('Loading original...')
        self.gspreadWrapper.loadAssessmentsFile()
        proposerDf = self.gspreadWrapper.prepareDataFromExport()
        spreadsheet = self.gspreadWrapper.createDoc(
            self.opt.proposerDocumentName
        )

        # Define all the columns needed in the file
        headings = [
            self.opt.proposalKeyCol, self.opt.ideaURLCol, self.opt.questionCol,
            self.opt.assessorCol, self.opt.assessmentCol, self.opt.ratingCol,
            self.opt.assessmentsIdCol, self.opt.tripletIdCol, self.opt.proposalIdCol,
            self.opt.blankCol, self.opt.topQualityCol, self.opt.profanityCol,
            self.opt.nonConstructiveCol, self.opt.scoreCol, self.opt.copyCol,
            self.opt.incompleteReadingCol, self.opt.notRelatedCol,
            self.opt.otherCol, self.opt.otherRationaleCol
        ]

        print('Assign blanks...')
        # Assign 'x' for blank assessments
        proposerDf[self.opt.blankCol] = proposerDf[self.opt.assessmentCol].apply(
            lambda r: 'x' if (r.strip() == "") else ''
        )

        print('Format columns...')
        widths = [
            ('A:D', 150), ('E', 400),
            ('F', 60),('G:R', 30), ('S', 400)
        ]

        formats = [
            ('F:R', self.utils.counterFormat),
            ('A1:S1', self.utils.headingFormat),
            ('F1:R1', self.utils.verticalHeadingFormat),
            ('K2:K', self.utils.greenFormat),
            ('L2:L', self.utils.redFormat),
            ('M2:R', self.utils.yellowFormat),
            ('A2:E', self.utils.textFormat),
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
