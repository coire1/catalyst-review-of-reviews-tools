from gspreadWrapper import GspreadWrapper
from options import Options
from utils import Utils

from gspread.models import Cell
from gspread_formatting import *

class CreateVCAMaster():
    def __init__(self):
        self.opt = Options()
        self.utils = Utils()
        self.gspreadWrapper = GspreadWrapper()

    def createDoc(self):
        spreadsheet = self.gspreadWrapper.createDoc(self.opt.VCAMasterFileName)

        # Define headings for VCAMasterFile
        print('Define headings...')
        headings = [
            self.opt.assessmentsIdCol,
            self.opt.proposalKeyCol, self.opt.ideaURLCol, self.opt.assessorCol,
            self.opt.tripletIdCol, self.opt.proposalIdCol,
            self.opt.q0Col, self.opt.q0Rating, self.opt.q1Col, self.opt.q1Rating,
            self.opt.q2Col, self.opt.q2Rating, self.opt.proposerMarkCol,
            self.opt.proposersRationaleCol, self.opt.excellentCol,
            self.opt.goodCol, self.opt.notValidCol
        ]

        print('Load proposers flagged reviews...')
        self.gspreadWrapper.getProposersAggregatedData()
        #self.gspreadWrapper.dfProposers.to_csv('test.csv')

        # Extract assessors
        assessors = self.gspreadWrapper.dfProposers.groupby(
            self.opt.assessorCol
        ).agg(
            total=(self.opt.tripletIdCol, 'count'),
            blanks=(self.opt.blankCol, (lambda x: (x == 'x').sum()))
        )

        # Calculate and extract assessors by blanks
        assessors['blankPercentage'] = assessors['blanks'] / assessors['total']
        assessors['excluded'] = (assessors['blankPercentage'] >= self.opt.allowedBlankPerAssessor)
        excludedAssessors = assessors[(assessors['excluded'] == True)].index.tolist()
        includedAssessors = assessors[(assessors['excluded'] != True)].index.tolist()

        # Exclude assessors that are also proposers (get from options)
        includedAssessors = [x for x in includedAssessors if (x not in self.opt.excludedCAProposers)]
        excludedAssessors.extend(self.opt.excludedCAProposers)

        assessors['assessor'] = assessors.index

        # Filter out assessments made by excluded assessors
        validAssessments = self.gspreadWrapper.dfProposers[
            self.gspreadWrapper.dfProposers[self.opt.assessorCol].isin(includedAssessors)
        ]

        # Filter out blank assessments
        validAssessments = validAssessments[validAssessments[self.opt.blankCol] != 'x']

        # Remove proposers marks
        criteria = [self.opt.excellentCol, self.opt.goodCol, self.opt.notValidCol]
        for col in criteria:
            validAssessments[col] = ''

        # Assign 'x' for marks
        validAssessments[self.opt.proposerMarkCol] = validAssessments[self.opt.proposerMarkCol].apply(
            lambda r: 'x' if (r > 0) else ''
        )

        # Write sheet with assessments
        assessmentsWidths = [
            ('A', 30), ('B:C', 150), ('D', 100), ('E:F', 40), ('G', 300),
            ('H', 30), ('I', 300), ('J', 30), ('K', 300), ('L:N', 30),
            ('N', 300), ('O:Q', 30)
        ]
        assessmentsFormats = [
            ('A', self.utils.counterFormat),
            ('H', self.utils.counterFormat),
            ('J', self.utils.counterFormat),
            ('L', self.utils.counterFormat),
            ('M', self.utils.counterFormat),
            ('A1:Q1', self.utils.headingFormat),
            ('M1:M1', self.utils.verticalHeadingFormat),
            ('H1', self.utils.verticalHeadingFormat),
            ('J1', self.utils.verticalHeadingFormat),
            ('L1', self.utils.verticalHeadingFormat),
            ('O1:Q1', self.utils.verticalHeadingFormat),
            ('G2:G', self.utils.textFormat),
            ('I2:I', self.utils.textFormat),
            ('K2:K', self.utils.textFormat),
            ('O2:O', self.utils.greenFormat),
            ('P2:P', self.utils.greenFormat),
            ('Q2:Q', self.utils.yellowFormat),
        ]

        self.gspreadWrapper.createSheetFromDf(
            spreadsheet,
            'Assessments',
            validAssessments,
            headings,
            columnWidths=assessmentsWidths,
            formats=assessmentsFormats
        )

        # Write sheet with CAs summary
        self.gspreadWrapper.createSheetFromDf(
            spreadsheet,
            'Community Advisors',
            assessors,
            ['assessor', 'total', 'blanks', 'blankPercentage', 'excluded'],
            columnWidths=[('A', 140), ('B:D', 60), ('E', 100)],
            formats=[
                ('B:C', self.utils.counterFormat),
                ('D2:D', self.utils.percentageFormat),
                ('A1:E1', self.utils.headingFormat),
                ('B1:D1', self.utils.verticalHeadingFormat),
            ]
        )
        print('Master Document for vCAs created')
        print('Link: {}'.format(spreadsheet.url))

cvca = CreateVCAMaster()
cvca.createDoc()
