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
            self.opt.assessmentsIdCol, self.opt.tripletIdCol, self.opt.ideaURLCol,
            self.opt.proposalIdCol, self.opt.questionCol, self.opt.questionIdCol,
            self.opt.ratingCol, self.opt.assessorCol, self.opt.assessmentCol,
            self.opt.proposerMarkCol, self.opt.fairCol, self.opt.topQualityCol,
            self.opt.abstainCol, self.opt.strictCol, self.opt.lenientCol,
            self.opt.profanityCol, self.opt.nonConstructiveCol,
            self.opt.scoreCol, self.opt.copyCol, self.opt.incompleteReadingCol,
            self.opt.notRelatedCol, self.opt.otherCol, self.opt.otherRationaleCol
        ]

        print('Load proposers flagged reviews...')
        self.gspreadWrapper.getProposersData()
        #self.gspreadWrapper.dfProposers.to_csv('test.csv')

        # Extract assessors
        assessors = self.gspreadWrapper.dfProposers.groupby(
            self.opt.assessorCol
        ).agg(
            total=(self.opt.assessmentCol, 'count'),
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
        criteria = self.gspreadWrapper.infringementsColumns + [self.opt.topQualityCol, self.opt.otherRationaleCol]
        for col in criteria:
            validAssessments[col] = ''

        # Assign 'x' for marks
        validAssessments[self.opt.proposerMarkCol] = validAssessments[self.opt.proposerMarkCol].apply(
            lambda r: 'x' if (r) else ''
        )

        # Write sheet with assessments
        assessmentsWidths = [
            ('A', 40), ('B', 60), ('D', 40), ('E', 200), ('F', 40), ('G', 60), ('H', 120), ('I', 400),
            ('J:V', 30), ('W', 300)
        ]
        assessmentsFormats = [
            ('G:G', self.utils.counterFormat),
            ('I:I', self.utils.noteFormat),
            ('J:V', self.utils.counterFormat),
            ('A1:W1', self.utils.headingFormat),
            ('B1', self.utils.verticalHeadingFormat),
            ('D1', self.utils.verticalHeadingFormat),
            ('F1:G1', self.utils.verticalHeadingFormat),
            ('J1:V1', self.utils.verticalHeadingFormat),
            ('K2:L', self.utils.greenFormat),
            ('P2:P', self.utils.redFormat),
            ('Q2:V', self.utils.yellowFormat),
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
