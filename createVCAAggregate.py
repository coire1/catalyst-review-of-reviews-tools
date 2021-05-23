from gspreadWrapper import GspreadWrapper
from options import Options
from utils import Utils

from gspread.models import Cell
from gspread_formatting import *

from time import sleep
import pandas as pd

class createVCAAggregate():
    def __init__(self):
        self.opt = Options()
        self.utils = Utils()
        self.gspreadWrapper = GspreadWrapper()

        self.infringementsColumns = [
            self.opt.profanityCol, self.opt.nonConstructiveCol,
            self.opt.scoreCol, self.opt.copyCol, self.opt.incompleteReadingCol,
            self.opt.notRelatedCol, self.opt.otherCol
        ]
        self.positiveColumns = [self.opt.fairCol, self.opt.topQualityCol]
        '''
        self.neutralColumns = [
            self.opt.abstainCol, self.opt.lenientCol, self.opt.strictCol
        ]
        '''
        self.neutralColumns = [self.opt.abstainCol]
        self.indicatorColumns = [self.opt.lenientCol, self.opt.strictCol]
        self.yellowColumns = [
            self.opt.nonConstructiveCol, self.opt.scoreCol, self.opt.copyCol,
            self.opt.incompleteReadingCol, self.opt.notRelatedCol,
            self.opt.otherCol
        ]
        self.redColumns = [self.opt.profanityCol]
        self.feedbackColumns = self.infringementsColumns + self.positiveColumns
        self.allColumns = self.infringementsColumns + self.positiveColumns + self.neutralColumns + self.indicatorColumns

    def prepareBaseData(self):
        self.gspreadWrapper.getVCAMasterData()
        self.dfVca = self.gspreadWrapper.dfVca.set_index('id')
        # Set all counters to 0
        self.dfVca[self.opt.noVCAReviewsCol] = 0
        self.dfVca[self.opt.yellowCardCol] = 0
        self.dfVca[self.opt.redCardCol] = 0
        for col in self.allColumns:
            self.dfVca[col] = 0

        self.gspreadWrapper.getVCAMasterAssessors()
        self.gspreadWrapper.dfVcaAssessors[self.opt.yellowCardCol] = 0
        self.gspreadWrapper.dfVcaAssessors[self.opt.redCardCol] = 0

    def loadVCAsFiles(self):
        self.prepareBaseData()
        self.vcaData = []
        self.vcaDocs = []
        for vcaFile in self.opt.VCAsFiles:
            print(vcaFile)
            vcaDocument = self.gspreadWrapper.gc.open_by_key(vcaFile)
            vcaSheet = vcaDocument.worksheet("Assessments")
            data = pd.DataFrame(vcaSheet.get_all_records())
            data.set_index('id', inplace=True)
            self.vcaData.append(data)
            self.vcaDocs.append(vcaDocument)
            #sleep(35)

    def createDoc(self):
        self.loadVCAsFiles()
        # Loop over master ids as reference
        for id, row in self.dfVca.iterrows():
            # Loop over all vca files
            for vcaDf in self.vcaData:
                if (id in vcaDf.index):
                    locAss = vcaDf.loc[id]
                    integrity = self.checkIntegrity(id, row, locAss)
                    if (integrity is False):
                        print('Error')
                        break
                        break

                    good = self.goodFeedback(locAss)
                    bad = self.badFeedback(locAss)
                    neutral = self.neutralFeedback(locAss)
                    if (self.isVCAfeedbackValid(locAss, good, bad, neutral)):
                        if (good or bad):
                            self.dfVca.loc[id, self.opt.noVCAReviewsCol] = self.dfVca.loc[id, self.opt.noVCAReviewsCol] + 1
                        for col in self.allColumns:
                            colVal = self.checkIfMarked(locAss, col)
                            if (colVal > 0):
                                self.dfVca.loc[id, col] = self.dfVca.loc[id, col] + colVal

            (yellow, red) = self.calculateCards(self.dfVca.loc[id])
            self.dfVca.loc[id, self.opt.yellowCardCol] = yellow
            self.dfVca.loc[id, self.opt.redCardCol] = red

        # Extract red card assessors and update List
        redCards = self.dfVca[self.dfVca[self.opt.redCardCol] > 0]
        self.redCardsAssessors = list(redCards[self.opt.assessorCol].unique())


        # Select valid assessments (no red card assessors, no yellow card assessments, no blank assessments)
        validAssessments = self.dfVca[(
            (self.dfVca[self.opt.yellowCardCol] == 0) &
            ~self.dfVca[self.opt.assessorCol].isin(self.redCardsAssessors)
        )]
        validAssessments[self.opt.assessmentsIdCol] = validAssessments.index
        # generate Assessor Recap
        assessors = self.assessorRecap()
        # generate nonValidAssessment recap
        nonValidAssessments = self.nonValidAssessments(validAssessments)

        # Generate Doc
        validAssessments.to_csv('cache/valid.csv')
        nonValidAssessments.to_csv('cache/non-valid.csv')
        assessors.to_csv('cache/assessors.csv')
        spreadsheet = self.gspreadWrapper.createDoc(self.opt.VCAAggregateFileName)

        # Print valid assessments
        assessmentsHeadings = [
            self.opt.assessmentsIdCol, self.opt.tripletIdCol, self.opt.ideaURLCol,
            self.opt.proposalIdCol, self.opt.questionCol, self.opt.questionIdCol,
            self.opt.ratingCol, self.opt.assessorCol, self.opt.assessmentCol,
            self.opt.proposerMarkCol, self.opt.fairCol, self.opt.topQualityCol,
            self.opt.abstainCol, self.opt.strictCol, self.opt.lenientCol,
            self.opt.profanityCol, self.opt.nonConstructiveCol,
            self.opt.scoreCol, self.opt.copyCol, self.opt.incompleteReadingCol,
            self.opt.notRelatedCol, self.opt.otherCol, self.opt.noVCAReviewsCol,
            self.opt.yellowCardCol, self.opt.redCardCol
        ]
        assessmentsWidths = [
            ('A', 40), ('B', 60), ('C', 120), ('D', 40), ('E', 200), ('F', 40), ('G', 60), ('H', 120), ('I', 400),
            ('J:Y', 30)
        ]
        assessmentsFormats = [
            ('G:G', self.utils.counterFormat),
            ('I:I', self.utils.noteFormat),
            ('J:Y', self.utils.counterFormat),
            ('A1:W1', self.utils.headingFormat),
            ('B1', self.utils.verticalHeadingFormat),
            ('D1', self.utils.verticalHeadingFormat),
            ('F1:G1', self.utils.verticalHeadingFormat),
            ('J1:Y1', self.utils.verticalHeadingFormat),
            ('K2:L', self.utils.greenFormat),
            ('P2:P', self.utils.redFormat),
            ('Q2:V', self.utils.yellowFormat),
            ('X2:X', self.utils.yellowFormat),
            ('Y2:Y', self.utils.redFormat),
        ]

        self.gspreadWrapper.createSheetFromDf(
            spreadsheet,
            'Valid Assessments',
            validAssessments,
            assessmentsHeadings,
            columnWidths=assessmentsWidths,
            formats=assessmentsFormats
        )

        # Print assessors recap

        # Write sheet with CAs summary
        self.gspreadWrapper.createSheetFromDf(
            spreadsheet,
            'Community Advisors',
            assessors,
            [
                'assessor', 'total', 'blanks', 'blankPercentage', 'excluded',
                'Yellow Card', 'Red Card', 'Constructive Feedback', 'note'
            ],
            columnWidths=[('A', 140), ('B:D', 60), ('E', 100), ('F:H', 40), ('I', 200)],
            formats=[
                ('B:C', self.utils.counterFormat),
                ('F:H', self.utils.counterFormat),
                ('D2:D', self.utils.percentageFormat),
                ('A1:E1', self.utils.headingFormat),
                ('B1:D1', self.utils.verticalHeadingFormat),
                ('F1:H1', self.utils.verticalHeadingFormat),
                ('F2:F', self.utils.yellowFormat),
                ('G2:G', self.utils.redFormat),
                ('H2:H', self.utils.greenFormat),
            ]
        )


        # Print non valid assessments -> reason, extract from proposer Doc

        nonAssessmentsHeadings = [
            self.opt.assessmentsIdCol, self.opt.tripletIdCol, self.opt.ideaURLCol,
            self.opt.proposalIdCol, self.opt.questionCol,
            self.opt.ratingCol, self.opt.assessorCol, self.opt.assessmentCol,
            self.opt.blankCol,
            self.opt.proposerMarkCol, self.opt.fairCol, self.opt.topQualityCol,
            self.opt.abstainCol, self.opt.strictCol, self.opt.lenientCol,
            self.opt.profanityCol, self.opt.nonConstructiveCol,
            self.opt.scoreCol, self.opt.copyCol, self.opt.incompleteReadingCol,
            self.opt.notRelatedCol, self.opt.otherCol, self.opt.noVCAReviewsCol,
            self.opt.yellowCardCol, self.opt.redCardCol, 'reason'
        ]
        nonAssessmentsWidths = [
            ('A', 40), ('B', 60), ('C', 120), ('D', 40), ('E', 200), ('F', 40), ('G', 120), ('H', 400),
            ('I:Y', 30), ('Z', 200)
        ]
        nonAssessmentsFormats = [
            ('G:G', self.utils.counterFormat),
            ('H:H', self.utils.noteFormat),
            ('I:Y', self.utils.counterFormat),
            ('A1:W1', self.utils.headingFormat),
            ('B1', self.utils.verticalHeadingFormat),
            ('D1', self.utils.verticalHeadingFormat),
            ('F1:G1', self.utils.verticalHeadingFormat),
            ('I1:Y1', self.utils.verticalHeadingFormat),
            ('K2:L', self.utils.greenFormat),
            ('P2:P', self.utils.redFormat),
            ('Q2:V', self.utils.yellowFormat),
            ('X2:X', self.utils.yellowFormat),
            ('Y2:Y', self.utils.redFormat),
        ]

        self.gspreadWrapper.createSheetFromDf(
            spreadsheet,
            'Excluded Assessments',
            nonValidAssessments,
            nonAssessmentsHeadings,
            columnWidths=nonAssessmentsWidths,
            formats=nonAssessmentsFormats
        )

        # Print vca recap

        print('Aggregated Document created')
        print('Link: {}'.format(spreadsheet.url))

    def nonValidAssessments(self, validAssessments):
        self.gspreadWrapper.getProposersData()
        dfProposers = self.gspreadWrapper.dfProposers.set_index('id')
        dfProposers[self.opt.assessmentsIdCol] = dfProposers.index
        for col in self.allColumns:
            dfProposers[col] = ''
        dfProposers[self.opt.proposerMarkCol] = ''
        dfProposers[self.opt.otherRationaleCol] = ''
        nonValidAssessments = dfProposers[~dfProposers[self.opt.assessmentsIdCol].isin(validAssessments.index)].copy()
        for id, row in nonValidAssessments.iterrows():
            if (id in self.dfVca.index):
                for col in self.allColumns:
                    nonValidAssessments.loc[id, col] = int(self.dfVca.loc[id, col])
                nonValidAssessments.loc[id, self.opt.yellowCardCol] = int(self.dfVca.loc[id, self.opt.yellowCardCol])
                nonValidAssessments.loc[id, self.opt.redCardCol] = int(self.dfVca.loc[id, self.opt.redCardCol])
                nonValidAssessments.loc[id, self.opt.noVCAReviewsCol] = int(self.dfVca.loc[id, self.opt.noVCAReviewsCol])
        nonValidAssessments['reason'] = nonValidAssessments.apply(self.describeReason, axis=1)
        nonValidAssessments.fillna('', inplace=True)
        return nonValidAssessments

    def describeReason(self, row):
        reason = []
        if (row[self.opt.blankCol] == 'x'):
            reason.append(self.opt.blankCol)
        if (row['id'] in self.dfVca.index):
            tot = row[self.opt.noVCAReviewsCol]
            for col in self.infringementsColumns:
                if (tot > 0):
                    if ((row[col]/tot) >= self.opt.cardLimit):
                        reason.append(col)
        excludedAssessors = list(self.gspreadWrapper.dfVcaAssessors[self.gspreadWrapper.dfVcaAssessors['excluded'] == 'TRUE']['assessor'])
        if (row[self.opt.assessorCol] in excludedAssessors):
            reason.append('Assessor excluded')
        return ', '.join(reason)

    def assessorRecap(self):
        self.gspreadWrapper.dfVcaAssessors.loc[self.gspreadWrapper.dfVcaAssessors['assessor'].isin(self.redCardsAssessors), 'excluded'] = 'TRUE'
        self.gspreadWrapper.dfVcaAssessors.loc[self.gspreadWrapper.dfVcaAssessors['assessor'].isin(self.redCardsAssessors), 'note'] = "red card"

        # Extract assessors
        locAssessors = self.dfVca.groupby(
            self.opt.assessorCol
        ).agg(
            constructiveFeedback=(self.opt.topQualityCol, 'sum'),
            red=(self.opt.redCardCol, 'sum'),
            yellow=(self.opt.yellowCardCol, 'sum'),
        )
        for id, row in locAssessors.iterrows():
            self.gspreadWrapper.dfVcaAssessors.loc[self.gspreadWrapper.dfVcaAssessors['assessor'] == id, self.opt.redCardCol] = row['red']
            self.gspreadWrapper.dfVcaAssessors.loc[self.gspreadWrapper.dfVcaAssessors['assessor'] == id, self.opt.yellowCardCol] = row['yellow']
            self.gspreadWrapper.dfVcaAssessors.loc[self.gspreadWrapper.dfVcaAssessors['assessor'] == id, self.opt.topQualityCol] = row['constructiveFeedback']

        self.gspreadWrapper.dfVcaAssessors.fillna('', inplace=True)
        return self.gspreadWrapper.dfVcaAssessors


    def checkIntegrity(self, id, ass1, ass2):
        if (
            (ass1[self.opt.proposalIdCol] != ass2[self.opt.proposalIdCol]) or
            (ass1[self.opt.questionIdCol] != ass2[self.opt.questionIdCol]) or
            (ass1[self.opt.ratingCol] != ass2[self.opt.ratingCol]) or
            (ass1[self.opt.assessorCol] != ass2[self.opt.assessorCol])
        ):
            print("Something wrong with assessment {}".format(id))
            return False
        return True

    def checkIfMarked(self, row, column):
        if (row[column].strip() != ''):
            return 1
        return 0

    def calculateCards(self, row):
        yellow = 0
        red = 0
        tot = row[self.opt.noVCAReviewsCol]
        if (tot >= self.opt.minimumVCA):
            if ((row[self.opt.profanityCol]/tot) >= self.opt.cardLimit):
                red = red + 1
            for col in self.yellowColumns:
                if ((row[col]/tot) >= self.opt.cardLimit):
                    yellow = yellow + 1
        return (yellow, red)

    def goodFeedback(self, row):
        for col in self.positiveColumns:
            if (self.checkIfMarked(row, col) > 0):
                return True
        return False

    def badFeedback(self, row):
        for col in self.infringementsColumns:
            if (self.checkIfMarked(row, col) > 0):
                return True
        return False

    def badValid(self, row):
        if (
            (self.checkIfMarked(row, self.opt.otherCol) == 1) and
            (self.checkIfMarked(row, self.opt.otherRationaleCol) == 0)
        ):
            return False
        return True

    def neutralFeedback(self, row):
        for col in self.neutralColumns:
            if (self.checkIfMarked(row, col) > 0):
                return True
        return False

    def isVCAfeedbackValid(self, row, good, bad, neutral):
        if (bad):
            if (self.badValid(row) is False):
                return False
        if (sum([good, bad, neutral]) <= 1):
            return True
        return False




c = createVCAAggregate()
c.createDoc()
