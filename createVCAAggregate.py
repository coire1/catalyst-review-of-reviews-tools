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
        redCardsAssessors = list(redCards[self.opt.assessorCol].unique())


        # Select valid assessments (no red card assessors, no yellow card assessments, no blank assessments)
        validAssessments = self.dfVca[(
            (self.dfVca[self.opt.yellowCardCol] == 0) &
            ~self.dfVca[self.opt.assessorCol].isin(redCardsAssessors)
        )]
        validAssessments.to_csv('valid.csv')

        # Print valid assessments

        # Print non valid assessments -> reason, extract from proposer Doc
        self.gspreadWrapper.getProposersData()

        # Print assessors recap

        self.assessorRecap(redCardsAssessors)

        # Print vca recap

        with pd.option_context('display.max_rows', None):
            #toPrint = self.dfVca.loc[350:400]
            toPrint = self.dfVca[self.dfVca["# of vCAs Reviews"] > 4]
            print(toPrint[self.allColumns + [self.opt.noVCAReviewsCol, self.opt.yellowCardCol, self.opt.redCardCol]])

    def assessorRecap(self, redCardsAssessors):
        self.gspreadWrapper.dfVcaAssessors.loc[self.gspreadWrapper.dfVcaAssessors['assessor'].isin(redCardsAssessors), 'excluded'] = True
        self.gspreadWrapper.dfVcaAssessors.loc[self.gspreadWrapper.dfVcaAssessors['assessor'].isin(redCardsAssessors), 'note'] = "red card"

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

        self.gspreadWrapper.dfVcaAssessors.to_csv('assessors.csv')


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
