from gspreadWrapper import GspreadWrapper
from options import Options
from utils import Utils

from gspread.models import Cell
from gspread_formatting import *

class createVCAAggregate():
    def __init__(self):
        self.options = Options()
        self.utils = Utils()
        self.gspreadWrapper = GspreadWrapper()

        self.infringementsColumns = [
            self.options.profanityColumn, self.options.scoreColumn,
            self.options.copyColumn, self.options.wrongChallengeColumn,
            self.options.wrongCriteriaColumn, self.options.otherColumn
        ]
        self.positiveColumns = [
            self.options.fairColumn, self.options.topQualityColumn
        ]
        self.feedbackColumns = self.infringementsColumns + self.positiveColumns

    def loadVCAsFiles(self):
        masterDocument = self.gspreadWrapper.gc.open_by_key(self.options.VCAMasterFile)
        masterSheet = masterDocument.worksheet("Assessments")
        masterData = masterSheet.get_all_records()
        self.masterDataByIds = self.gspreadWrapper.groupById(masterData)
        self.vcaRawData = []
        self.vcaData = []
        self.vcaDocs = []
        for vcaFile in self.options.VCAsFiles:
            vcaDocument = self.gspreadWrapper.gc.open_by_key(vcaFile)
            vcaSheet = vcaDocument.worksheet("Assessments")
            data = vcaSheet.get_all_records()
            dataByIds = self.gspreadWrapper.groupById(data)
            self.vcaData.append(dataByIds)
            self.vcaRawData.append(data)
            self.vcaDocs.append(vcaDocument)



    def createDoc(self):
        self.loadVCAsFiles()
        print('Create new document...')
        spreadsheet = self.gspreadWrapper.gc.create(self.options.VCAAggregateFileName)
        spreadsheet.share(
            self.options.accountEmail,
            perm_type='user',
            role='writer'
        )
        print('vCA aggregate document created.')

        print('Create sheet...')

        excludedAssessors = []
        validAssessments = []
        yellowAssessments = []
        assessments = []
        # Loop over master ids as reference
        for id in self.masterDataByIds:
            assessment = {}
            assessment[self.options.assessmentsIdColumn] = id
            assessment[self.options.tripletIdColumn] = self.masterDataByIds[id][self.options.tripletIdColumn]
            assessment[self.options.assessorColumn] = self.masterDataByIds[id][self.options.assessorColumn]
            assessment[self.options.noVCAReviewsColumn] = 0
            assessment[self.options.fairColumn] = 0
            assessment[self.options.topQualityColumn] = 0
            for col in self.infringementsColumns:
                assessment[col] = 0
            assessment[self.options.yellowCardColumn] = 0
            assessment[self.options.redCardColumn] = 0

            # Loop over all vca files
            for vcaFile in self.vcaData:
                fair = self.checkIfMarked(vcaFile[id], self.options.fairColumn)
                if (self.isVCAfeedbackValid(fair, vcaFile[id])):
                    assessment[self.options.noVCAReviewsColumn] = assessment[self.options.noVCAReviewsColumn] + self.checkIfReviewed(vcaFile[id])
                    assessment[self.options.fairColumn] = assessment[self.options.fairColumn] + fair
                    assessment[self.options.topQualityColumn] = assessment[self.options.topQualityColumn] + self.checkIfMarked(vcaFile[id], self.options.topQualityColumn)
                    for col in self.infringementsColumns:
                        assessment[col] = assessment[col] + self.checkIfMarked(vcaFile[id], col)

            (yellow, red) = self.calculateCards(assessment)
            assessment[self.options.yellowCardColumn] = yellow
            assessment[self.options.redCardColumn] = red

            if (red >= 1):
                excludedAssessors.append(assessment[self.options.assessorColumn])
            if (yellow >= 1):
                yellowAssessments.append(assessment)
            assessments.append(assessment)

        validAssessments, excludedAssessments = self.filterAssessments(yellowAssessments, excludedAssessors)

        self.gspreadWrapper.createSheetFromList(
            spreadsheet,
            'Aggregated',
            assessments,
            []
        )

        print(validAssessments[0])
        self.gspreadWrapper.createSheetFromList(
            spreadsheet,
            'Valid Assessments',
            validAssessments,
            [
                self.options.proposerMarkColumn, self.options.fairColumn,
                self.options.topQualityColumn, self.options.profanityColumn,
                self.options.scoreColumn, self.options.copyColumn,
                self.options.wrongChallengeColumn, self.options.wrongCriteriaColumn,
                self.options.otherColumn, self.options.otherRationaleColumn,
                self.options.proposerMarkColumn
            ]
        )

        # Create sheet for invalid assessemnts
        self.gspreadWrapper.createSheetFromList(
            spreadsheet,
            'Excluded Assessments (r/y cards)',
            excludedAssessments,
            [
                self.options.proposerMarkColumn, self.options.fairColumn,
                self.options.topQualityColumn, self.options.profanityColumn,
                self.options.scoreColumn, self.options.copyColumn,
                self.options.wrongChallengeColumn, self.options.wrongCriteriaColumn,
                self.options.otherColumn, self.options.otherRationaleColumn,
                self.options.proposerMarkColumn
            ]
        )


        # Create sheet with excluded assessors (merge excluded assessor for
        # blank ratio + excluded assessors for red card)
        originalAssessments = self.gspreadWrapper.getProposersData()
        originalAssessors = self.gspreadWrapper.groupByAssessor(originalAssessments)
        blankExcludedAssessors = {}
        blankIncludedAssessors = {}
        for k in originalAssessors:
            if (originalAssessors[k]['blankPercentage'] >= self.options.allowedBlankPerAssessor):
                blankExcludedAssessors[k] = originalAssessors[k]
            else:
                blankIncludedAssessors[k] = originalAssessors[k]

        mergedExcludedAssessors = self.mergeExcludedAssessors(
            excludedAssessors,
            blankExcludedAssessors,
            blankIncludedAssessors
        )
        self.gspreadWrapper.createSheetFromList(
            spreadsheet,
            'Excluded assessors',
            mergedExcludedAssessors,
            []
        )

        # Append each VCA sheet to the current document.
        for i, vcaRawData in enumerate(self.vcaRawData):
            self.gspreadWrapper.createSheetFromList(
                spreadsheet,
                self.vcaDocs[i].title,
                vcaRawData,
                []
            )

        worksheet = spreadsheet.get_worksheet(0)
        spreadsheet.del_worksheet(worksheet)

        print('Aggregated document created')

        print('Link: {}'.format(spreadsheet.url))

    def checkIfReviewed(self, row):
        result = False
        for col in self.feedbackColumns:
            result = result or (row[col] == 'x')
        if (result):
            return 1
        return 0

    def checkIfMarked(self, row, column):
        if (row[column] == 'x'):
            return 1
        return 0

    def calculateCards(self, row):
        yellow = 0
        red = 0
        tot = row[self.options.noVCAReviewsColumn]
        if (tot >= self.options.minimumVCA):
            if ((row[self.options.profanityColumn]/tot) >= self.options.profanityLimit):
                red = red + 1
            if ((row[self.options.scoreColumn]/tot) >= self.options.scoreLimit):
                yellow = yellow + 1
            if ((row[self.options.copyColumn]/tot) >= self.options.copyLimit):
                yellow = yellow + 1
            if ((row[self.options.wrongChallengeColumn]/tot) >= self.options.wrongChallengeLimit):
                yellow = yellow + 1
            if ((row[self.options.wrongCriteriaColumn]/tot) >= self.options.wrongCriteriaLimit):
                yellow = yellow + 1
            if ((row[self.options.otherColumn]/tot) >= self.options.otherLimit):
                yellow = yellow + 1
        return (yellow, red)

    def isVCAfeedbackValid(self, fairCount, row):
        if (fairCount == 1):
            for col in self.infringementsColumns:
                if (self.checkIfMarked(row, col) > 0):
                    return False
        return True


    def filterAssessments(self, yellowAssessments, excludedAssessors):
        filtered = []
        excluded = []
        yellowRelatedTripletsIds = self.getTripletsIds(yellowAssessments)
        assessments = self.masterDataByIds
        for id in assessments:
            if (
                (assessments[id][self.options.tripletIdColumn] not in yellowRelatedTripletsIds) and
                (assessments[id][self.options.assessorColumn] not in excludedAssessors)
            ):
                filtered.append(assessments[id])
            else:
                excluded.append(assessments[id])

        return filtered, excluded

    def getTripletsIds(self, assessments):
        tripletIds = [el[self.options.tripletIdColumn] for el in assessments]
        return tripletIds

    def mergeExcludedAssessors(self, excludedByCard, excludedByBlank, includedByBlank):
        assessors = []
        for assessor in excludedByCard:
            assessors.append({
                'name': assessor,
                'By Card': 1,
                'By Blanks': '',
                'Blank percentage': includedByBlank[assessor]['blankPercentage'],
            })
        for k in excludedByBlank:
            assessors.append({
                'name': k,
                'By Card': '',
                'By Blanks': 1,
                'Blank percentage': excludedByBlank[k]['blankPercentage'],
            })
        return assessors




c = createVCAAggregate()
c.createDoc()
