from gspreadWrapper import GspreadWrapper
from options import Options
from utils import Utils

from gspread.models import Cell

class CreateVCAMaster():
    def __init__(self):
        self.options = Options()
        self.utils = Utils()
        self.gspreadWrapper = GspreadWrapper()

    def createDoc(self):
        print('Create new document...')
        spreadsheet = self.gspreadWrapper.gc.create(self.options.VCAMasterFileName)
        spreadsheet.share(
            self.options.accountEmail,
            perm_type='user',
            role='writer'
        )

        print('Create sheet...')
        worksheet = spreadsheet.add_worksheet(title="Flags", rows=1000, cols=17)

        cellsToAdd = []
        # Set headings
        print('Set headings...')
        headings = [
            self.options.assessmentsIdColumn, self.options.ideaURLColumn,
            self.options.assessorColumn, self.options.assessmentColumn,
            self.options.proposerMarkColumn, self.options.fairColumn,
            self.options.topQualityColumn, self.options.profanityColumn,
            self.options.scoreColumn, self.options.copyColumn,
            self.options.wrongChallengeColumn, self.options.wrongCriteriaColumn,
            self.options.otherColumn, self.options.otherRationaleColumn
        ]

        for i, value in enumerate(headings):
            cellsToAdd.append(
                Cell(row=1, col=(i + 1), value=value)
            )

        print('Set column width...')
        self.utils.setColWidth(
            spreadsheet,
            worksheet,
            1,
            len(headings) - 2,
            50
        )

        print('Load proposers flagged reviews...')
        assessments = self.gspreadWrapper.getProposersData()

        # extract Assessors
        assessors = self.gspreadWrapper.groupByAssessor(assessments)

        # filter assessors with more than allowed blank reviews.
        excludedAssessors = [k for k in assessors if (assessors[k]['blankPercentage'] >= self.options.allowedBlankPerAssessor)]
        includedAssessors = [k for k in assessors if (assessors[k]['blankPercentage'] < self.options.allowedBlankPerAssessor)]

        proposersDoc = self.gspreadWrapper.gc.open_by_key(self.options.proposersFile)
        self.gspreadWrapper.createSheetFromGroup(
            proposersDoc,
            'Excluded CAs',
            assessors,
            excludedAssessors,
            ['assessments']
        )

        self.gspreadWrapper.createSheetFromGroup(
            proposersDoc,
            'Included CAs',
            assessors,
            includedAssessors,
            ['assessments']
        )

        # Add sheet for excluded/included assessors

        index = 2
        print('Cloning flagged reviews...')
        for assessment in assessments:
            if (assessment[self.options.assessorColumn] not in excludedAssessors):
                marked = (
                    (assessment[self.options.profanityColumn] == 'x') or
                    (assessment[self.options.scoreColumn] == 'x') or
                    (assessment[self.options.copyColumn] == 'x') or
                    (assessment[self.options.wrongChallengeColumn] == 'x') or
                    (assessment[self.options.wrongCriteriaColumn] == 'x') or
                    (assessment[self.options.otherColumn] == 'x')
                )
                cellsToAdd.extend([
                    Cell(row=index, col=1, value=assessment[self.options.assessmentsIdColumn]),
                    Cell(row=index, col=2, value=assessment[self.options.ideaURLColumn]),
                    Cell(row=index, col=3, value=assessment[self.options.assessorColumn]),
                    Cell(row=index, col=4, value=assessment[self.options.assessmentColumn]),
                    Cell(row=index, col=5, value=marked)
                ])

                index = index + 1
        worksheet.update_cells(cellsToAdd, value_input_option='USER_ENTERED')
        print('Master Document for vCAs created')
        print('Link: {}'.format(spreadsheet.url))

cvca = CreateVCAMaster()
cvca.createDoc()
