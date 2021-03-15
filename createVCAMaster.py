from gspreadWrapper import GspreadWrapper
from options import Options
from utils import Utils

from gspread.models import Cell
from gspread_formatting import *

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
        worksheet = spreadsheet.get_worksheet(0)
        worksheet.update_title("Asessments")

        cellsToAdd = []
        # Set headings
        print('Set headings...')
        headings = [
            self.options.assessmentsIdColumn, self.options.ideaURLColumn,
            self.options.questionColumn, self.options.ratingColumn,
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
        set_column_widths(worksheet, [
            ('A', 40), ('B:C', 200), ('D', 40), ('E', 120), ('F', 400),
            ('G:P', 30), ('Q', 300)
        ])

        print('Format columns')
        noteFormat = cellFormat(
            wrapStrategy='CLIP'
        )
        flagFormat = cellFormat(
            textFormat=textFormat(bold=True),
            horizontalAlignment='CENTER'
        )
        format_cell_ranges(worksheet, [
            ('D:D', flagFormat), ('F:F', noteFormat), ('G:P', flagFormat)]
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
                marked = 'x' if (
                    (assessment[self.options.profanityColumn] == 'x') or
                    (assessment[self.options.scoreColumn] == 'x') or
                    (assessment[self.options.copyColumn] == 'x') or
                    (assessment[self.options.wrongChallengeColumn] == 'x') or
                    (assessment[self.options.wrongCriteriaColumn] == 'x') or
                    (assessment[self.options.otherColumn] == 'x')
                ) else ''
                cellsToAdd.extend([
                    Cell(row=index, col=1, value=assessment[self.options.assessmentsIdColumn]),
                    Cell(row=index, col=2, value=assessment[self.options.ideaURLColumn]),
                    Cell(row=index, col=3, value=assessment[self.options.questionColumn]),
                    Cell(row=index, col=4, value=assessment[self.options.ratingColumn]),
                    Cell(row=index, col=5, value=assessment[self.options.assessorColumn]),
                    Cell(row=index, col=6, value=assessment[self.options.assessmentColumn]),
                    Cell(row=index, col=7, value=marked)
                ])

                index = index + 1
        worksheet.update_cells(cellsToAdd, value_input_option='USER_ENTERED')
        print('Master Document for vCAs created')
        print('Link: {}'.format(spreadsheet.url))

cvca = CreateVCAMaster()
cvca.createDoc()
