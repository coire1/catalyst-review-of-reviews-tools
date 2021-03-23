from gspreadWrapper import GspreadWrapper
from options import Options
from utils import Utils

from gspread.models import Cell
from gspread_formatting import *

class CreateProposerDocument():
    def __init__(self):
        self.options = Options()
        self.utils = Utils()
        self.gspreadWrapper = GspreadWrapper()

    def createDoc(self):
        print('Loading original...')
        self.gspreadWrapper.loadAssessmentsFile()
        print('Make a new copy...')
        spreadsheet = self.gspreadWrapper.gc.copy(
            self.options.originalExportFromIdeascale,
            title=self.options.proposerDocumentName
        )
        spreadsheet.share(
            self.options.accountEmail,
            perm_type='user',
            role='writer'
        )

        worksheet = spreadsheet.worksheet(self.options.assessmentsSheet)

        # Delete column for avg rating
        worksheet.delete_columns(5)

        print('Setting headings for report...')
        # Add columns for y/r cards criteria
        currentColsCount = worksheet.col_count - 1
        cellsToAdd = []
        # Set headings
        headings = [
            self.options.assessmentsIdColumn, self.options.tripletIdColumn,
            self.options.blankColumn, self.options.topQualityColumn,
            self.options.profanityColumn, self.options.scoreColumn,
            self.options.copyColumn, self.options.wrongChallengeColumn,
            self.options.wrongCriteriaColumn, self.options.otherColumn,
            self.options.otherRationaleColumn
        ]
        worksheet.add_cols(len(headings))

        print('Set column width...')
        set_column_widths(worksheet, [
            ('A:D', 150), ('E', 120), ('F', 400),
            ('G', 60), ('H', 30), ('I', 60), ('J:Q', 30), ('R', 400)
        ])

        for i, value in enumerate(headings):
            cellsToAdd.append(
                Cell(row=1, col=(currentColsCount + i + 1), value=value)
            )

        print('Mark blank assessments...')
        # Autofill blank assessments
        assessments = self.gspreadWrapper.getAssessmentsData()
        for note in assessments:
            col = (currentColsCount + 1)
            cellsToAdd.append(
                Cell(row=note[self.options.assessmentsIdColumn], col=col, value=note[self.options.assessmentsIdColumn])
            )
            col = (currentColsCount + 2)
            cellsToAdd.append(
                Cell(row=note[self.options.assessmentsIdColumn], col=col, value=note[self.options.tripletIdColumn])
            )
            assessment = note[self.options.assessmentColumn].strip()
            if (assessment == ''):
                col = (currentColsCount + 3)
                cellsToAdd.append(
                    Cell(row=note[self.options.assessmentsIdColumn], col=col, value='x')
                )
        worksheet.update_cells(cellsToAdd, value_input_option='USER_ENTERED')
        worksheet.freeze(rows=1)

        format_cell_ranges(worksheet, [
            ('H:Q', self.utils.counterFormat),
            ('A1:G1', self.utils.headingFormat),
            ('R1', self.utils.headingFormat),
            ('G1:Q1', self.utils.verticalHeadingFormat),
            ('K2:K', self.utils.greenFormat),
            ('L2:L', self.utils.redFormat),
            ('M2:Q', self.utils.yellowFormat),
            ('A2:G', self.utils.textFormat),

        ])
        print('Document for proposers created')
        print('Link: {}'.format(spreadsheet.url))



c = CreateProposerDocument()
c.createDoc()
