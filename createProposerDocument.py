from gspreadWrapper import GspreadWrapper
from options import Options
from utils import Utils

from gspread.models import Cell

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

        print('Setting headings for report...')
        # Add columns for y/r cards criteria
        currentColsCount = worksheet.col_count
        cellsToAdd = []
        # Set headings
        headings = [
            'Blank', 'Offense/Profanity', 'Score doesn\'t match',
            'Copy', 'Wrong challenge', 'Wrong criteria', 'Other',
            'Other rationale'
        ]
        worksheet.add_cols(len(headings))

        print('Set column width...')
        self.utils.setColWidth(
            spreadsheet,
            worksheet,
            currentColsCount,
            currentColsCount + len(headings),
            70
        )

        for i, value in enumerate(headings):
            cellsToAdd.append(
                Cell(row=1, col=(currentColsCount + i + 1), value=value)
            )

        print('Mark blank assessments...')
        # Autofill blank assessments
        assessments = self.gspreadWrapper.getAssessmentsData()
        for note in assessments:
            assessment = note[self.options.assessmentColumn].strip()
            if (assessment == ''):
                col = (currentColsCount + 1)
                cellsToAdd.append(
                    Cell(row=note[self.options.assessmentsIdColumn], col=col, value='x')
                )
        worksheet.update_cells(cellsToAdd, value_input_option='USER_ENTERED')
        print('Document for proposers created')
        print('Link: {}'.format(spreadsheet.url))



c = CreateProposerDocument()
c.createDoc()
