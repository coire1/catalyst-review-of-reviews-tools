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
            'id', 'Offense/Profanity (Proposer)', 'Offense/Profanity (vCA)',
            'Score doesn\'t match (Proposer)', 'Score doesn\'t match (vCA)',
            'Copy (Proposer)', 'Copy (vCA)', 'Wrong challenge (Proposer)',
            'Wrong challenge (vCA)', 'Wrong criteria (Proposer)',
            'Wrong criteria (vCA)', 'Other (Proposer)', 'Other rationale (Proposer)',
            'Other (vCA)', 'Other rationale (vCA)', 'Assessment', 'Assessor'
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
        progressiveIndex = 2
        print('Cloning flagged reviews...')
        for note in assessments:
            assessment = note[self.options.assessmentColumn].strip()
            # Exclude blank assessments from vCA master file
            if (assessment != ''):
                cellsToAdd.extend([
                    Cell(row=progressiveIndex, col=1, value=note[self.options.assessmentsIdColumn]),
                    Cell(row=progressiveIndex, col=2, value=note['Offense/Profanity']),
                    Cell(row=progressiveIndex, col=4, value=note['Score doesn\'t match']),
                    Cell(row=progressiveIndex, col=6, value=note['Copy']),
                    Cell(row=progressiveIndex, col=8, value=note['Wrong challenge']),
                    Cell(row=progressiveIndex, col=10, value=note['Wrong criteria']),
                    Cell(row=progressiveIndex, col=12, value=note['Other']),
                    Cell(row=progressiveIndex, col=13, value=note['Other rationale']),
                    Cell(row=progressiveIndex, col=16, value=note[self.options.assessmentColumn]),
                    Cell(row=progressiveIndex, col=17, value=note[self.options.assessorColumn]),
                ])

                progressiveIndex = progressiveIndex + 1
        worksheet.update_cells(cellsToAdd, value_input_option='USER_ENTERED')
        print('Master Document for vCAs created')
        print('Link: {}'.format(spreadsheet.url))

cvca = CreateVCAMaster()
cvca.createDoc()
