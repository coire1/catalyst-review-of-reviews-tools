from gspreadWrapper import GspreadWrapper
from options import Options
from utils import Utils

from gspread_formatting import *
import pandas as pd
import json

class CreateProposerDocument():
    def __init__(self):
        self.opt = Options()
        self.utils = Utils()
        self.gspreadWrapper = GspreadWrapper()

    def filterOutCAProposers(self, proposersDf):
        toInclude = []
        toExclude = []
        proposals = json.load(open('proposals.json'))
        users = json.load(open('users.json'))
        for id, row in proposersDf.iterrows():
            ass = row.to_dict()
            user = next((item for item in users if item["id"] == ass[self.opt.assessorCol]), None)
            proposal = next((item for item in proposals if item["id"] == ass[self.opt.proposalIdCol]), None)
            if (user and proposal):
                if (proposal["category"] in user["campaigns"]):
                    toExclude.append(ass)
                    print(ass)
                else:
                    toInclude.append(row)
            else:
                toInclude.append(ass)
        return pd.DataFrame(toInclude), pd.DataFrame(toExclude)


    def createDoc(self):
        pd.options.display.max_columns = 100
        print('Loading original...')
        self.gspreadWrapper.loadAssessmentsFile()
        proposerDf = self.gspreadWrapper.prepareDataFromExport()
        spreadsheet = self.gspreadWrapper.createDoc(
            self.opt.proposerDocumentName
        )
        # Define all the columns needed in the file
        headings = [
            self.opt.assessmentsIdCol,
            self.opt.proposalKeyCol, self.opt.ideaURLCol, self.opt.assessorCol,
            self.opt.tripletIdCol, self.opt.proposalIdCol,
            self.opt.q0Col, self.opt.q0Rating, self.opt.q1Col, self.opt.q1Rating,
            self.opt.q2Col, self.opt.q2Rating, self.opt.blankCol,
            self.opt.notValidCol, self.opt.notValidRationaleCol
        ]

        print('Assign blanks...')
        # Assign 'x' for blank assessments
        proposerDf[self.opt.blankCol] = proposerDf.apply(
            lambda r: 'x' if (
                (str(r[self.opt.q0Col]).strip() == "") or
                (str(r[self.opt.q1Col]).strip() == "") or
                (str(r[self.opt.q2Col]).strip() == "") or
                (str(r[self.opt.q0Rating]).strip() == "NA") or
                (str(r[self.opt.q1Rating]).strip() == "NA") or
                (str(r[self.opt.q2Rating]).strip() == "NA")
            ) else ''
        , axis=1)

        toInclude, toExclude = self.filterOutCAProposers(proposerDf)

        print('Format columns...')
        widths = [
            ('A', 30), ('B:C', 150), ('D', 100), ('E:F', 40), ('G', 300),
            ('H', 30), ('I', 300), ('J', 30), ('K', 300), ('L:N', 30),
            ('O', 300)
        ]

        formats = [
            ('A', self.utils.counterFormat),
            ('H', self.utils.counterFormat),
            ('J', self.utils.counterFormat),
            ('L', self.utils.counterFormat),
            ('M', self.utils.counterFormat),
            ('A1:O1', self.utils.headingFormat),
            ('M1:N1', self.utils.verticalHeadingFormat),
            ('H1', self.utils.verticalHeadingFormat),
            ('J1', self.utils.verticalHeadingFormat),
            ('L1', self.utils.verticalHeadingFormat),
            ('N2:N', self.utils.yellowFormat),
            ('G2:G', self.utils.textFormat),
            ('I2:I', self.utils.textFormat),
            ('K2:K', self.utils.textFormat),
        ]

        self.gspreadWrapper.createSheetFromDf(
            spreadsheet,
            "Assessments",
            toInclude,
            headings,
            widths,
            formats
        )

        if (len(toExclude)):
            self.gspreadWrapper.createSheetFromDf(
                spreadsheet,
                "Assessments Excluded (CA proposer in same challenge)",
                toExclude,
                headings,
                widths,
                formats
            )
        print('Master Document for proposers created')
        print('Link: {}'.format(spreadsheet.url))

c = CreateProposerDocument()
c.createDoc()
