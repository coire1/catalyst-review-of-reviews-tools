from gspreadWrapper import GspreadWrapper
from options import Options
from utils import Utils

from gspread.models import Cell
from gspread_formatting import *

from time import sleep
import pandas as pd
import os

class createProposersAggregate():
    def __init__(self):
        self.opt = Options()
        self.utils = Utils()
        self.gspreadWrapper = GspreadWrapper()
        self.proposersFiles = []

        self.allColumns = [self.opt.notValidCol]

    def prepareBaseData(self):
        self.gspreadWrapper.getProposersMasterData()
        self.dfMasterProposers = self.gspreadWrapper.dfMasterProposers.set_index('id')
        self.dfMasterProposers[self.opt.proposerMarkCol] = 0
        self.dfMasterProposers[self.opt.proposersRationaleCol] = ''

    def prepareProposersFileList(self):
        for currentDirPath, currentSubdirs, currentFiles in os.walk('./proposers-files'):
            for aFile in currentFiles:
                if aFile.endswith(".csv") :
                    fpath = str(os.path.join(currentDirPath, aFile))
                    self.proposersFiles.append(fpath)

    def loadProposersFiles(self):
        self.prepareBaseData()
        self.prepareProposersFileList()
        self.proposersData = []
        for proposerFile in self.proposersFiles:
            print(proposerFile)
            data = pd.read_csv(proposerFile)
            data.set_index(self.opt.assessmentsIdCol, inplace=True)
            data.fillna('', inplace=True)
            self.proposersData.append(data)

    def createDoc(self):
        self.loadProposersFiles()
        # Loop over master ids as reference
        for id, row in self.dfMasterProposers.iterrows():
            # Loop over all vca files
            for proposerDf in self.proposersData:
                if (id in proposerDf.index):
                    locAss = proposerDf.loc[id]
                    integrity = self.checkIntegrity(id, row, locAss)
                    if (integrity is False):
                        print('Error')
                        break
                        break

                    if (self.isProposerFeedbackValid(locAss)):
                        col = self.opt.notValidCol
                        colVal = self.checkIfMarked(locAss, col)
                        if (colVal > 0):
                            self.dfMasterProposers.loc[id, self.opt.proposerMarkCol] = self.dfMasterProposers.loc[id, self.opt.proposerMarkCol] + colVal
                            self.dfMasterProposers.loc[id, self.opt.proposersRationaleCol] = locAss[self.opt.notValidRationaleCol]

        self.dfMasterProposers[self.opt.assessmentsIdCol] = self.dfMasterProposers.index
        self.dfMasterProposers.to_csv('cache/test-proposers-aggregate.csv')
        spreadsheet = self.gspreadWrapper.createDoc(self.opt.proposersAggregateFileName)

        # Print valid assessments
        assessmentsHeadings = [
            self.opt.assessmentsIdCol,
            self.opt.proposalKeyCol, self.opt.ideaURLCol, self.opt.assessorCol,
            self.opt.tripletIdCol, self.opt.proposalIdCol,
            self.opt.q0Col, self.opt.q0Rating, self.opt.q1Col, self.opt.q1Rating,
            self.opt.q2Col, self.opt.q2Rating, self.opt.blankCol,
            self.opt.proposerMarkCol, self.opt.proposersRationaleCol
        ]
        assessmentsWidths = [
            ('A', 40), ('B', 60), ('C', 120), ('D', 120), ('E', 40), ('F', 200), ('G', 40), ('H', 60), ('I', 120), ('J', 400),
            ('K:Z', 30)
        ]
        assessmentsFormats = [
            ('H:H', self.utils.counterFormat),
            ('J:J', self.utils.noteFormat),
            ('K:Z', self.utils.counterFormat),
            ('A1:X1', self.utils.headingFormat),
            ('B1', self.utils.verticalHeadingFormat),
            ('E1', self.utils.verticalHeadingFormat),
            ('G1:H1', self.utils.verticalHeadingFormat),
            ('K1:Z1', self.utils.verticalHeadingFormat),
            ('L2:L', self.utils.greenFormat),
            ('Q2:Q', self.utils.redFormat),
            ('R2:W', self.utils.yellowFormat),
            ('Y2:Y', self.utils.yellowFormat),
            ('Z2:Z', self.utils.redFormat),
        ]

        self.gspreadWrapper.createSheetFromDf(
            spreadsheet,
            'Assessments',
            self.dfMasterProposers,
            assessmentsHeadings,
            #columnWidths=assessmentsWidths,
            #formats=assessmentsFormats
        )

        print('Proposers Aggregated Document created')
        print('Link: {}'.format(spreadsheet.url))

    def checkIntegrity(self, id, ass1, ass2):
        if (
            (ass1[self.opt.proposalIdCol] != ass2[self.opt.proposalIdCol]) or
            (ass1[self.opt.q0Rating] != ass2[self.opt.q0Rating]) or
            (ass1[self.opt.q1Rating] != ass2[self.opt.q1Rating]) or
            (ass1[self.opt.q2Rating] != ass2[self.opt.q2Rating]) or
            (ass1[self.opt.assessorCol] != ass2[self.opt.assessorCol])
        ):
            print("Something wrong with assessment {}".format(id))
            return False
        return True

    def checkIfMarked(self, row, column):
        if (row[column].strip() != ''):
            return 1
        return 0

    def badFeedback(self, row):
        for col in self.infringementsColumns:
            if (self.checkIfMarked(row, col) > 0):
                return True
        return False

    def badValid(self, row):
        if (
            (self.checkIfMarked(row, self.opt.notValidCol) == 1) and
            (self.checkIfMarked(row, self.opt.notValidRationaleCol) == 0)
        ):
            return False
        return True

    def isProposerFeedbackValid(self, row):
        return self.badValid(row)

c = createProposersAggregate()
c.createDoc()
