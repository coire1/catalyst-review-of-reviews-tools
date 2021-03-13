from gspreadWrapper import GspreadWrapper

from options import Options

class ServiceAccountUtils():
    def __init__(self):
        self.options = Options()
        self.gspreadWrapper = GspreadWrapper()

    def listAll(self):
        for spreadsheet in self.gspreadWrapper.gc.openall():
            print(spreadsheet)

    def deleteAll(self):
        for spreadsheet in self.gspreadWrapper.gc.openall():
            self.gspreadWrapper.gc.del_spreadsheet(spreadsheet.id)

sa = ServiceAccountUtils()
sa.listAll()
sa.deleteAll()
sa.listAll()
