from gspreadWrapper import GspreadWrapper
import pandas as pd

class DownloadCsvFromDrive():
    def __init__(self):
        self.gspreadWrapper = GspreadWrapper()
        self.fileList = [
            "18D79kcl3jXkuCF43jtBdeiz59IXFkb_vvG8fBmzzq84"
        ]

    def downloadFiles(self):
        for gfile in self.fileList:
            doc = self.gspreadWrapper.gc.open_by_key(gfile)
            # doc = self.gspreadWrapper.gc.open_by_url(gfile)
            sheet = doc.worksheet("Assessments")
            df = pd.DataFrame(sheet.get_all_records())
            df.fillna('', inplace=True)
            df.to_csv('proposers-files/' + doc.title + '.csv', index=False)

down = DownloadCsvFromDrive()
down.downloadFiles()
