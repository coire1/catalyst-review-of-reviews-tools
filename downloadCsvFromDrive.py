from gspreadWrapper import GspreadWrapper
import pandas as pd
import gspread

class DownloadCsvFromDrive():
    def __init__(self):
        self.gspreadWrapper = GspreadWrapper()
        self.fileList = [
        ]

    def downloadFiles(self):
        for gfile in self.fileList:
            try:
                print("\n######\n")
                print("Downloading: {}".format(gfile))
                doc = self.gspreadWrapper.gc.open_by_key(gfile)
                # doc = self.gspreadWrapper.gc.open_by_url(gfile)
                sheet = doc.worksheet("Assessments")
                df = pd.DataFrame(sheet.get_all_records())
                df.fillna('', inplace=True)
                df.to_csv('proposers-files/' + doc.title + '.csv', index=False)
                print("Downloaded successfully: {}".format(gfile))
            except gspread.exceptions.APIError as e:
                print("GDrive error downloading: {}".format(gfile))
                print(e)
            except:
                print("Generic error downloading: {}".format(gfile))


down = DownloadCsvFromDrive()
down.downloadFiles()
