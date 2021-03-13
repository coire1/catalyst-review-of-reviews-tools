import json

class Utils():
    def setColWidth(self, spreadsheet, worksheet, startIndex, endIndex, size):
        sheetId = worksheet._properties['sheetId']
        body = {
            "requests": [
                {
                    "updateDimensionProperties": {
                        "range": {
                            "sheetId": sheetId,
                            "dimension": "COLUMNS",
                            "startIndex": startIndex,
                            "endIndex": endIndex
                        },
                        "properties": {
                            "pixelSize": size
                        },
                        "fields": "pixelSize"
                    }
                }
            ]
        }
        res = spreadsheet.batch_update(body)

    '''
    saveCache() saves the pulled records in a json file to cache the response.
    '''
    def saveCache(self, dicts, name):
        print('Saving cache..')
        with open('cache/' + name + '.json', 'w') as f:
            json.dump(dicts, f)

    '''
    loadCache() get records from cache if present.
    '''
    def loadCache(self, name):
        try:
            with open('cache/' + name + '.json', 'r') as f:
                data = json.load(f)
            return data
        except:
            return False
