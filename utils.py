import json

class Utils():
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
