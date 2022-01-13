from gspreadWrapper import GspreadWrapper
import gdown
import pandas as pd
import gspread
import json
import re
from time import sleep
from options import Options
from utils import Utils
import os

class FlashAssessmentCounter():
    def __init__(self):
        self.opt = Options()
        self.utils = Utils()
        self.gspreadWrapper = GspreadWrapper()
        doc = self.gspreadWrapper.gc.open_by_key(self.opt.flashResponses)
        worksheet = doc.worksheet('Form Responses 1')
        self.responses = worksheet.get_all_values()

        with open("flash-stage-proposals.json") as file:
            self.proposals = json.load(file)
        with open("proposals.json") as file:
            self.fullProposals = json.load(file)
        with open("categories.json") as file:
            self.categories = json.load(file)

    def get_proposal_name(self, response):
        name = ''
        for index in range(4,24):
            tname = response[index]
            if (tname != ''):
                name = tname
                break
        return name


    def start(self):
        for proposal in self.proposals:
            f_proposal = next((item for item in self.fullProposals if item["id"] == proposal["proposal_id"]), None)
            category = next((item for item in self.categories if item["id"] == f_proposal["category"]), None)
            proposal['challenge'] = category['title']
        for response in self.responses:
            name = self.get_proposal_name(response)
            proposal = next((item for item in self.proposals if item["title"] in name), None)
            if (proposal):
                pr_index = self.proposals.index(proposal)
                if pr_index > -1:
                    self.proposals[pr_index]['no_assessments'] = self.proposals[pr_index]['no_assessments'] + 1
            else:
                print("{} proposal not found".format(name))

        results = newlist = sorted(self.proposals, key=lambda d: d['no_assessments'])
        full = pd.DataFrame(results)
        full[['title', 'no_assessments', 'challenge']].to_csv('cache/flash-results.csv', index=False)
        for pp in results:
            print("{},{},{}".format(pp["no_assessments"], pp["title"], pp["challenge"]))
            print("# assessments: {} - {}".format(pp["no_assessments"], pp["title"]))


flash = FlashAssessmentCounter()
flash.start()
