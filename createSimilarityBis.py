import os
import sys
#sys.stderr = open(os.devnull, "w")  # silence stderr

from gspreadWrapper import GspreadWrapper
from options import Options
from utils import Utils

from gspread.models import Cell
from gspread_formatting import *
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

import pandas as pd

import time

opt = Options()
utils = Utils()
gspreadWrapper = GspreadWrapper()

print('Load vca data...')
gspreadWrapper.getProposersMasterData()

assessors = {}
similarities = []

similarityMinScore = 0.6


vectorize = lambda Text: TfidfVectorizer().fit_transform(Text).toarray()

def analyze_sentence(assessor_a, text_vector_a, assessor_b, text_vector_b, progress, k, notes):
    sys.stderr = open(os.devnull, "w")  # silence stderr
    #internal_index = s_vectors.index((assessor_b, text_vector_b))
    char_diff = abs(len(notes[progress]) - len(notes[k]))
    if (progress != k and char_diff < 125):
        sim_score = cosine_similarity([text_vector_a, text_vector_b])[0][1]
        assessor_pair = sorted((assessor_a, assessor_b))
        score = (assessor_pair[0], assessor_pair[1], sim_score)
        return score
    return (0,0,0)

data = gspreadWrapper.dfMasterProposers
criteria = [opt.q0Col, opt.q1Col, opt.q2Col]

start_time = time.perf_counter()
for criterium in criteria:
    notes = list(data[criterium])
    ids = list(data[opt.assessmentsIdCol])
    vectors = vectorize(notes)
    s_vectors = list(zip(ids, vectors))
    plagiarism_results = set()
    progress = 0
    for assessor_a, text_vector_a in s_vectors:
        print("{} of {}".format(progress, len(s_vectors)))
        k = 0
        for assessor_b, text_vector_b in s_vectors:
            score = analyze_sentence(assessor_a, text_vector_a, assessor_b, text_vector_b, progress, k, notes)
            plagiarism_results.add(score)
            k = k + 1
        progress = progress + 1
    for res in plagiarism_results:
        if (res[2] > similarityMinScore):
            ass_0 = data.loc[data[opt.assessmentsIdCol] == res[0]][opt.assessorCol].item()
            ass_1 = data.loc[data[opt.assessmentsIdCol] == res[1]][opt.assessorCol].item()
            assessment_0 = data.loc[data[opt.assessmentsIdCol] == res[0]][criterium].item()
            assessment_1 = data.loc[data[opt.assessmentsIdCol] == res[1]][criterium].item()
            if (ass_0 not in assessors):
                assessors[ass_0] = {
                    'Assessor': ass_0,
                    'similarity_other_assessors': [],
                    'similarity_count_others': 0,
                    'similarity_count_self': 0
                }
            if (ass_0 != ass_1):
                assessors[ass_0]['similarity_other_assessors'].append(ass_1)
                assessors[ass_0]['similarity_count_others']  = assessors[ass_0]['similarity_count_others'] + 1
            else:
                assessors[ass_0]['similarity_count_self']  = assessors[ass_0]['similarity_count_self'] + 1
            similarities.append({
                'id A': res[0],
                'id B': res[1],
                'Assessor A': ass_0,
                'Assessor B': ass_1,
                'Note A': assessment_0,
                'Note B': assessment_1,
                'Similarity Score': res[2]
            })

end_time = time.perf_counter()
print(end_time - start_time, "seconds")

for k in assessors:
    assessors[k]['similarity_other_assessors'] = ','.join(list(dict.fromkeys(assessors[k]['similarity_other_assessors'])))

assessors = list(assessors.values())

dfSimilarities = pd.DataFrame(similarities)
dfAssessors = pd.DataFrame(assessors)

dfSimilarities.to_csv('cache/sim8-light.csv')
dfAssessors.to_csv('cache/sim_ass8-light.csv')

spreadsheet = gspreadWrapper.createDoc('Similarity Analysis')

gspreadWrapper.createSheetFromDf(
    spreadsheet,
    'Assessments',
    dfSimilarities,
    ['id A', 'id B', 'Assessor A', 'Assessor B', 'Note A', 'Note B', 'Similarity Score'],
    columnWidths=[('A:B', 50), ('C:D', 150), ('E:F', 300), ('G', 60)],
    formats=[('G', utils.counterFormat),('A1:G1', utils.headingFormat)]
)

gspreadWrapper.createSheetFromDf(
    spreadsheet,
    'CAs',
    dfAssessors,
    ['Assessor', 'similarity_other_assessors', 'similarity_count_others', 'similarity_count_self'],
    columnWidths=[('A:B', 150), ('C:D', 60)],
    formats=[
        ('C:D', utils.counterFormat),
        ('A1:D1', utils.headingFormat),
        ('C1:D1', utils.verticalHeadingFormat),
    ]
)
worksheet = spreadsheet.get_worksheet(0)
spreadsheet.del_worksheet(worksheet)

print('Link: {}'.format(spreadsheet.url))
