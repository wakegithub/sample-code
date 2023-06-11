import re
import pandas as pd
from word_forms.word_forms import get_word_forms
from tqdm.auto import tqdm

def Breakdown_phrase(phrase):
    temp_list = re.split('[^a-zA-Z0-9\']', phrase)
    temp_list2 = []
    for temp in temp_list:
        if temp != '':
            temp_list2.append(temp)
    return temp_list2

def Get_forms(word):
    forms = get_word_forms(word)
    return(set([word] + list(forms['n']) + list(forms['a']) + list(forms['v']) + list(forms['r'])))

def Find_matches(phrase1, phrase2):
    breakdown1 = Breakdown_phrase(phrase1.lower())
    breakdown2 = Breakdown_phrase(phrase2.lower())
    breakdown1x = []
    breakdown2x = []
    for x in breakdown1:
        breakdown1x.append(Get_forms(x))
    for y in breakdown2:
        breakdown2x.append(Get_forms(y))

    temp1 = []
    temp2 = []
    for x in range(len(breakdown1x)):
        forms1 = breakdown1x[x]
        for y in range(len(breakdown2x)):
            forms2 = breakdown2x[y]
            if forms1.intersection(forms2):
                if x not in temp1:
                    temp1.append(x)
                if y not in temp2:
                    temp2.append(y)

    match1 = []
    for index1 in temp1:
        match1.append(breakdown1[index1])
    match2 = []
    for index2 in temp2:
        match2.append(breakdown2[index2])

    match_all1 = ';'.join(match1)
    match_all2 = ';'.join(match2)
    return([match_all1,
            str(round(100 * len(match1)/len(breakdown1))) + '%',
            match_all2,
            str(round(100 * len(match2)/len(breakdown2))) + '%'])

#Open file
df_phrases = pd.read_csv(r'text_compare.csv',encoding='utf-8')

#Find matches
tqdm.pandas(desc='Phrase Breakdown', colour='green')
df_phrases['Matches'] = df_phrases.progress_apply(lambda x: Find_matches(x['Phrase1'], x['Phrase2']), axis=1)
tqdm.pandas(desc='Adjusting Matches', colour='green')
df_phrases['Matches1'] = df_phrases.progress_apply(lambda x: x['Matches'][0], axis=1)
df_phrases['Matches1 %'] = df_phrases.progress_apply(lambda x: x['Matches'][1], axis=1)
df_phrases['Matches2'] = df_phrases.progress_apply(lambda x: x['Matches'][2], axis=1)
df_phrases['Matches2 %'] = df_phrases.progress_apply(lambda x: x['Matches'][3], axis=1)
df_final = df_phrases[['Phrase1', 'Matches1', 'Matches1 %', 'Phrase2', 'Matches2', 'Matches2 %']]

#Write file
df_final.to_csv('text_compare_matches.csv', index=False)

print('Done!')
