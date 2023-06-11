import re
import pandas as pd
from statistics import mean
from tqdm.auto import tqdm

def DF_find(df, word):
    try:
        index_found = df[df[df.columns[0]] == word].index.values.astype(int)[0]
    except:
        index_found = None
    return index_found

def Breakdown_phrase(phrase):
    phrase2 = re.sub('[\'\,\"\(\)\[\]\{\}\.\?\!\@\#\$\%\^\&\*\-\_\+\=\<\>\:]', ' WWW ', phrase)
    temp_list = re.split('\W', phrase2)
    temp_list2 = []
    for temp in temp_list:
        if temp != '':
            temp_list2.append(temp.lower())
    return temp_list2

def Search_name(breakdown):
    names = []
    for x in range(len(breakdown)-1):
        word1 = breakdown[x]
        word2 = breakdown[x+1]
        index1 = DF_find(df_firstnames,word1)
        if index1 is not None:
            vocab1 = DF_find(df_vocabulary, word1)
            if vocab1 == None:
                index2 = DF_find(df_lastnames,word2)
                if index2 is not None:
                    vocab2 = DF_find(df_vocabulary, word2)
                    if vocab2 == None:
                        score1 = 100 - df_firstnames.iloc[index1][2]
                        score2 = 100 - df_lastnames.iloc[index2][2]
                        if score1 + score2 >= min_name_score:
                            names.append(word1 + ' ' + word2 + ' (' + str(mean([score1, score2])) + '%)')
    names_string = ';'.join(names)
    return names_string

#Open files
min_name_score = 110
df_firstnames = pd.read_csv(r'firstnames.csv', encoding='utf-8')
df_lastnames = pd.read_csv(r'lastnames.csv', encoding='utf-8')
df_vocabulary = pd.read_csv(r'vocabulary.csv', encoding='utf-8')
df_phrases = pd.read_csv(r'phrases2.csv', encoding='utf-8')

#Search names
tqdm.pandas(desc='Phrase Breakdown', colour='green')
df_phrases['Breakdown'] = df_phrases.progress_apply(lambda x: Breakdown_phrase(x['Phrases']), axis=1)
tqdm.pandas(desc='Searching Names', colour='green')
df_phrases['Names Found'] = df_phrases.progress_apply(lambda x: Search_name(x['Breakdown']), axis=1)

#Write file
header = ['Phrases', 'Names Found']
df_phrases.to_csv('phrases2_names_found.csv', columns=header, index=False)

print('Done!')
