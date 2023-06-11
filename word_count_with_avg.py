import re
import math
import pandas as pd

def Breakdown_phrase(phrase):
    temp_list = re.split('[^a-zA-Z0-9\']', phrase)
    return temp_list

def Generate_phrases(temp_list, max_word_length, numbers):
    phrases = [[], [], [], []]
    for x in range(0,max_word_length + 1,1):
        phrase = ''
        for y in range(x,max_word_length + 1,1):
            try:
                phrase += ' ' + temp_list[y]
                phrase = phrase.strip()
            except:
                phrase = None
            if phrase != None and len(phrase) > 0:
                word_length = len(phrase) - len(phrase.replace(' ','')) + 1
                if word_length > max_word_length:
                    break
                else:
                    phrases[word_length - 1].append([phrase, numbers])
    return phrases

def Generate_counts(file):
    #Open File
    df_temp = pd.read_csv(file, encoding='utf-8', usecols=['Keyword', 'Numbers'])
    counts = [[], [], [], []]

    #Generate phrases
    for i in df_temp.index:
        search_volume = df_temp[df_temp.columns[1]][i]
        breakdown = Breakdown_phrase(df_temp[df_temp.columns[0]][i].lower())
        phrases = Generate_phrases(breakdown, 4, search_volume)

        for x in range(len(phrases)):
            for pair in phrases[x]:
                counts[x].append(pair)

    #Count phrases
    counts_all = []
    for c in range(len(counts)):
        phrase = [x[0] for x in counts[c]]
        values = [x[1] for x in counts[c]]
        values = [0 if math.isnan(x) else x for x in values]

        df_phrase = pd.DataFrame(list(zip(phrase, values)),columns=['Phrase','Numbers'])
        df_phrase_count = df_phrase.groupby('Phrase').size().reset_index(name='Count')
        df_sv_sum = df_phrase.groupby('Phrase')['Numbers'].sum().reset_index(name='Numbers Sum')
        df_combined = pd.merge(left=df_phrase_count, right=df_sv_sum, left_on='Phrase', right_on='Phrase')
        df_combined.insert(1, 'Words' ,c + 1)
        df_combined['Numbers Avg'] = df_combined.apply(lambda x: round(x['Numbers Sum']/x['Count'],1), axis=1)
        df_combined = df_combined.sort_values(by=['Count'], ascending=False)
        df_combined = df_combined[df_combined['Count'] >= min_count]
        df_final = df_combined[['Phrase', 'Words', 'Count', 'Numbers Avg']]

        counts_all.append(df_final)
        print(str(c + 1) + '-counts done.')

    return counts_all

min_count = 2
df_file = Generate_counts('data.csv')
df_all = pd.concat(df_file, axis=0, ignore_index=True)
header = ['Phrase', 'Words', 'Count', 'Numbers Avg']
df_all.to_csv('data_counted.csv', columns=header, index=False)
print('Done!')
