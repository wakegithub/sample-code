import pandas as pd

missed_keys_dict = {'a': 'qwsxz', 'b': 'vghn', 'c': 'xdfv', 'd': 'serfcx', 'e': 'w34rfds', 'f': 'drtgvc',
                        'g': 'ftyhbv', 'h': 'gyujnb', 'i': 'u89olkj', 'j': 'huikmn', 'k': 'jiolm', 'l': 'kop',
                        'm': 'njk', 'n': 'bhjm', 'o': 'i90plk', 'p': 'o0l', 'q': '12wsa', 'r': 'e45tgfd', 's': 'aqwdxz',
                        't': 'r56yhgf', 'u': 'y78ikjh', 'v': 'cfgb', 'w': 'q23edsa', 'x': 'zsdc', 'y': 't67ujhg',
                        'z': 'asx'}

def Skipped_letters(word):
    temp_list = []
    for x in range(len(word)):
        misspelling = word[0:x] + word[x + 1:len(word)]
        if misspelling not in temp_list:
            temp_list.append(misspelling)
    return temp_list

def Double_letters(word):
    temp_list = []
    for x in range(len(word)):
        misspelling = word[0:x +1] + word[x] + word[x + 1:len(word)]
        if misspelling not in temp_list:
            temp_list.append(misspelling)
    return temp_list

def Reverse_letters(word):
    temp_list = []
    for x in range(0,len(word)-1,1):
        misspelling = word[0:x] + word[x+1] + word[x] + word[x+2:len(word)]
        if misspelling not in temp_list and misspelling != word:
            temp_list.append(misspelling)
    return temp_list

def Skipped_spaces(word):
    if ' ' in word:
        temp_list = []
        for x in range(len(word)):
            if word[x] == ' ':
                misspelling = word[0:x] + word[x+1:len(word)]
                if misspelling not in temp_list and misspelling != word:
                    temp_list.append(misspelling)
        return temp_list
    else:
        return []

def Missed_keys(word):
    temp_list = []
    for x in range(len(word)):
        letter = word[x]
        missed_keys_letters = missed_keys_dict[letter]
        for y in range(len(missed_keys_letters)):
            misspelling = word[0:x] + missed_keys_letters[y] + word[x + 1:len(word)]
            if misspelling not in temp_list:
                temp_list.append(misspelling)
    return temp_list

def Inserted_keys(word):
    temp_list = []
    for x in range(len(word)):
        letter = word[x]
        missed_keys_letters = missed_keys_dict[letter]
        for y in range(len(missed_keys_letters)):
            misspelling = word[0:x] + missed_keys_letters[y] + word[x:len(word)]
            if misspelling not in temp_list:
                temp_list.append(misspelling)
            misspelling2 = word[0:x+1] + missed_keys_letters[y] + word[x+1:len(word)]
            if misspelling2 not in temp_list:
                temp_list.append(misspelling2)
    return temp_list

#Open file
df_words = pd.read_csv(r'words.csv', encoding='utf-8', usecols=['Words'])
list_methods = [Skipped_letters, Double_letters, Reverse_letters, Skipped_spaces, Missed_keys, Inserted_keys]
words = []
for i in df_words.index:
    word = df_words[df_words.columns[0]][i].lower()
    words.append(word)

#Generate misspellings
misspellings = []
df_misspellings_all = pd.DataFrame(columns=['Word', 'Method', 'Misspelling'])
for word in words:
    word.strip()
    word.replace('\"', '\'')

    for m in list_methods:
        misspellings = m(word)
        df_misspellings = pd.DataFrame(misspellings,columns=['Misspelling'])
        df_misspellings.insert(0, 'Word', word)
        df_misspellings.insert(1, 'Method', m.__name__)
        df_misspellings_all = pd.concat([df_misspellings_all, df_misspellings], axis=0, ignore_index=True)

#Write file
df_misspellings_all.to_csv('misspellings.csv', index=False)

print('Done!')
