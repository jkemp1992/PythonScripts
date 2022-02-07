import nltk
# nltk.download('punkt')
# nltk.download('stopwords')

import matplotlib.pyplot as plt
import pandas
from nltk.collocations import *
from nltk.stem import PorterStemmer
from nltk.corpus import stopwords

# Testing manual counting of trigram
from operator import itemgetter
from collections import Counter


# Returns a list of (instance, count) sorted in total order and then from most to least common
def most_common(instances):
    return sorted(sorted(Counter(instances).items(), key=itemgetter(0)), key=itemgetter(1), reverse=True)

line = ""
final_list = []
data_list = []
open_file = open("C:/Users/Justin.Kemp/Python/test.txt", 'r')

# Open text file and read each line
for val in open_file:
    line += val

# Create 'tokens' from sentences
# tokens = line.split()
tokens = nltk.word_tokenize(line.lower())

# Using Stemmer to find base version of word
# ps = PorterStemmer()

# Common words like 'the', 'no', and 'am'
# stop_words = set(stopwords.words('english'))

# Replace unwanted characters in strings
# Add only new words to list (distinct final_list)
unwanted_chars = ".,-_ ()[]*?!/\\+:'"""
for word in tokens:
    new_word = word.strip(unwanted_chars)
    new_word = new_word.strip()
    if not new_word == "":
            # if new_word not in stop_words:
                # final_list.append(ps.stem(new_word))
        final_list.append(new_word)

# Create trigrams
tgs = nltk.trigrams(final_list)

# Compute frequency distribution for all the trigrams in the text
# Add to list variable
fdist = nltk.FreqDist(tgs)
for k, v in fdist.items():
    data_list.append([k, v])

# fdist.plot(30, cumulative=False)
# plt.bar()

# Count trigrams
total_trigrams = 0
for k, v in data_list:
    total_trigrams += int(v)

# Calculate frequency percentage
percent_list = []
for k, v in data_list:
    freq_percent = (int(v) / total_trigrams) * 100
    freq_percent = round(freq_percent, 2)
    percent_list.append([k, v, freq_percent])

# Create Pandas dataframe tables
df = pandas.DataFrame(data_list, columns=['Trigram', 'Frequency'])
df = df.sort_values(by='Frequency', ascending=False)

df2 = pandas.DataFrame(percent_list, columns=['Trigram', 'Frequency', '% Trigrams'])
df2 = df2.sort_values(by='Frequency', ascending=False)
print(df2[:50])
