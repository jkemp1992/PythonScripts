from os import error
import nltk
import pandas as pd
from nltk.collocations import *
import pyodbc

# Script pulls data from SQL Server, loops through a top level group (Main Keyword1), then creates sub-groups (String) and usage frequency
# Used to classify support ticket data and also server monitoring alert data

# Function that replace unwanted start/end characters in strings
def clean_string(k: str) -> str:
    if k.startswith("[ "): k = k.replace("[ ", "")
    elif k.startswith("( "): k = k.replace("( ", "")
    
    if k.endswith(" )"): k = k.replace(" )", "")
    elif k.endswith(" ]"): k = k.replace(" ]", "")
    
    return k


# Function converts NLTK freq dist tuple object into a list, then builds dataframe
def compute_frequency_dataframe(my_freq: nltk.probability.FreqDist) -> pd.DataFrame:
    temp_list = []
    # Loops through key, value items in tuple
    for k, v in my_freq.items():
        # If k has more than 1 string inside a tuple, join them
        if isinstance(k, tuple):
            k = ' '.join(k)  
            # If k is a string now, replace start/end special characters
            if isinstance(k, str): k = clean_string(k)
        temp_list.append([k, v])
    df_freq = pd.DataFrame(temp_list, columns=['String', 'Frequency'])
    df_freq['Main Keyword'] = keyword1
    
    return df_freq


SQLString = """
            Select * 
            From MY_SQLTABLE as t (nolock) 
            INNER JOIN Table2 r on r.ID = t.ID
            """

# Connection to SQL
SQL_Server_CN = pyodbc.connect('Driver={SQL Server};'
                                'Server=MY-SQLSERVER;'
                                'Database=MY-DB;'
                                f'UID=username;'
                                f'PWD=password;'
                                'Trusted_Connection=no;'
                                )

# Execute SQL query and add to local SQL database
df = pd.read_sql_query(SQLString, SQL_Server_CN)

# If column is string/object make it lowercase
try: df['combined'] = df['combined'].str.lower()
except AttributeError: pass

# Create distinct list of keywords and ticket types
main_kewords = list(df['Keyword1'].unique())

# Create final dataframe
df_final = pd.DataFrame(columns={'String', 'Frequency', 'Main Keyword'})

# Loop through each group and analyze the data
for keyword1 in main_kewords:
    word_list = []

    # Filter the table, split checks by space delimeter, and create a list of words
    df_filtered = df.loc[df['Keyword1'] == keyword1]
    word_list = df_filtered['combined'].to_list()
    tokens = nltk.word_tokenize(' '.join(word_list))

    # Compute frequency distribution and build dataframe
    fdist = nltk.FreqDist(tokens)
    # fdist = fdist.most_common(50)  # Top 50 keywords
    df_final = df_final.append(compute_frequency_dataframe(fdist))
    filter_list = [":", "-", "_", "\\", "/", ";"]
    df_final = df_final.where(~df_final["String"].isin(filter_list))

    # Compute bigram frequency distribution
    bigrams = nltk.bigrams(tokens)
    fdist = nltk.FreqDist(bigrams)
    df_final = df_final.append(compute_frequency_dataframe(fdist))

    # Compute trigram frequency distribution
    trigrams = nltk.trigrams(tokens)
    fdist = nltk.FreqDist(trigrams)
    df_final = df_final.append(compute_frequency_dataframe(fdist))

df_final = df_final.groupby(['String', 'Main Keyword'])['Frequency'].sum().reset_index()
df_final.sort_values(by='Frequency', ascending=False, inplace=True)
df_final = df_final[df_final['Frequency'] > 2]
df_final.to_csv("ExampleFile.csv", index=False)
