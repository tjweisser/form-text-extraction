import os
import docx2txt as dt
import pandas as pd

# select directory of documents with text to be extracted
path = os.chdir('/Users/User/Desktop/Pythonproject/extract/documents')

# get docx files
files = [file for file in os.listdir(path) if file.endswith('.docx')]

# extract specific portion and save to list
list = [(dt.process(files[text])[41:46], dt.process(files[text])[90:93])
        for text in range(len(files))]

# make dataframe from list
df = pd.DataFrame(list, columns=['team', 'grade'])

# and export to excel
df.to_csv('df.csv', encoding='utf-8', index=False)
