import re
import os
import pandas as pd
import docx2txt as doc
import numpy as np

# define directory with documents with text to be extracted
directory = ('/Users/User/documents')

# define regular expressions
groupsPattern = re.compile(r'(Group:)\s?(\d+\.?\d+)')
gradesPattern = re.compile(r'(Grade:)\s?(\d\d?\.?\d?)')

# set path to directory with documents with text to be extracted
path = os.chdir(directory)

# get docx files
files = [file for file in os.listdir(path) if file.endswith('.docx')]

# loop through files and populate two lists (grades and groups)
grades = []
groups = []
for file in files:
    text = doc.process(file)
    # getting every second item from list of tuples [x[1] for x in list]
    groupMatch = [match[1] for match in groupsPattern.findall(text)]
    gradeMatch = [match[1] for match in gradesPattern.findall(text)]
    groups.append(groupMatch)
    grades.append(gradeMatch)

# transform list of lists to flat list
groupsAsList = [item for sublist in groups for item in sublist]
gradesAsList = [item for sublist in grades for item in sublist]

# make dataframe from list
df = pd.DataFrame(np.column_stack([groupsAsList, gradesAsList]),
                  columns=['Group', 'Grade'])

# and export to csv
try:
    df.to_csv('df.csv', encoding='utf-8', index=False)
    print("Table successfully generated!")
except Exception as e:
    print(e)
