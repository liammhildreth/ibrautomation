from textblob import TextBlob
import xlrd
import spacy
import docx2txt
import pandas as pd
from stanza.server import CoreNLPClient
import re
import nltk
import stanza
from pycorenlp import *

nlp = StanfordCoreNLP("http://localhost:9000/")

# input word doc from your C:\Users drive
article = input("Word document name: ")
text = docx2txt.process(str(article) + ".docx")
nlpfacts = []

# applying nlp library to the article
processed = TextBlob(text)

# parsing through the article to pull all sentences
for line in processed.sentences:
    if line != '':
        sent = re.sub(r"(\.|,|\?|\(|\)|\[|\])", " ", line.string)
        if sent != '' and sent != ' 'and len(sent)!=1:
            sent = TextBlob(sent)
            
            # setting threshold to 1 in order to pull every single sentence
            if sent.subjectivity <= 1:
                nlpfacts.append(line)


# creating pandas dataframe to store the sentences to be uploaded to factchecking spreadsheet                
df = pd.DataFrame()
df['Facts']= nlpfacts

# exporting sentences to pre-existing spreadsheet for each article
df.to_excel(str(article) +'_Factcheck1.xlsx', index=False)
