from textblob import TextBlob
import xlrd
import docx2txt
import pandas as pd
import re
import Levenshtein as lev

# inputs
article = input("Article name: ")
iteration = int(input("Iteration #: "))
text = docx2txt.process(str(article) +"_"+str(iteration)+ ".docx")

# previous factcheck sheet
previous = str(article)+"_Factcheck"+str(iteration-1)+".xlsx"
wb = xlrd.open_workbook(previous)

sheet = wb.sheet_by_index(0)
row_count = sheet.nrows
previousfacts = {}

# getting all the facts from the previous sheet and storing the corresponding column items in a dict
for curr_row in range(1,row_count):
    fact = str(sheet.cell_value(curr_row, 1))
    if fact != '':
        previousfacts[fact]=[sheet.cell_value(curr_row, 2),sheet.cell_value(curr_row, 3), sheet.cell_value(curr_row, 4), sheet.cell_value(curr_row, 5), sheet.cell_value(curr_row, 6)]

# applying nlp library to the article
processed = TextBlob(text)
allfacts = {}


# parsing through the article to pull all facts
for line in processed.sentences:
    stop = False
    line =str(line)
    if not stop:
        
        # parsing through the previous facts
        for fact in previousfacts:
            
            # cool python library to test for word matching between two different strings, some heavy math i dont know how to do
            ratio = lev.ratio(fact, line)
            
            # setting the partial match threshold at 90% in order for the facts to from previous draft to be considered the same as the latest draft
            if ratio >0.8:
                allfacts[fact] = previousfacts[fact]
                stop = True
                break
    
    # if the fact is new
    if line != '' and not stop:
        sent = re.sub(r"(\.|,|\?|\(|\)|\[|\])", " ", line)
        if sent != '' and sent != ' 'and len(sent)!=1:
            sent = sent.replace('\n',' ')
            sent = TextBlob(sent)
            
            # subjectivity threshold for new facts at 80% subjective (this is very high but i'd rather overfit than underfit)
            if sent.subjectivity <=0.8:
                allfacts[str(sent)] = []


# creating pandas dataframe to store the sentences to be uploaded to factchecking spreadsheet                
df = pd.DataFrame.from_dict(allfacts, orient = 'index', columns = ["Colour", "Evidence", "Impact", "Source", "Link"])

# exporting sentences to pre-existing spreadsheet for each article
df.to_excel(str(article) +'_Factcheck' + str(iteration)+ ".xlsx")