from textblob import TextBlob
import xlrd
import spacy
import docx2txt
from stanza.server import CoreNLPClient
import re
import nltk
import stanza
from pycorenlp import *

nlp = StanfordCoreNLP("http://localhost:9000/")

# get verb phrases

def get_verb_phrases(t):
    verb_phrases = []
    num_children = len(t)
    num_VP = sum(1 if t[i].label() == "VP" else 0 for i in range(0, num_children))

    if t.label() != "VP":
        for i in range(0, num_children):
            if t[i].height() > 2:
                verb_phrases.extend(get_verb_phrases(t[i]))
    elif t.label() == "VP" and num_VP > 1:
        for i in range(0, num_children):
            if t[i].label() == "VP":
                if t[i].height() > 2:
                    verb_phrases.extend(get_verb_phrases(t[i]))
    else:
        verb_phrases.append(' '.join(t.leaves()))

    return verb_phrases


# get position of first node "VP" while traversing from top to bottom

def get_pos(t):
    vp_pos = []
    sub_conj_pos = []
    num_children = len(t)
    children = [t[i].label() for i in range(0,num_children)]

    flag = re.search(r"(S|SBAR|SBARQ|SINV|SQ)", ' '.join(children))

    if "VP" in children and not flag:
        for i in range(0, num_children):
            if t[i].label() == "VP":
                vp_pos.append(t[i].treeposition())
    elif not "VP" in children and not flag:
        for i in range(0, num_children):
            if t[i].height() > 2:
                temp1,temp2 = get_pos(t[i])
                vp_pos.extend(temp1)
                sub_conj_pos.extend(temp2)
    # comment this "else" part, if want to include subordinating conjunctions
    else:
        for i in range(0, num_children):
            if t[i].label() in ["S","SBAR","SBARQ","SINV","SQ"]:
                temp1, temp2 = get_pos(t[i])
                vp_pos.extend(temp1)
                sub_conj_pos.extend(temp2)
            else:
                sub_conj_pos.append(t[i].treeposition())

    return (vp_pos,sub_conj_pos)


# get all clauses
def get_clause_list(sent):
    parser = nlp.annotate(sent, properties={"annotators":"parse","outputFormat": "json"})
    sent_tree = nltk.tree.ParentedTree.fromstring(parser["sentences"][0]["parse"])
    clause_level_list = ["S","SBAR","SBARQ","SINV","SQ"]
    clause_list = []
    sub_trees = []
    # sent_tree.pretty_print()

    # break the tree into subtrees of clauses using
    # clause levels "S","SBAR","SBARQ","SINV","SQ"
    for sub_tree in reversed(list(sent_tree.subtrees())):
        if sub_tree.label() in clause_level_list:
            if sub_tree.parent().label() in clause_level_list:
                continue

            if (len(sub_tree) == 1 and sub_tree.label() == "S" and sub_tree[0].label() == "VP"
                and not sub_tree.parent().label() in clause_level_list):
                continue

            sub_trees.append(sub_tree)
            del sent_tree[sub_tree.treeposition()]

    # for each clause level subtree, extract relevant simple sentence
    for t in sub_trees:
        # get verb phrases from the new modified tree
        verb_phrases = get_verb_phrases(t)

        # get tree without verb phrases (mainly subject)
        # remove subordinating conjunctions
        vp_pos,sub_conj_pos = get_pos(t)
        for i in vp_pos:
            del t[i]
        for i in sub_conj_pos:
            del t[i]

        subject_phrase = ' '.join(t.leaves())

        # update the clause_list
        for i in verb_phrases:
            clause_list.append(subject_phrase + " " + i)

    return clause_list

# excel file of factcheck done by IBR researchers
xlfile = r"C:\Users\liamh\Downloads\Test_1.xlsx"
wb = xlrd.open_workbook(xlfile)

# reading the excel file
sheet = wb.sheet_by_index(0)
row_count = sheet.nrows
humanfacts = []

# creating a list to store the sentences selected as facts by IBR researchers
for curr_row in range(1,row_count):
    fact = sheet.cell_value(curr_row, 1)
    fact2 = re.sub(r"\W","",fact)
    humanfacts.append(fact2)

# opening an example of an IBR article to parse through looking for facts
text = docx2txt.process("11. Modern Health - Draft 4.docx")
nlpfacts = []

processed = TextBlob(text)

# parsing through the IBR article to look for facts
for line in processed.sentences:
    if line != '':
        sent = re.sub(r"(\.|,|\?|\(|\)|\[|\])", " ", line.string)
        if sent != '' and sent != ' 'and len(sent)!=1:
            clauses = get_clause_list(sent)
            for clause in clauses:
                clause = TextBlob(clause)
                
                # major decision to tell if a sentence contains a fact or not, arbitrarily chose 0.5 on a scale of 0:1 to test
                if clause.subjectivity < 0.5:
                    line2 = re.sub(r"\W","",line.string)
                    nlpfacts.append(line2)

numbercorrect = 0
totalfacts = len(humanfacts)
both = {}

# testing how many of the facts picked up through the nlp were also selected by IBR researchers
for autof in nlpfacts:
    for humanf in humanfacts:
        if autof in humanf:
            numbercorrect += 1
            humanfacts.remove(humanf)
            both[autof] = humanf
            break


if totalfacts != 0:
    accuracy = numbercorrect/totalfacts
else:
    accuracy = 0

print(humanfacts)
print(len(humanfacts))
print(len(nlpfacts))
print(accuracy)
print(both)
