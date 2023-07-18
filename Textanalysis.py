import pandas as pd
import re
import os
import requests
from bs4 import BeautifulSoup
import nltk
from nltk.tokenize import sent_tokenize
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords

textdir = 'text_files'
os.makedirs(textdir, exist_ok=True)
stopwordsdir = 'Stopwords'
masterdicdir = 'MasterDictionary'

input = pd.read_excel("Input.xlsx")
output = pd.read_excel("Output Data Structure.xlsx")

urls = input["URL"]
url_id = input["URL_ID"]
error = []
for i in range (len(urls)):
  titles= []
  req = requests.get(urls[i])
  request = req.text
  soup = BeautifulSoup(req.content, 'html.parser')
  title = soup.title.get_text()
  title_clean = title[0:title.index( " | Blackcoffer" )]
  print(title_clean)
  if title_clean == "Page not found":
    error.append(url_id[i])
    print(error)
  titles.append(title_clean)
  artic= soup.find_all(attrs= {'class' : "td-post-content" })
  for articl in artic:
    article = articl.text  
    article = article.replace('\n', '')
  file_path = os.path.join(textdir, f'{url_id[i]}.txt')
  with open(file_path, 'w', encoding='utf-8') as file:
        file.write(title_clean + '\n\n')
        file.write(article)
  print(f"Article saved: {url_id[i]}.txt")

stop_words = set()
for files in os.listdir(stopwordsdir):
  with open(os.path.join(stopwordsdir,files),'r',encoding='ISO-8859-1') as f:
    stop_words.update(set(word.lower() for word in f.read().splitlines()))

pos=set()
neg=set()

for files in os.listdir(masterdicdir):
  if files =='positive-words.txt':
    with open(os.path.join(masterdicdir,files),'r',encoding='ISO-8859-1') as f:
      pos.update(f.read().splitlines())
  else:
    with open(os.path.join(masterdicdir,files),'r',encoding='ISO-8859-1') as f:
      neg.update(f.read().splitlines())


docs = []
for text_file in os.listdir(textdir):
  with open(os.path.join(textdir,text_file),'r',encoding='ISO-8859-1') as f:
    text = f.read()
    words = word_tokenize(text)
    filtered_text = [word for word in words if word.lower() not in stop_words]
    docs.append(filtered_text)


pos_words = []
Neg_words =[]
pos_score = []
neg_score = []
pol_score = []
sub_score = []

for i in range(len(docs)):
  pos_words.append([word for word in docs[i] if word.lower() in pos])
  Neg_words.append([word for word in docs[i] if word.lower() in neg])
  pos_score.append(len(pos_words[i]))
  neg_score.append(len(Neg_words[i]))
  pol_score.append((pos_score[i] - neg_score[i]) / ((pos_score[i] + neg_score[i]) + 0.000001))
  sub_score.append((pos_score[i] + neg_score[i]) / ((len(docs[i])) + 0.000001))

word_count = []
avg_words_p_sent = []
avg_wordlength = []
avg_senlength = []
syllable_per_word = []
comp_word_count = []
compwords_percent = []
fog_index = []
personal_pronouns= []

stopwords = set(stopwords.words('english'))
def values(file):
   with open(os.path.join(textdir,file), 'r', encoding='ISO-8859-1') as f:
     text = f.read()
     text = re.sub(r'[^\w\s.]','',text)
     sentences = text.split('.')
     numsen = len(sentences)
     words = [word  for word in text.split() if word.lower() not in stopwords]
     length = sum(len(word) for word in words)
     numwords = len(words)
     avg_wordspersent = numwords / numsen
     avg_word_length = length / numwords
     avg_senlength = avg_wordspersent
     syllable_words = []
     complex_words = []
     for word in words:
       if word.endswith('es'):
         word = word[:-2]
       elif word.endswith('ed'):
          word = word[:-2]
       vowels = 'aeiou'
       syllable_count_word = sum( 1 for letter in word if letter.lower() in vowels)
       if syllable_count_word >= 1:
         syllable_words.append(word)         
       if syllable_count_word > 2:
        complex_words.append(word)  
       syllableperword = syllable_count_word / numwords
       compwordlen = len(complex_words)
       compwordpercent = (compwordlen * 100) / numwords 
       fogind = 0.4 * (avg_senlength + compwordpercent)
       personal_pronouns = ["I", "we", "We", "my", "My", "ours", "Ours", "us"]
       proncount = 0
       for pronoun in personal_pronouns:
         proncount += len(re.findall(r"\b" + pronoun + r"\b", text))



     return numwords, avg_wordspersent, avg_word_length, avg_senlength, syllableperword, compwordlen, compwordpercent, fogind, proncount

  
for file in os.listdir(textdir):
  a, b, c, d, e, f, g, h, i = values(file)
  word_count.append(a)
  avg_words_p_sent.append(b)
  avg_wordlength.append(c)
  avg_senlength.append(d)
  syllable_per_word.append(e)
  comp_word_count.append(f)
  compwords_percent.append(g)
  fog_index.append(h)
  personal_pronouns.append(i)


scores = [pos_score,
          neg_score,
          pol_score,
          sub_score,
          avg_senlength,
          compwords_percent,
          fog_index,
          avg_senlength,
          comp_word_count,
          word_count,
          syllable_per_word,
          personal_pronouns,
          avg_wordlength]

for i, var in enumerate(scores):
  print(i, var)
  output.iloc[:,i+2] = var

for i in error: 
  output.drop([i-35], axis= 0, inplace= True )

output.to_excel('output_final.xlsx', index=False)