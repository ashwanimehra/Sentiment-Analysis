import pandas as pd
import nltk
import string
from nltk.corpus import stopwords
nltk.download('punkt')
from textblob import TextBlob
nltk.download('wordnet')
from nltk.stem import WordNetLemmatizer 
from collections import Counter


def sentiment(x):
    sentiment = TextBlob(x)
    return sentiment.sentiment.polarity

def remove_case(ip):
    return (ip.lower())
  
def remove_allstop(ip):
    ip_tokens = nltk.word_tokenize(ip)
    stop = set(stopwords.words('english'))
    all_stops = stop | set(string.punctuation)
    text_nostop = [t for t in ip_tokens if t not in all_stops]
    return text_nostop

def lem(ip):
    lema = WordNetLemmatizer()
    lem_text = [lema.lemmatize(t) for t in ip ]
    return lem_text

def top_words(ip):
    c = Counter(ip)
    return (c.most_common(n=10))

def make_sentence(ip):
    sen = ' '.join(ip)
    return sen

def senti(ip):
    if ip > 0:
        return('positive')
    elif ip < 0:
        return('negative')
    elif ip == 0:
        return('neutral')

if __name__ == '__main__':
	
	data = pd.read_excel('ReviewData.xlsx')

	data['Review_Paragraph'] = data['Review_Paragraph'].apply(remove_case)
	data.head()
	data['Review_Paragraph'] = data['Review_Paragraph'].apply(remove_allstop)
	data.head()
	data['Review_Paragraph'] = data['Review_Paragraph'].apply(lem)
	data.head()

	#for top 10 words
	cor = []
	for d in data['Review_Paragraph']:
	    cor.extend(d)
	tw = top_words(cor)
	print('Top 10 words are:')
	for item in tw:
	    print(item[0])

	#Sentiment Analysis
	data['sentiment'] = data['Review_Paragraph'].apply(sentiment)
	data['senti'] = data['sentiment'].apply(senti)
	data.head()

	#Print Most Positive, Most Negative and Neutral review based on the score
	print('Most Positive : {}'.format(data[data['sentiment']==data['sentiment'].max()].Review_Title.values[0]))

	print('Most Negative : {}'.format(data[data['sentiment']==data['sentiment'].min()].Review_Title.values[0]))

	print('Neutral : {}'.format(data[data['sentiment']==0].Review_Title.values[0]))
	data['Review_Paragraph'] = data['Review_Paragraph'].apply(make_sentence)