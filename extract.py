import customtkinter as ctk
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
import numpy as np

from konlpy.tag import Okt
from tkinter import filedialog

import os
from glob import glob
import pandas as pd
import csv


# 리뷰 파일 로드
def load_excel(input_path:str):
    file_list = glob(os.path.join(input_path,"*.xlsx"))
    return pd.read_excel(file_list[0])


from kiwipiepy import Kiwi
kiwi = Kiwi()

# 형태소 분석(명사 추출)
def tokenize(text:str):
    # from konlpy.tag import Okt
    # okt = Okt()
    # pos = okt.pos(text, stem=True)
    # tokens = [word for word, tag in pos if tag in ['Noun', 'Adjective']]
    tokens = [t.form for t in kiwi.tokenize(text) if t.tag in ["NNG", "NNP", "VA"]]
    return tokens

# 불용어 제거
def clean_text(nouns_text:list) -> list():
    # 불용어 파일 로드
    with open('stop_words_kr.txt', encoding='utf-8') as f:
        stop_words = f.read()
        stop_words = set(w.strip() for w in stop_words.split('\n') if w.strip())

    # 불용어를 제외하고 반환
    return [ noun for noun in nouns_text if noun not in stop_words]

def preprocess(input_file:pd.DataFrame) -> np.ndarray:
    nouns_text = input_file['REVIEW'].apply(lambda text: tokenize(text))
    removed_stopwords = nouns_text.apply(lambda nouns_text: clean_text(nouns_text))
    docs = removed_stopwords.apply(lambda word: ' '.join(word))
    return docs

def extract_top_keywords(corpus, top_n=10):
    """
    corpus : list[str] 또는 pandas.Series[str]
        문서 전체 (문서별 하나의 문자열)
    top_n : int
        상위 키워드 개수
    """
    vectorizer = TfidfVectorizer()
    tfidf = vectorizer.fit_transform(corpus)
    feature_names = vectorizer.get_feature_names_out()
    
    # 전체 문서 기준 TF-IDF 합산
    scores = tfidf.toarray().sum(axis=0)
    top_indices = scores.argsort()[::-1][:top_n]

    top_keywords = [(feature_names[i], scores[i]) for i in top_indices]
    top_keywords = np.array(top_keywords)
    return top_keywords[:, 0], top_keywords[:,1]

# 저장
def save_to_csv(keywords, scores:None, filename="result.csv"):
    with open(filename, "w", newline="", encoding='utf-8-sig') as f:
        writer = csv.writer(f, delimiter=',')
        writer.writerow(["키워드", "중요도"])
        if scores:
            for word, score in zip(keywords, scores):
                writer.writerow([word, f"{score:.4f}"])
        else:
            for word in keywords:
                writer.writerow([word])
                

def main():
    input_path="input_text" # 후기 데이터 입력 폴더경로
    input_file = load_excel(input_path)
    print(input_file.head())
    
    docs = preprocess(input_file)
    keywords, scores = extract_top_keywords(docs)
    scores = list(map(float, scores))
    keywords = list(map(str, keywords))

    save_to_csv(keywords, scores)
    print("Done!") 
    
if __name__ == "__main__":
    main()