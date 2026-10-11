"""
Train Machine Learning Models Exclusively on American Amazon Reviews Dataset
100% English Natural Language Processing Pipeline
"""
import os
import json
import re
import joblib
import pandas as pd
import numpy as np

from sklearn.model_selection import train_test_split, cross_val_score
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.naive_bayes import MultinomialNB
from sklearn.linear_model import LogisticRegression
from sklearn.svm import LinearSVC
from sklearn.calibration import CalibratedClassifierCV
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import (
    accuracy_score, precision_score, recall_score, f1_score, confusion_matrix
)

DATA_RAW_PATH = "D:/project_model/data/raw/amazon_reviews_raw.csv"
MODEL_DIR = "D:/project_model/models"

def clean_english_text(text: str) -> str:
    if not isinstance(text, str):
        return ""
    text = text.lower()
    text = re.sub(r'<.*?>', ' ', text)
    text = re.sub(r'http\S+|www\S+|https\S+', ' ', text)
    text = re.sub(r"won't", "will not", text)
    text = re.sub(r"can\'t", "can not", text)
    text = re.sub(r"n\'t", " not", text)
    text = re.sub(r'[^a-zA-Z\s]', ' ', text)
    return ' '.join(text.split())

def build_american_dataset():
    print(f"Loading American Amazon reviews from {DATA_RAW_PATH}...")
    df = pd.read_csv(DATA_RAW_PATH)
    
    df['Text'] = df['Text'].fillna('')
    df['Summary'] = df['Summary'].fillna('')
    df['Combined'] = df['Summary'] + ' ' + df['Text']
    df = df[df['Combined'].str.strip() != ''].copy()
    
    # 4 & 5 stars -> 1 (Positive), 1 & 2 stars -> 0 (Negative)
    # Neutral 3 stars excluded for distinct binary sentiment boundary
    df = df[df['Score'] != 3].copy()
    df['Sentiment'] = (df['Score'] > 3).astype(int)
    
    pos_df = df[df['Sentiment'] == 1]
    neg_df = df[df['Sentiment'] == 0]
    
    n_neg = len(neg_df)
    print(f"Available American reviews: {len(pos_df)} Positive, {n_neg} Negative")
    
    # Balance classes to prevent majority class bias (sample equal positive)
    pos_sample = pos_df.sample(n=n_neg, random_state=42)
    balanced_df = pd.concat([pos_sample, neg_df]).sample(frac=1.0, random_state=42).reset_index(drop=True)
    
    print(f"Preprocessing {len(balanced_df)} English reviews...")
    balanced_df['Cleaned'] = balanced_df['Combined'].apply(clean_english_text)
    balanced_df = balanced_df[balanced_df['Cleaned'].str.strip() != ''].copy()
    
    print(f"Final training corpus: {len(balanced_df)} reviews (50% Pos / 50% Neg)")
    return balanced_df

def train_and_evaluate():
    df = build_american_dataset()
    
    X = df['Cleaned']
    y = df['Sentiment']
    
    X_train, X_test, y_train, y_test = train_test_split(
        X, y, test_size=0.20, random_state=42, stratify=y
    )
    print(f"Train set: {len(X_train)} reviews | Test set: {len(X_test)} reviews")
    
    print("Fitting English TF-IDF Vectorizer (Unigrams + Bigrams, 12,000 features)...")
    vectorizer = TfidfVectorizer(
        ngram_range=(1, 2),
        max_features=12000,
        sublinear_tf=True,
        stop_words='english'
    )
    X_train_vec = vectorizer.fit_transform(X_train)
    X_test_vec = vectorizer.transform(X_test)
    
    print("Evaluating models with 5-Fold Stratified Cross Validation on American reviews...")
    models = {
        "Multinomial Naive Bayes": MultinomialNB(alpha=0.5),
        "Logistic Regression": LogisticRegression(max_iter=1000, C=2.0, random_state=42),
        "Linear SVM (Calibrated)": CalibratedClassifierCV(LinearSVC(C=1.0, random_state=42), cv=5),
        "Random Forest": RandomForestClassifier(n_estimators=100, max_depth=20, random_state=42, n_jobs=-1)
    }
    
    results = {}
    best_model_name = None
    best_f1 = -1
    trained_objs = {}
    
    for name, m in models.items():
        print(f"\n--- Training {name} ---")
        cv_scores = cross_val_score(m, X_train_vec, y_train, cv=5, scoring='f1', n_jobs=-1)
        cv_mean = float(np.mean(cv_scores))
        
        m.fit(X_train_vec, y_train)
        trained_objs[name] = m
        
        y_pred = m.predict(X_test_vec)
        acc = float(accuracy_score(y_test, y_pred))
        prec = float(precision_score(y_test, y_pred))
        rec = float(recall_score(y_test, y_pred))
        f1 = float(f1_score(y_test, y_pred))
        cm = confusion_matrix(y_test, y_pred).tolist()
        
        results[name] = {
            "accuracy": round(acc, 4),
            "precision": round(prec, 4),
            "recall": round(rec, 4),
            "f1_score": round(f1, 4),
            "cv_f1_mean": round(cv_mean, 4),
            "confusion_matrix": cm
        }
        print(f"Result for {name}:")
        print(f"  Accuracy:  {acc*100:.2f}%")
        print(f"  Precision: {prec*100:.2f}%")
        print(f"  Recall:    {rec*100:.2f}%")
        print(f"  F1-Score:  {f1*100:.2f}%")
        print(f"  5-Fold CV: {cv_mean*100:.2f}%")
        
        if f1 > best_f1:
            best_f1 = f1
            best_model_name = name
            
    print(f"\n==========================================")
    print(f"Winning Production Model: {best_model_name} (F1: {best_f1*100:.2f}%)")
    print(f"==========================================")
    
    final_model = trained_objs[best_model_name]
    
    model_path = os.path.join(MODEL_DIR, "sentiment_model.pkl")
    vec_path = os.path.join(MODEL_DIR, "tfidf_vectorizer.pkl")
    meta_path = os.path.join(MODEL_DIR, "model_comparison.json")
    
    joblib.dump(final_model, model_path)
    joblib.dump(vectorizer, vec_path)
    
    metadata = {
        "best_model": best_model_name,
        "final_accuracy": results[best_model_name]["accuracy"],
        "final_f1": results[best_model_name]["f1_score"],
        "final_confusion_matrix": results[best_model_name]["confusion_matrix"],
        "all_models_comparison": results
    }
    
    with open(meta_path, "w", encoding="utf-8") as f:
        json.dump(metadata, f, indent=4)
        
    print(f"Saved model binary: {model_path}")
    print(f"Saved vectorizer binary: {vec_path}")
    print(f"Saved benchmark metadata: {meta_path}")
    print("Pure American English model training completed successfully!")

if __name__ == "__main__":
    train_and_evaluate()
