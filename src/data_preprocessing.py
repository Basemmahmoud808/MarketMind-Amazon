"""
Amazon Product Reviews Preprocessing Pipeline
Cleaning text, handling nulls, and preparing sentiment labels.
"""

import re
import string
import pandas as pd
import numpy as np

# Standard common English stopwords list (lightweight and self-contained)
STOPWORDS = set([
    "a", "about", "above", "after", "again", "against", "all", "am", "an", "and", "any", "are", 
    "as", "at", "be", "because", "been", "before", "being", "below", "between", "both", "but", 
    "by", "could", "did", "do", "does", "doing", "down", "during", "each", "few", "for", "from", 
    "further", "had", "has", "have", "having", "he", "her", "here", "hers", "herself", "him", 
    "himself", "his", "how", "i", "if", "in", "into", "is", "it", "its", "itself", "just", 
    "me", "more", "most", "my", "myself", "no", "nor", "not", "of", "off", "on", "once", "only", 
    "or", "other", "ought", "our", "ours", "ourselves", "out", "over", "own", "same", "she", 
    "should", "so", "some", "such", "than", "that", "the", "their", "theirs", "them", "themselves", 
    "then", "there", "these", "they", "this", "those", "through", "to", "too", "under", "until", 
    "up", "very", "was", "we", "were", "what", "when", "where", "which", "while", "who", "whom", 
    "why", "with", "would", "you", "your", "yours", "yourself", "yourselves"
])

# Keep negation words that strongly affect sentiment
NEGATIONS = {"no", "not", "nor", "neither", "never"}
STOPWORDS = STOPWORDS - NEGATIONS

def clean_text(text: str) -> str:
    """
    Cleans raw review text:
    - Lowercase conversion
    - Stripping HTML tags and URLs
    - Removing special characters, punctuation, and digits
    - Removing non-essential stopwords while preserving sentiment signals
    """
    if not isinstance(text, str):
        return ""
    
    # 1. Lowercase
    text = text.lower()
    
    # 2. Remove HTML tags
    text = re.sub(r'<.*?>', ' ', text)
    
    # 3. Remove URLs
    text = re.sub(r'http\S+|www\S+|https\S+', ' ', text, flags=re.MULTILINE)
    
    # 4. Remove contractions expansions (basic)
    text = re.sub(r"won't", "will not", text)
    text = re.sub(r"can\'t", "can not", text)
    text = re.sub(r"n\'t", " not", text)
    
    # 5. Remove punctuation and numbers
    text = re.sub(r'[^a-zA-Z\s]', ' ', text)
    
    # 6. Tokenize & filter stopwords
    tokens = text.split()
    tokens = [t for t in tokens if len(t) > 2 and (t not in STOPWORDS or t in NEGATIONS)]
    
    return " ".join(tokens)

def load_and_preprocess(raw_data_path: str, sample_size: int = 25000) -> pd.DataFrame:
    """
    Loads raw Amazon reviews, creates binary sentiment target,
    handles missing values, cleans text, and balances the classes.
    """
    print(f"Loading data from {raw_data_path}...")
    df = pd.read_csv(raw_data_path)
    
    # Handle missing values
    df['Text'] = df['Text'].fillna('')
    df['Summary'] = df['Summary'].fillna('')
    
    # Combine Summary and Text for richer context
    df['Combined_Text'] = df['Summary'] + ' ' + df['Text']
    df = df[df['Combined_Text'].str.strip() != ''].copy()
    
    # Define Sentiment Target:
    # 4 & 5 stars -> 1 (Positive)
    # 1 & 2 stars -> 0 (Negative)
    # 3 stars are neutral / borderline, excluded for sharp binary classification
    df = df[df['Score'] != 3].copy()
    df['Sentiment'] = (df['Score'] > 3).astype(int)
    
    # Class balancing (sample to avoid massive positive imbalance)
    pos_count = (df['Sentiment'] == 1).sum()
    neg_count = (df['Sentiment'] == 0).sum()
    print(f"Original distribution: {pos_count} Positive, {neg_count} Negative")
    
    # Sample balanced dataset
    n_each = min(sample_size // 2, neg_count, pos_count)
    df_pos = df[df['Sentiment'] == 1].sample(n=n_each, random_state=42)
    df_neg = df[df['Sentiment'] == 0].sample(n=n_each, random_state=42)
    df_balanced = pd.concat([df_pos, df_neg]).sample(frac=1.0, random_state=42).reset_index(drop=True)
    
    print(f"Cleaning {len(df_balanced)} balanced reviews...")
    df_balanced['Cleaned_Text'] = df_balanced['Combined_Text'].apply(clean_text)
    
    # Remove empty cleaned texts
    df_balanced = df_balanced[df_balanced['Cleaned_Text'].str.strip() != ''].copy()
    
    print(f"Preprocessed dataset ready with {len(df_balanced)} rows.")
    return df_balanced

if __name__ == "__main__":
    df = load_and_preprocess("D:/project_model/data/raw/amazon_reviews_raw.csv")
    df.to_csv("D:/project_model/data/processed/cleaned_reviews.csv", index=False)
    print("Saved to D:/project_model/data/processed/cleaned_reviews.csv")
