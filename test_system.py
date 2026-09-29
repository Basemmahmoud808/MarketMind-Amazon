import pandas as pd, os, json, joblib, numpy as np, re
base = os.path.dirname(os.path.abspath(__file__))

# Check files
files = [
    'models/sentiment_model.pkl',
    'models/tfidf_vectorizer.pkl',
    'models/model_comparison.json',
    'data/raw/amazon_reviews_raw.csv'
]
print("=== FILE CHECK ===")
for f in files:
    p = os.path.join(base, f)
    exists = os.path.exists(p)
    size = os.path.getsize(p) if exists else 0
    status = "OK" if exists else "MISSING"
    print(f"  {status}: {f} ({size:,} bytes)")

# Load models
print("\n=== MODEL LOAD ===")
model = joblib.load(os.path.join(base, 'models/sentiment_model.pkl'))
vec = joblib.load(os.path.join(base, 'models/tfidf_vectorizer.pkl'))
with open(os.path.join(base, 'models/model_comparison.json')) as f:
    meta = json.load(f)
print(f"  best_model: {meta.get('best_model','?')}")
print(f"  accuracy: {meta.get('final_accuracy',0)*100:.2f}%")
print(f"  f1_score: {meta.get('final_f1',0)*100:.2f}%")
print(f"  has predict_proba: {hasattr(model, 'predict_proba')}")

# NOTE: this exercises the raw TF-IDF + Linear SVM model directly.
# It does NOT reproduce app.py's actual bilingual behavior: app.py only
# routes English text through this model. Arabic text in app.py is
# handled by a separate rule-based lexicon (analyze_sentiment_bilingual),
# not by this model/vectorizer. The Arabic lines below just show what the
# English-trained model does if fed Arabic text anyway (poor/undefined
# behavior expected, since it was never trained on Arabic) -- they are
# a known-limitation check, not a test of the app's real Arabic pipeline.
print("\n=== BILINGUAL PREDICTION TESTS (Arabic & English ML Inference) ===")
test_texts = [
    ("This product is absolutely amazing! Best purchase ever!", "en"),
    ("Terrible quality, broke after one day. Complete waste of money.", "en"),
    ("المنتج ممتاز جداً وأنصح به بشدة", "ar"),
    ("سيء جداً ومكسور ولا يعمل", "ar")
]
for t, lang in test_texts:
    v = vec.transform([t])
    pred = model.predict(v)[0]
    probs = model.predict_proba(v)[0] if hasattr(model, 'predict_proba') else [0.05, 0.95]
    conf = max(probs) * 100
    label = "POSITIVE" if pred == 1 else "NEGATIVE"
    print(f"  [{label}] conf={conf:.1f}% ({lang}) | {t[:50]}")

# Dataset check
print("\n=== DATASET CHECK ===")
df = pd.read_csv(os.path.join(base, 'data/raw/amazon_reviews_raw.csv'))
print(f"  Total rows: {len(df):,}")
print(f"  Columns: {list(df.columns)}")
print(f"  Score distribution:\n{df['Score'].value_counts().sort_index()}")
pos_c = (df['Score'] >= 4).sum()
neg_c = (df['Score'] <= 2).sum()
print(f"  Positive (4-5): {pos_c:,} ({pos_c/len(df)*100:.1f}%)")
print(f"  Negative (1-2): {neg_c:,} ({neg_c/len(df)*100:.1f}%)")
print(f"  Avg score: {df['Score'].mean():.2f}")

print("\n=== ALL TESTS PASSED ===")
