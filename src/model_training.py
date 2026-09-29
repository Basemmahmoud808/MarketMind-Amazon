"""
Build Comprehensive Bilingual (Arabic + English) Dataset and Train ML Models
"""
import os
import json
import joblib
import random
import pandas as pd
import numpy as np
import re

from sklearn.model_selection import train_test_split, cross_val_score
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.naive_bayes import MultinomialNB
from sklearn.linear_model import LogisticRegression
from sklearn.svm import LinearSVC
from sklearn.calibration import CalibratedClassifierCV
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import (
    accuracy_score, precision_score, recall_score, f1_score,
    roc_auc_score, confusion_matrix
)

DATA_RAW_PATH = "D:/project_model/data/raw/amazon_reviews_raw.csv"
MODEL_DIR = "D:/project_model/models"

# 1. Arabic Seed Templates & Variations
arabic_positive_templates = [
    "المنتج ممتاز وعالي الجودة وشغال بكفاءة عالية جدا",
    "تحفة فنية وخامات محترمة جدا وسرعة في التوصيل",
    "انصح به بشدة يستحق كل قرش اندفع فيه",
    "هاتف رائع جدا والبطارية بتكفي يومين استخدام شاق",
    "قيمة ممتازة مقابل السعر وأداء مستقر وسريع جدا",
    "أصلي تماما ومطابق للمواصفات ومريح في الاستخدام",
    "كاميرا خرافية وشاشة واضحة وسريعة جدا 120 هرتز",
    "تجربة استخدام رائعة وسعيد جدا بالشراء من امازون",
    "افضل اختيار في فئته السعرية ولا يوجد اي عيوب",
    "روعة وتغليف ممتاز والشحن وصل في المعاد مضبوط",
    "جهاز جبار وصوته نقي وسريع جدا في الالعاب",
    "الخامات ممتازة وناعمة ولا تترك بصمات ومريحة",
    "شكرا للبائع المنتج وصل بحالة ممتازة وبدون خدش",
    "مفيد جدا وعملي ومطابق للوصف بالظبط وبجودة فاخرة",
    "شحن سريع جدا وبطارية خارقة وتصميم شيك جدا",
    "شغال تمام وممتاز ومفيش اي مشاكل فنية خالص",
    "منتج محترم جدا واصلي ويستاهل 5 نجوم عن جدارة",
    "احسن هاتف اشتريته هذا العام خفيف وسريع وانيق",
    "جودة التصنيع عالية جدا واداء المعالج فوق الممتاز",
    "توصيل محترم ومندوب محترم والمنتج تحفة",
    "مظبوط تماما وجودة فائقة وتغليف محكم للغاية",
    "ممتاز وسعره ارخص من المحلات بكتير وأنصح به",
    "شاشة جميلة وألوان مبهجة والسطوع ممتاز تحت الشمس",
    "المنتج اصلي مية بالمية وضمان معتمد وكل شيء تمام",
    "مبسوط جدا منه وبستخدمه بقالي اسبوعين ممتاز جدا",
    "بصراحة تفوق على توقعاتي أداء وكاميرا وشحن سريع",
    "خدمة عملاء ممتازة والمنتج وصل سليم وبكل ملحقاته",
    "جيد جدا ومناسب للاستخدام اليومي وبسعر اقتصادي",
    "من أفضل المنتجات اللي اشتريتها اونلاين جودة وفخامة",
    "رائع جدا وسلس وسهل التعامل وخفيف الوزن",
    "الماوس باد خامتها ناعمة جدا وممتازة ومش بتزحلق",
    "بادة ماوس ممتازة مقاسها مناسب وخياطة الحواف متينة",
    "جراب حماية ممتاز ومقاسه مضبوط جدا على التليفون",
    "سماعة صوتها نقي والعزل ممتاز ومايكروفون واضح",
    "شاحن سريع اصلي وما بيسخنش الموبايل خالص وممتاز"
]

arabic_negative_templates = [
    "المنتج سيء جدا وخامات رديئة ولا يعمل بشكل صحيح",
    "تجربة سيئة للغاية ومقلد وليس اصلي ولا انصح به",
    "خسارة الفلوس اللي ادفعت فيه وصل مكسور وغير مطابق",
    "رديء جدا ويسخن بطريقة غريبة والبطارية بتفضى بسرعة",
    "للاسف وصل تالف والكرتونة مقطوعة والشاحن لا يعمل",
    "بايظ ومش شغال نهائيا وبطلب استرجاع والموقع لا يستجيب",
    "نصابين حسبي الله ونعم الوكيل منتج رديء ومغشوش",
    "لا انصح بشرائه اطلاقا خامات زبالة ومكسور من الداخل",
    "غالي جدا على الفاضي ومفيش فيه اي ميزة والجودة رديئة",
    "بيهنج كتير والكاميرا سيئة جدا ومش واضحة خالص",
    "وصل مختلف تماما عن الصور والمواصفات المعروضة",
    "اسوأ تجربة شراء اونلاين والمنتج غير صالح للاستخدام",
    "صوت السماعة واطي جدا ومكتوم وبيقطع على طول",
    "الشاشة بهتانة ومكسورة واللمس لا يستجيب",
    "الجهاز مش بيشحن خالص ومطابق للمنتجات المقلدة",
    "تغليف مهمل جدا والبضاعة وصلت متبهدلة ومخدوشة",
    "توقف عن العمل بعد يومين فقط وضاعت الفلوس على الفاضي",
    "جودة ضعيفة جدا خفيف كأنه لعبة اطفال مش منتج حقيقي",
    "مندوب الشحن اتأخر جدا والمنتج فيه عيب صناعة واضح",
    "مقلد وفيك وبايظ وحجمه غير المكتوب تماما ومخيب للامل",
    "خدمة ما بعد البيع سيئة ورفضوا يرجعوا المنتج المعيوب",
    "بيسخن نار اثناء الشحن والبطارية اتنفخت وخطر جدا",
    "الماوس باد ريحتها وحشة جدا وبتزحلق وخامتها رديئة",
    "الخياطة مقطوعة ومنسولة ومش كويسة خالص",
    "سعر مبالغ فيه على جودة زبالة ميكملش اسبوع ويبوظ",
    "منتج تجاري رخيص ومضروب ومفيش معاه ضمان",
    "خيبة امل كبيرة وميستاهلش حتى نجمة واحدة",
    "غير مطابق للوصف تماما مقاس صغير وخامة سيئة",
    "وصل مستعمل وفيه بصمات وتراب وعلبة مفتوحة",
    "لا يعمل وتالف وضياع وقت وفلوس حسبي الله"
]

# Multilingual seeds (Italian, German, French)
multilingual_positive = [
    "Ottimo sotto tutti punti di vista prodotto eccellente e consigliato",
    "Starkes Smartphone mit großem Display sehr gute Akkulaufzeit",
    "Super Qualität schnelle Lieferung bin vollkommen zufrieden",
    "Parfait envoi rapide et sécurisé produit original de grande qualité",
    "Sehr gut verarbeitet funktioniert einwandfrei klare Kaufempfehlung",
    "Tres bon produit conforme a la description et performant",
    "Eccellente acquisto qualita ottima e spedizione velocissima",
    "Top Produkt super Akku und schones Design gerne wieder"
]

multilingual_negative = [
    "Pessimo prodotto rotto e non funzionante pessima qualita",
    "Sehr schlechte Qualität defekt angekommen nicht zu empfehlen",
    "Decevant tres lent et bloque tout le temps produit defectueux",
    "Schlechtes Produkt nach zwei Tagen kaputt geldverschwendung",
    "Horrible et abime ne fonctionne pas du tout a eviter"
]

def build_bilingual_corpus():
    print("Loading English raw Amazon reviews...")
    df_raw = pd.read_csv(DATA_RAW_PATH)
    df_raw['Text'] = df_raw['Text'].fillna('')
    df_raw['Summary'] = df_raw['Summary'].fillna('')
    df_raw['Combined'] = df_raw['Summary'] + ' ' + df_raw['Text']
    
    # Filter 4 & 5 stars (Pos), 1 & 2 stars (Neg)
    df_pos_en = df_raw[df_raw['Score'] >= 4].sample(n=8000, random_state=42)[['Combined']].copy()
    df_pos_en['Sentiment'] = 1
    
    df_neg_en = df_raw[df_raw['Score'] <= 2].sample(n=min(6500, len(df_raw[df_raw['Score'] <= 2])), random_state=42)[['Combined']].copy()
    df_neg_en['Sentiment'] = 0
    
    # Generate Arabic augmented corpus (5,000 samples)
    print("Generating Arabic and Multilingual NLP training corpus...")
    ar_pos = []
    adverbs_pos = ["بصراحة", "فعلا", "جدا", "للغاية", "والله", "بدون تردد", "تماماً", "عن تجربة", "بأمانة"]
    for _ in range(3500):
        t = random.choice(arabic_positive_templates)
        adv = random.choice(adverbs_pos)
        if random.random() > 0.5:
            text = f"{adv} {t}"
        else:
            text = f"{t} {adv}"
        ar_pos.append(text)
        
    ar_neg = []
    adverbs_neg = ["للأسف", "بصراحة", "جدا", "للغاية", "والله", "عن تجربة", "للأسف الشديد", "بجد"]
    for _ in range(3500):
        t = random.choice(arabic_negative_templates)
        adv = random.choice(adverbs_neg)
        if random.random() > 0.5:
            text = f"{adv} {t}"
        else:
            text = f"{t} {adv}"
        ar_neg.append(text)
        
    # Add multilingual samples
    multi_pos = multilingual_positive * 50
    multi_neg = multilingual_negative * 50
    
    df_pos_ar = pd.DataFrame({'Combined': ar_pos + multi_pos, 'Sentiment': 1})
    df_neg_ar = pd.DataFrame({'Combined': ar_neg + multi_neg, 'Sentiment': 0})
    
    df_all = pd.concat([df_pos_en, df_neg_en, df_pos_ar, df_neg_ar]).sample(frac=1.0, random_state=42).reset_index(drop=True)
    print(f"Total Bilingual Corpus: {len(df_all)} (Positive: {(df_all['Sentiment']==1).sum()}, Negative: {(df_all['Sentiment']==0).sum()})")
    return df_all

def train_and_save():
    df = build_bilingual_corpus()
    
    X = df['Combined']
    y = df['Sentiment']
    
    X_train, X_test, y_train, y_test = train_test_split(
        X, y, test_size=0.20, random_state=42, stratify=y
    )
    print(f"Train samples: {len(X_train)}, Test samples: {len(X_test)}")
    
    print("Fitting Bilingual TF-IDF Vectorizer (Unicode + N-Grams)...")
    vectorizer = TfidfVectorizer(
        token_pattern=r'(?u)\b\w+\b',
        ngram_range=(1, 2),
        max_features=12000,
        sublinear_tf=True
    )
    X_train_vec = vectorizer.fit_transform(X_train)
    X_test_vec = vectorizer.transform(X_test)
    
    print("Training Candidate Models with 5-Fold Cross Validation...")
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
        print(f"  Evaluating {name}...")
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
        print(f"    Acc: {acc*100:.2f}% | F1: {f1*100:.2f}% | 5-Fold CV: {cv_mean*100:.2f}%")
        
        if f1 > best_f1:
            best_f1 = f1
            best_model_name = name
            
    print(f"\nWinning Model: {best_model_name} (F1: {best_f1*100:.2f}%)")
    final_model = trained_objs[best_model_name]
    
    # Save Model & Vectorizer
    model_path = os.path.join(MODEL_DIR, "sentiment_model.pkl")
    vec_path = os.path.join(MODEL_DIR, "tfidf_vectorizer.pkl")
    meta_path = os.path.join(MODEL_DIR, "model_comparison.json")
    
    joblib.dump(final_model, model_path)
    joblib.dump(vectorizer, vec_path)
    
    test_acc = results[best_model_name]["accuracy"]
    test_f1 = results[best_model_name]["f1_score"]
    
    metadata = {
        "best_model": best_model_name,
        "final_accuracy": test_acc,
        "final_f1": test_f1,
        "final_confusion_matrix": results[best_model_name]["confusion_matrix"],
        "all_models_comparison": results
    }
    
    with open(meta_path, "w", encoding="utf-8") as f:
        json.dump(metadata, f, indent=4)
        
    print(f"Model saved to {model_path}")
    print(f"Vectorizer saved to {vec_path}")
    print("Training finished successfully!")

if __name__ == "__main__":
    train_and_save()
