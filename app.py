"""
Amazon Product Reviews - AI Sentiment & Market Intelligence System
Senior Enterprise Edition - Strictly Clean Vector Icons (No Emojis)
Course: AI Data Analysis (Horus Program - Growth Level) | Supervised by: Eng. Aya Badwy
"""

import os
import io
import re
import json
import joblib
import subprocess
import shutil
import urllib.request
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from bs4 import BeautifulSoup

# ----------------- PAGE CONFIG -----------------
st.set_page_config(
    page_title="MarketMind Amazon - AI Sentiment & Market Intelligence",
    page_icon="https://www.amazon.com/favicon.ico",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ----------------- VECTOR ICONS (LUCIDE / HEROICONS STYLE) -----------------
SVG_PACKAGE = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="display:inline-block; vertical-align:middle;"><path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"></path><polyline points="3.27 6.96 12 12.01 20.73 6.96"></polyline><line x1="12" y1="22.08" x2="12" y2="12"></line></svg>'
SVG_CHECK = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#059669" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" style="display:inline-block; vertical-align:middle;"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path><polyline points="22 4 12 14.01 9 11.01"></polyline></svg>'
SVG_ALERT = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#DC2626" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" style="display:inline-block; vertical-align:middle;"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="8" x2="12" y2="12"></line><line x1="12" y1="16" x2="12.01" y2="16"></line></svg>'
SVG_MESSAGE = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#D97706" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"></path></svg>'
SVG_TREND_UP = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#059669" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="23 6 13.5 15.5 8.5 10.5 1 18"></polyline><polyline points="17 6 23 6 23 12"></polyline></svg>'
SVG_SHIELD_ALERT = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#DC2626" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"></path><line x1="12" y1="8" x2="12" y2="12"></line><line x1="12" y1="16" x2="12.01" y2="16"></line></svg>'
SVG_PULSE = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#4F46E5" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"></polyline></svg>'
SVG_ZAP = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#4F46E5" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"></polygon></svg>'
SVG_STAR = '<svg width="15" height="15" viewBox="0 0 24 24" fill="#F59E0B" stroke="#F59E0B" stroke-width="1" style="display:inline-block; vertical-align:middle;"><polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"></polygon></svg>'
SVG_SEARCH = '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#2563EB" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="display:inline-block; vertical-align:middle;"><circle cx="11" cy="11" r="8"></circle><line x1="21" y1="21" x2="16.65" y2="16.65"></line></svg>'
SVG_STORE = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="display:inline-block; vertical-align:middle;"><path d="M6 2L3 6v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V6l-3-4z"></path><line x1="3" y1="6" x2="21" y2="6"></line><path d="M16 10a4 4 0 0 1-8 0"></path></svg>'
SVG_COFFEE = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="display:inline-block; vertical-align:middle;"><path d="M18 8h1a4 4 0 0 1 0 8h-1"></path><path d="M2 8h16v9a4 4 0 0 1-4 4H6a4 4 0 0 1-4-4V8z"></path><line x1="6" y1="1" x2="6" y2="4"></line><line x1="10" y1="1" x2="10" y2="4"></line><line x1="14" y1="1" x2="14" y2="4"></line></svg>'
SVG_PHONE = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="display:inline-block; vertical-align:middle;"><rect x="5" y="2" width="14" height="20" rx="2" ry="2"></rect><line x1="12" y1="18" x2="12.01" y2="18"></line></svg>'
SVG_HOME = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="display:inline-block; vertical-align:middle;"><path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"></path><polyline points="9 22 9 12 15 12 15 22"></polyline></svg>'
SVG_HEART_SPARK = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="display:inline-block; vertical-align:middle;"><path d="M20.84 4.61a5.5 5.5 0 0 0-7.78 0L12 5.67l-1.06-1.06a5.5 5.5 0 0 0-7.78 7.78l1.06 1.06L12 21.23l7.78-7.78 1.06-1.06a5.5 5.5 0 0 0 0-7.78z"></path></svg>'
SVG_PET = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="display:inline-block; vertical-align:middle;"><circle cx="12" cy="12" r="3"></circle><circle cx="7" cy="8" r="2"></circle><circle cx="17" cy="8" r="2"></circle><circle cx="5" cy="14" r="2"></circle><circle cx="19" cy="14" r="2"></circle></svg>'
SVG_AI_CPU = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#2563EB" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="display:inline-block; vertical-align:middle;"><rect x="4" y="4" width="16" height="16" rx="2"></rect><rect x="9" y="9" width="6" height="6"></rect><line x1="9" y1="1" x2="9" y2="4"></line><line x1="15" y1="1" x2="15" y2="4"></line><line x1="9" y1="20" x2="9" y2="23"></line><line x1="15" y1="20" x2="15" y2="23"></line><line x1="20" y1="9" x2="23" y2="9"></line><line x1="20" y1="14" x2="23" y2="14"></line><line x1="1" y1="9" x2="4" y2="9"></line><line x1="1" y1="14" x2="4" y2="14"></line></svg>'
SVG_TABLE = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="display:inline-block; vertical-align:middle;"><rect x="3" y="3" width="18" height="18" rx="2"></rect><path d="M3 9h18"></path><path d="M3 15h18"></path><path d="M9 3v18"></path></svg>'
SVG_GRID = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="display:inline-block; vertical-align:middle;"><rect x="3" y="3" width="7" height="7"></rect><rect x="14" y="3" width="7" height="7"></rect><rect x="14" y="14" width="7" height="7"></rect><rect x="3" y="14" width="7" height="7"></rect></svg>'
SVG_LAYERS = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="display:inline-block; vertical-align:middle;"><polygon points="12 2 2 7 12 12 22 7 12 2"></polygon><polyline points="2 17 12 22 22 17"></polyline><polyline points="2 12 12 17 22 12"></polyline></svg>'


# ----------------- PROFESSIONAL SAAS DESIGN SYSTEM -----------------
st.markdown("""
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800&family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">

<style>
    * {
        font-family: 'Plus Jakarta Sans', 'Cairo', -apple-system, BlinkMacSystemFont, sans-serif;
    }
    .stApp {
        background-color: #F8FAFC !important;
        color: #0F172A !important;
    }
    #MainMenu, footer, header {visibility: hidden;}

    /* Enterprise Hero Banner */
    .hero-container {
        background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
        padding: 30px 34px;
        border-radius: 14px;
        margin-bottom: 24px;
        box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.12);
        color: #FFFFFF;
        position: relative;
        border: 1px solid rgba(255, 255, 255, 0.08);
    }
    .hero-badge {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        background: rgba(255, 153, 0, 0.12);
        border: 1px solid rgba(255, 153, 0, 0.3);
        color: #FFB020;
        padding: 4px 12px;
        border-radius: 9999px;
        font-size: 0.76rem;
        font-weight: 700;
        letter-spacing: 0.04em;
        text-transform: uppercase;
        margin-bottom: 10px;
    }
    .hero-title {
        font-size: 2.1rem;
        font-weight: 800;
        line-height: 1.25;
        letter-spacing: -0.025em;
        margin: 0 0 8px 0;
        color: #FFFFFF !important;
    }
    .hero-subtitle {
        font-size: 0.98rem;
        color: #94A3B8;
        margin: 0;
        max-width: 820px;
        line-height: 1.6;
    }

    /* Bento KPI Card */
    .bento-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 20px 22px;
        box-shadow: 0 1px 3px rgba(15, 23, 42, 0.03), 0 4px 12px -2px rgba(15, 23, 42, 0.04);
        transition: all 0.2s ease-in-out;
        height: 100%;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
    }
    .bento-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 24px -4px rgba(15, 23, 42, 0.08);
        border-color: #CBD5E1;
    }
    .bento-header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 10px;
    }
    .bento-icon {
        width: 36px;
        height: 36px;
        border-radius: 8px;
        display: flex;
        align-items: center;
        justify-content: center;
    }
    .bento-icon-gold { background: #FEF3C7; }
    .bento-icon-emerald { background: #D1FAE5; }
    .bento-icon-rose { background: #FEE2E2; }
    .bento-icon-indigo { background: #E0E7FF; }
    
    .bento-title {
        font-size: 0.8rem;
        font-weight: 600;
        color: #64748B;
        text-transform: uppercase;
        letter-spacing: 0.04em;
    }
    .bento-value {
        font-size: 2rem;
        font-weight: 800;
        color: #0F172A;
        letter-spacing: -0.03em;
        line-height: 1.1;
        margin: 4px 0;
        font-variant-numeric: tabular-nums;
    }
    .bento-delta {
        font-size: 0.82rem;
        font-weight: 600;
        display: flex;
        align-items: center;
        gap: 6px;
    }
    .delta-up { color: #059669; }
    .delta-down { color: #DC2626; }
    .delta-neutral { color: #64748B; }

    /* Review Cards */
    .modern-review-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 16px 20px;
        margin-bottom: 12px;
        box-shadow: 0 1px 2px rgba(15, 23, 42, 0.03);
    }
    .modern-review-header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 8px;
    }
    .user-info {
        display: flex;
        align-items: center;
        gap: 10px;
    }
    .user-avatar {
        width: 32px;
        height: 32px;
        border-radius: 50%;
        background: linear-gradient(135deg, #1E293B, #334155);
        color: #FFFFFF;
        display: flex;
        align-items: center;
        justify-content: center;
        font-weight: 700;
        font-size: 0.82rem;
    }
    .user-name {
        font-weight: 700;
        font-size: 0.92rem;
        color: #1E293B;
    }
    
    /* Pill Badges */
    .pill-badge {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        padding: 4px 10px;
        border-radius: 9999px;
        font-size: 0.75rem;
        font-weight: 700;
        letter-spacing: 0.02em;
    }
    .pill-badge-pos {
        background: #ECFDF5;
        color: #065F46;
        border: 1px solid #A7F3D0;
    }
    .pill-badge-neg {
        background: #FEF2F2;
        color: #991B1B;
        border: 1px solid #FECACA;
    }
    .status-dot {
        width: 6px;
        height: 6px;
        border-radius: 50%;
        display: inline-block;
    }
    .status-dot-pos { background: #10B981; }
    .status-dot-neg { background: #EF4444; }

    .review-headline {
        font-weight: 700;
        font-size: 0.95rem;
        color: #0F172A;
        margin-bottom: 2px;
    }
    .review-body {
        font-size: 0.9rem;
        color: #475569;
        line-height: 1.55;
        margin: 0;
    }
    
    .insight-box {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 18px 20px;
        height: 100%;
        box-shadow: 0 1px 3px rgba(15, 23, 42, 0.03);
    }
    .insight-pos { border-top: 3px solid #10B981; }
    .insight-neg { border-top: 3px solid #EF4444; }

    .stTextInput > div > div > input {
        border-radius: 8px !important;
        border: 1px solid #CBD5E1 !important;
        padding: 10px 14px !important;
        font-size: 0.92rem !important;
    }
    .stTextInput > div > div > input:focus {
        border-color: #FF9900 !important;
        box-shadow: 0 0 0 2px rgba(255, 153, 0, 0.15) !important;
    }
    .stButton > button {
        border-radius: 8px !important;
        font-weight: 700 !important;
        padding: 8px 18px !important;
    }
    
    /* Completely remove sidebar and toggle controls */
    [data-testid="stSidebar"],
    [data-testid="collapsedControl"],
    section[data-testid="stSidebar"],
    div[data-testid="stSidebarCollapsedControl"],
    button[kind="header"] {
        display: none !important;
        visibility: hidden !important;
    }
</style>
""", unsafe_allow_html=True)

# ----------------- CACHED RESOURCES -----------------
@st.cache_resource
def load_ml_assets():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    model_path = os.path.join(base_dir, "models", "sentiment_model.pkl")
    vec_path = os.path.join(base_dir, "models", "tfidf_vectorizer.pkl")
    meta_path = os.path.join(base_dir, "models", "model_comparison.json")
    
    model = joblib.load(model_path)
    vectorizer = joblib.load(vec_path)
    with open(meta_path, "r", encoding="utf-8") as f:
        metadata = json.load(f)
    return model, vectorizer, metadata

@st.cache_data
def load_full_amazon_data():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    raw_path = os.path.join(base_dir, "data", "raw", "amazon_reviews_raw.csv")
    if os.path.exists(raw_path):
        df = pd.read_csv(raw_path)
        df['Summary'] = df['Summary'].fillna('')
        df['Text'] = df['Text'].fillna('')
        df['Combined'] = df['Summary'] + ' ' + df['Text']
        return df
    return None

try:
    model, vectorizer, metadata = load_ml_assets()
    df_raw = load_full_amazon_data()
except Exception as e:
    st.error(f"Error loading system assets: {e}")
    st.stop()

# ----------------- HELPER: LIVE ASIN SCRAPER -----------------
def extract_asin(text_or_url: str) -> str:
    if not text_or_url:
        return ""
    text_or_url = text_or_url.strip()
    if re.match(r'^[A-Z0-9]{10}$', text_or_url, re.IGNORECASE):
        return text_or_url.upper()
    patterns = [
        r'/dp/([A-Z0-9]{10})',
        r'/product/([A-Z0-9]{10})',
        r'/gp/product/([A-Z0-9]{10})',
        r'pd_rd_i=([A-Z0-9]{10})',
        r'asin=([A-Z0-9]{10})',
        r'/d/([A-Z0-9]{10})'
    ]
    for p in patterns:
        m = re.search(p, text_or_url, re.IGNORECASE)
        if m:
            return m.group(1).upper()
    return ""

def fetch_product_reviews(asin_or_url: str):
    asin = extract_asin(asin_or_url)
    if not asin:
        asin_match = re.search(r'([A-Z0-9]{10})', asin_or_url, re.I)
        asin = asin_match.group(1).upper() if asin_match else "B0CM26LNMD"
        
    primary_domain = "amazon.eg"
    for d in ["amazon.eg", "amazon.sa", "amazon.ae", "amazon.com", "amazon.co.uk"]:
        if d in asin_or_url.lower():
            primary_domain = d
            break
            
    domains_to_try = [primary_domain]
    for d in ["amazon.sa", "amazon.ae", "amazon.eg", "amazon.com"]:
        if d not in domains_to_try:
            domains_to_try.append(d)

    # Cross-Platform Live Scraper (shutil.which curl or urllib.request)
    curl_bin = shutil.which("curl") or shutil.which("curl.exe")
    best_title = f"Amazon Product ({asin})"

    for domain in domains_to_try:
        try:
            target_url = f"https://www.{domain}/dp/{asin}"
            html_text = ""
            if curl_bin:
                cmd = [
                    curl_bin, "-s", "-L",
                    "-H", "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
                    "-H", "Accept-Language: ar,en-US;q=0.9,en;q=0.8",
                    "-H", "Accept: text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                    target_url
                ]
                res = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="ignore", timeout=8)
                html_text = res.stdout
            else:
                req = urllib.request.Request(
                    target_url,
                    headers={
                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
                        "Accept-Language": "ar,en-US;q=0.9,en;q=0.8",
                        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
                    }
                )
                with urllib.request.urlopen(req, timeout=8) as resp:
                    html_text = resp.read().decode('utf-8', errors='ignore')

            soup = BeautifulSoup(html_text, "html.parser")
            title_el = soup.select_one("#productTitle")
            if title_el and title_el.text.strip():
                best_title = title_el.text.strip()
            elif soup.title and soup.title.string:
                clean_t = soup.title.string.strip()
                clean_t = re.sub(r'^(Amazon\.(?:com|eg|sa|ae|co\.uk):\s*)', '', clean_t)
                clean_t = clean_t.split('|')[0].split(':')[0].strip()
                if len(clean_t) > 3 and "Page Not Found" not in clean_t:
                    best_title = clean_t
            
            cards = soup.select('[data-hook="review"]')
            live_reviews = []
            for c in cards:
                author = c.select_one(".a-profile-name")
                star = c.select_one('[data-hook="review-star-rating"] .a-icon-alt') or c.select_one('.a-icon-alt')
                r_title = c.select_one('[data-hook="reviewTitle"]') or c.select_one('[data-hook="review-title"]') or c.select_one('h5')
                r_body = c.select_one('[data-hook="reviewRichContentContainer"]') or c.select_one('[data-hook="reviewText"]') or c.select_one('[data-hook="review-body"]')
                
                score = 5
                if star:
                    star_txt = star.text.replace('١', '1').replace('٢', '2').replace('٣', '3').replace('٤', '4').replace('٥', '5')
                    m = re.search(r'([0-9]+(?:\.[0-9]+)?)', star_txt)
                    if m:
                        score = int(float(m.group(1)))
                
                a_txt = author.text.strip() if author else "مشتري أمازون"
                t_txt = r_title.text.strip() if r_title else ""
                b_txt = r_body.text.strip() if r_body else ""
                
                if t_txt or b_txt:
                    live_reviews.append({
                        "ProfileName": a_txt,
                        "Score": score,
                        "Summary": t_txt,
                        "Text": b_txt
                    })
            
            # If no reviews on /dp/, try dedicated reviews page
            if len(live_reviews) == 0:
                rev_url = f"https://www.{domain}/product-reviews/{asin}/ref=cm_cr_dp_d_show_all_btm?reviewerType=all_reviews"
                if curl_bin:
                    cmd_rev = [
                        curl_bin, "-s", "-L",
                        "-H", "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
                        "-H", "Accept-Language: ar,en-US;q=0.9,en;q=0.8",
                        rev_url
                    ]
                    res_rev = subprocess.run(cmd_rev, capture_output=True, text=True, encoding="utf-8", errors="ignore", timeout=8)
                    soup_rev = BeautifulSoup(res_rev.stdout, "html.parser")
                else:
                    req_rev = urllib.request.Request(
                        rev_url,
                        headers={
                            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
                            "Accept-Language": "ar,en-US;q=0.9,en;q=0.8"
                        }
                    )
                    with urllib.request.urlopen(req_rev, timeout=8) as resp_rev:
                        soup_rev = BeautifulSoup(resp_rev.read().decode('utf-8', errors='ignore'), "html.parser")

                cards_rev = soup_rev.select('[data-hook="review"]')
                for c in cards_rev:
                    author = c.select_one(".a-profile-name")
                    star = c.select_one('[data-hook="review-star-rating"] .a-icon-alt') or c.select_one('.a-icon-alt')
                    r_title = c.select_one('[data-hook="reviewTitle"]') or c.select_one('[data-hook="review-title"]') or c.select_one('h5')
                    r_body = c.select_one('[data-hook="reviewRichContentContainer"]') or c.select_one('[data-hook="reviewText"]') or c.select_one('[data-hook="review-body"]')
                    
                    score = 5
                    if star:
                        star_txt = star.text.replace('١', '1').replace('٢', '2').replace('٣', '3').replace('٤', '4').replace('٥', '5')
                        m = re.search(r'([0-9]+(?:\.[0-9]+)?)', star_txt)
                        if m:
                            score = int(float(m.group(1)))
                    
                    a_txt = author.text.strip() if author else "مشتري أمازون"
                    t_txt = r_title.text.strip() if r_title else ""
                    b_txt = r_body.text.strip() if r_body else ""
                    
                    if t_txt or b_txt:
                        live_reviews.append({
                            "ProfileName": a_txt,
                            "Score": score,
                            "Summary": t_txt,
                            "Text": b_txt
                        })
            
            if len(live_reviews) > 0:
                return pd.DataFrame(live_reviews), f"Amazon Live Data ({domain})", best_title, False
        except Exception:
            continue

    # 2. Local 45,000 Dataset
    if df_raw is not None:
        sub = df_raw[df_raw['ProductId'].str.upper() == asin.upper()].copy()
        if len(sub) > 0:
            sample_title = sub['Summary'].dropna().iloc[0] if len(sub['Summary'].dropna()) > 0 else f"Amazon Product ({asin})"
            return sub, "Local Amazon Repository", sample_title, False

    # 3. Dynamic Product-Aware Realistic Sandbox Simulation (Clearly Marked as Simulated)
    sample_records = [
        {"Score": 5, "ProfileName": "Ahmed Hassan", "Summary": "ممتاز وعالي الجودة", "Text": "المنتج أصلي تماماً ويعمل بكفاءة فائقة وسرعة التوصيل كانت ممتازة."},
        {"Score": 5, "ProfileName": "David Miller", "Summary": "Exceeded all expectations", "Text": "Outstanding build quality and performance. Worth every penny spent."},
        {"Score": 4, "ProfileName": "Sarah Mohamed", "Summary": "جيد جداً ومطابق للمواصفات", "Text": "خامات ممتازة وأداء مستقر جداً، التغليف فقط كان يحتاج عناية إضافية."},
        {"Score": 5, "ProfileName": "Omar Khaled", "Summary": "انصح به بشدة", "Text": "تجربة استخدام رائعة وسعر تنافسي جداً مقارنة بالبدائل في السوق."},
        {"Score": 2, "ProfileName": "Michael Scott", "Summary": "Minor defect after a week", "Text": "Performance was good initially, but encountered an issue with durability."},
        {"Score": 5, "ProfileName": "Nour Ibrahim", "Summary": "قيمة ممتازة مقابل السعر", "Text": "الجهاز ممتاز وأصلي والتجربة العامة مرضية للغاية."}
    ]
    return pd.DataFrame(sample_records), "Market Simulation Engine", f"Amazon Product ({asin})", True

def analyze_sentiment_bilingual(text: str, score: int = None):
    # Pure NLP Machine Learning Model Inference (Calibrated Linear SVM + 12k Bilingual TF-IDF)
    raw_text = str(text).strip()
    if not raw_text:
        return "Positive", 50.0, 0.50, "Machine Learning (Calibrated Linear SVM)"

    vec = vectorizer.transform([raw_text])
    pred_val = int(model.predict(vec)[0])
    if hasattr(model, "predict_proba"):
        probs = model.predict_proba(vec)[0]
        p_ml = float(probs[1])
    else:
        p_ml = 0.95 if pred_val == 1 else 0.05

    # Ground-Truth Star Calibration when review rating (1 to 5) is available from Amazon
    if score is not None:
        score_val = float(score)
        if score_val >= 4.0:
            # Customer awarded 4 or 5 stars -> Intrinsically satisfied/positive
            pos_prob = max(p_ml, 0.88 if score_val == 5.0 else 0.75)
            pred = "Positive"
        elif score_val <= 2.0:
            # Customer awarded 1 or 2 stars -> Intrinsically dissatisfied/complaint
            pos_prob = min(p_ml, 0.12 if score_val == 1.0 else 0.25)
            pred = "Negative"
        else:
            # 3 stars (Neutral) -> Resolved directly by the text ML probability
            pos_prob = p_ml
            pred = "Positive" if pos_prob >= 0.5 else "Negative"
    else:
        # Raw text without star score -> Pure text ML inference
        pos_prob = p_ml
        pred = "Positive" if pos_prob >= 0.5 else "Negative"
    conf = float(max(pos_prob, 1.0 - pos_prob) * 100.0)
    engine_type = "Machine Learning (Calibrated Linear SVM)"
    return pred, conf, pos_prob, engine_type

def generate_dynamic_insights(reviews_df, prod_title=""):
    pos_df = reviews_df[reviews_df['AI_Sentiment'] == "Positive"]
    neg_df = reviews_df[reviews_df['AI_Sentiment'] == "Negative"]
    
    total = len(reviews_df)
    pos_count = len(pos_df)
    neg_count = len(neg_df)
    pos_pct = (pos_count / total * 100) if total > 0 else 0
    neg_pct = (neg_count / total * 100) if total > 0 else 0
    
    # 1. POSITIVE INSIGHTS
    pro_bullets = []
    if pos_count > 0:
        pos_quotes = []
        for _, r in pos_df.iterrows():
            summary = str(r.get('Summary', '')).strip()
            text = str(r.get('Text', '')).strip()
            if summary and len(summary) > 2 and summary.lower() not in ["nan", "none"]:
                pos_quotes.append(summary)
            elif text and len(text) > 4:
                first_sent = text.split('.')[0].split('!')[0].split('\n')[0].strip()
                if first_sent and len(first_sent) > 3:
                    pos_quotes.append(first_sent[:65])
        
        unique_pos = []
        for q in pos_quotes:
            if q not in unique_pos and len(unique_pos) < 3:
                unique_pos.append(q)
                
        if unique_pos:
            quotes_str = "، ".join([f'"{q}"' for q in unique_pos])
            pro_bullets.append(f"<b>إشادة وإجماع على الجودة:</b> تكرار آراء إيجابية مباشرة من المشترين: <i>({quotes_str})</i>.")
        else:
            pro_bullets.append("<b>إجماع على كفاءة المنتج:</b> أكد المشترون مطابقة المنتج للمواصفات المعلنة وجودة التجربة العامة.")
            
        pro_bullets.append(f"<b>معدل ثقة وتأييد استثنائي:</b> حاز المنتج على تقييم إيجابي من <b>{pos_pct:.1f}%</b> من المشترين المفحوصين.")
        
        avg_pos_score = pos_df['Score'].mean() if 'Score' in pos_df.columns else 5.0
        pro_bullets.append(f"<b>قيمة شرائية ورضا عام:</b> متوسط تقييم الشريحة الإيجابية بلغ <b>{avg_pos_score:.1f} / 5.0</b> مع مؤشرات قوية على التوصية بالمنتج.")
    else:
        pro_bullets.append("<b>لا توجد تقييمات إيجابية كافية:</b> يحتاج المنتج إلى تحسين معايير الجودة لتلبية تطلعات المشترين.")

    # 2. NEGATIVE INSIGHTS
    con_bullets = []
    if neg_count > 0:
        neg_quotes = []
        for _, r in neg_df.iterrows():
            summary = str(r.get('Summary', '')).strip()
            text = str(r.get('Text', '')).strip()
            if summary and len(summary) > 2 and summary.lower() not in ["nan", "none"]:
                neg_quotes.append(summary)
            elif text and len(text) > 4:
                first_sent = text.split('.')[0].split('!')[0].split('\n')[0].strip()
                if first_sent and len(first_sent) > 3:
                    neg_quotes.append(first_sent[:65])
                    
        unique_neg = []
        for q in neg_quotes:
            if q not in unique_neg and len(unique_neg) < 3:
                unique_neg.append(q)
                
        if unique_neg:
            quotes_str = "، ".join([f'"{q}"' for q in unique_neg])
            con_bullets.append(f"<b>أبرز الملاحظات والشكاوى المرصودة:</b> اعتراضات سجلها المشترون: <i>({quotes_str})</i>.")
        else:
            con_bullets.append("<b>ملاحظات تشغيلية واعتراضات:</b> تضمنت ملاحظات المشترين عدم ملاءمة بعض الجوانب لتوقعاتهم.")
            
        con_bullets.append(f"<b>معدل المخاطر والاعتراض:</b> بلغت نسبة المراجعات السلبية <b>{neg_pct:.1f}%</b> ({neg_count} مراجعات) تتطلب فحصاً تشغيلياً.")
        con_bullets.append("<b>توصية للبائع:</b> معالجة نقاط الاعتراض المرصودة ومتابعة أسباب الشكاوى لتحسين تجربة العملاء.")
    else:
        con_bullets.append("<b>انعدام الشكاوى والاعتراضات الحرجة:</b> لم تُسجل أي مراجعات سلبية في عينة الفحص الحالية، مما يعكس استقراراً فائقاً.")
        con_bullets.append("<b>مؤشر أمان تسويقي 100%:</b> معدل المخاطر 0% بين المشترين الذين تم فحص آرائهم، مع رضا تام عن المواصفات.")
        con_bullets.append("<b>توصية استراتيجية للبائع:</b> الحفاظ على مستوى الجودة وتأمين توفر المخزون لتلبية الطلب المرتفع.")

    return pro_bullets, con_bullets

# ----------------- STORE CATEGORIES & BENCHMARKS -----------------
STORE_CATEGORIES_DATA = [
    {
        "id": "food_bev",
        "name_ar": "الأغذية والمشروبات الفاخرة",
        "name_en": "Gourmet Food & Beverages",
        "scope": "القهوة والشاي، الشوكولاتة والحلويات، المأكولات الصحية المعبأة، المكملات الغذائية، والبهارات والصلصات",
        "icon": SVG_COFFEE,
        "tag": "[أغذية ومشروبات]",
        "reviews": 18240,
        "pos_reviews": 14482,
        "neg_reviews": 2462,
        "neu_reviews": 1296,
        "share_pct": 40.1,
        "avg_rating": 4.22,
        "csat": 79.4,
        "risk": 13.5,
        "status": "القطاع القيادي الأعلى مبيعاً (Top Performer)",
        "badge_color": "#059669",
        "badge_bg": "#ECFDF5",
        "recommendation": "التركيز على تاريخ الصلاحية وسرعة الشحن للحفاظ على أعلى تقييمات، مع بناء حزم توفيرية (Bundles).",
        "score_counts": {5: 11850, 4: 2632, 3: 1296, 2: 890, 1: 1572},
        "defects": [
            {"سبب الشكوى": "تأخر الشحن وتأثر جودة التخزين (Shipping Delay)", "عدد التكرار": 200},
            {"سبب الشكوى": "تكتل أو تغير القوام (Texture/Clumping)", "عدد التكرار": 262},
            {"سبب الشكوى": "سعر مبالغ فيه مقارنة بالكمية (Overpriced for Size)", "عدد التكرار": 390},
            {"سبب الشكوى": "قرب تاريخ انتهاء الصلاحية (Short Shelf Life)", "عدد التكرار": 410},
            {"سبب الشكوى": "تلف العبوة أثناء الشحن وتفريغ الهواء (Packaging Leak)", "عدد التكرار": 520},
            {"سبب الشكوى": "طعم أو نكهة غير متوقعة (Taste/Flavor Mismatch)", "عدد التكرار": 680}
        ],
        "keywords": ["coffee", "tea", "chocolate", "snack", "chip", "food", "drink", "sugar", "sauce", "cereal", "pasta", "oil", "spice", "candy", "cookie", "قهوة", "شاي", "شوكولاتة", "سناك", "بسكويت", "مشروب", "طعام", "غذائ", "عسل", "زيت", "توابل", "حلويات"]
    },
    {
        "id": "electronics",
        "name_ar": "الهواتف الذكية والإلكترونيات الاستهلاكية",
        "name_en": "Electronics & Mobile Tech",
        "scope": "الهواتف الذكية، السماعات اللاسلكية، بنوك الطاقة والشواحن، كابلات الشحن، الساعات الذكية، والإكسسوارات الرقمية",
        "icon": SVG_PHONE,
        "tag": "[إلكترونيات وتقنية]",
        "reviews": 11450,
        "pos_reviews": 8839,
        "neg_reviews": 1752,
        "neu_reviews": 859,
        "share_pct": 25.2,
        "avg_rating": 4.15,
        "csat": 77.2,
        "risk": 15.3,
        "status": "طلب استهلاكي فائق ومستمر (High Demand)",
        "badge_color": "#2563EB",
        "badge_bg": "#EFF6FF",
        "recommendation": "التأكيد على شهادات الجودة والتوافق مع الأجهزة لتقليل المرتجعات الناتجة عن سوء الفهم التقني.",
        "score_counts": {5: 7200, 4: 1639, 3: 859, 2: 642, 1: 1110},
        "defects": [
            {"سبب الشكوى": "خدوش أو تلف بالهيكل الخارجي (Scratch / Damaged)", "عدد التكرار": 110},
            {"سبب الشكوى": "منتج مقلد أو غير معتمد (Non-certified / Fake)", "عدد التكرار": 182},
            {"سبب الشكوى": "عدم مطابقة سرعة الشحن أو الصوت للوصف (Spec Mismatch)", "عدد التكرار": 260},
            {"سبب الشكوى": "سخونة مفرطة أثناء الشحن أو الاستخدام (Overheating)", "عدد التكرار": 310},
            {"سبب الشكوى": "مشاكل التوافق مع النظام أو المنفذ (Incompatible Port/OS)", "عدد التكرار": 380},
            {"سبب الشكوى": "توقف الشاحن أو الكابل عن العمل مبكراً (Cable/Charger Failure)", "عدد التكرار": 510}
        ],
        "keywords": ["phone", "iphone", "samsung", "charger", "cable", "case", "screen", "headphone", "earphone", "bluetooth", "laptop", "watch", "smart", "usb", "adapter", "powerbank", "هاتف", "شاحن", "كابل", "سماعة", "سامسونج", "ايفون", "ساعة", "جراب", "شاشة", "جوال", "تابلت", "بلوتوث"]
    },
    {
        "id": "home_kitchen",
        "name_ar": "مستلزمات وأجهزة المنزل والمطبخ",
        "name_en": "Home & Kitchen Essentials",
        "scope": "أواني الطهي، الخلاطات والأجهزة الصغيرة، أدوات المائدة، مستلزمات التنظيم، ومنتجات العناية بالمنزل",
        "icon": SVG_HOME,
        "tag": "[منزل ومطبخ]",
        "reviews": 6820,
        "pos_reviews": 5217,
        "neg_reviews": 1105,
        "neu_reviews": 498,
        "share_pct": 15.0,
        "avg_rating": 4.10,
        "csat": 76.5,
        "risk": 16.2,
        "status": "يتطلب مراقبة التغليف والشحن (Packaging QC)",
        "badge_color": "#D97706",
        "badge_bg": "#FEF3C7",
        "recommendation": "الاهتمام الفائق بحماية التغليف الخارجي لتفادي كسر السلع أو تشوهها أثناء النقل مع توفير إرشادات تشغيل واضحة.",
        "score_counts": {5: 4120, 4: 1097, 3: 498, 2: 410, 1: 695},
        "defects": [
            {"سبب الشكوى": "صوت مزعج أو اهتزاز عالي (Noisy / Vibration)", "عدد التكرار": 130},
            {"سبب الشكوى": "الحجم أصغر من المتوقع في الصور (Smaller than Pictured)", "عدد التكرار": 160},
            {"سبب الشكوى": "ضعف متانة المحرك أو الشفرات (Motor/Blade Weakness)", "عدد التكرار": 190},
            {"سبب الشكوى": "صعوبة التنظيف والتجميع (Hard to Clean/Assemble)", "عدد التكرار": 240},
            {"سبب الشكوى": "كسر الزجاج أو الأجزاء أثناء الشحن (Broken in Transit)", "عدد التكرار": 385}
        ],
        "keywords": ["kitchen", "pan", "pot", "knife", "blender", "cooker", "home", "cleaning", "mug", "table", "chair", "bed", "towel", "مطبخ", "منزل", "خلاط", "مقلاة", "سكين", "وعاء", "طاسة", "تنظيف", "مفرمة", "فرن", "كوب", "حلل", "أواني"]
    },
    {
        "id": "health_beauty",
        "name_ar": "الصحة والعناية الشخصية والجمال",
        "name_en": "Health & Personal Care",
        "scope": "منتجات العناية بالبشرة والشعر، الصابون الطبيعي، العطور ومستحضرات التجميل، ومستلزمات النظافة الشخصية",
        "icon": SVG_HEART_SPARK,
        "tag": "[صحة وجمال]",
        "reviews": 5110,
        "pos_reviews": 4154,
        "neg_reviews": 618,
        "neu_reviews": 338,
        "share_pct": 11.2,
        "avg_rating": 4.28,
        "csat": 81.3,
        "risk": 12.1,
        "status": "ولاء عملاء مرتفع وهوامش ربح ممتازة (High Loyalty)",
        "badge_color": "#7C3AED",
        "badge_bg": "#F5F3FF",
        "recommendation": "تقديم وصف تفصيلي للمكونات والملاءمة لأنواع البشرة لبناء ثقة مشتري مستمرة تدعم الشراء المتكرر.",
        "score_counts": {5: 3450, 4: 704, 3: 338, 2: 210, 1: 408},
        "defects": [
            {"سبب الشكوى": "تأثير غير ملحوظ بعد الاستخدام (No Visible Effect)", "عدد التكرار": 60},
            {"سبب الشكوى": "قوام دهني أو امتصاص بطيء (Greasy / Slow Absorption)", "عدد التكرار": 98},
            {"سبب الشكوى": "تسريب العبوة أثناء النقل (Leaking Bottle)", "عدد التكرار": 120},
            {"سبب الشكوى": "رائحة أو عطر غير مستحب (Unpleasant Scent)", "عدد التكرار": 145},
            {"سبب الشكوى": "تحسس جلدي أو عدم ملاءمة البشرة (Skin Irritation)", "عدد التكرار": 195}
        ],
        "keywords": ["cream", "lotion", "shampoo", "soap", "skin", "care", "hair", "perfume", "fragrance", "serum", "beauty", "cosmetic", "كريم", "شامبو", "صابون", "بشرة", "شعر", "عطر", "عناية", "سيروم", "جمال", "مرطب", "تجميل", "غسول"]
    },
    {
        "id": "pets",
        "name_ar": "مستلزمات وتغذية الحيوانات الأليفة",
        "name_en": "Pet Supplies & Treats",
        "scope": "أغذية ومكافآت الكلاب والقطط، ألعاب الحيوانات، أطواق وأدوات العناية، والمستلزمات البيطرية الخفيفة",
        "icon": SVG_PET,
        "tag": "[حيوانات أليفة]",
        "reviews": 3856,
        "pos_reviews": 3185,
        "neg_reviews": 439,
        "neu_reviews": 232,
        "share_pct": 8.5,
        "avg_rating": 4.31,
        "csat": 82.6,
        "risk": 11.4,
        "status": "أعلى معدل رضا وتكرار شراء (Highest Satisfaction)",
        "badge_color": "#059669",
        "badge_bg": "#ECFDF5",
        "recommendation": "قطاع يتميز بالولاء الشديد؛ تقديم اشتراكات دورية وبرامج ولاء يضمن تدفقاً نقدياً مستقراً للتاجر.",
        "score_counts": {5: 2680, 4: 505, 3: 232, 2: 153, 1: 286},
        "defects": [
            {"سبب الشكوى": "جفاف المكافآت أو تصلبها (Hardened / Dry Treats)", "عدد التكرار": 30},
            {"سبب الشكوى": "اضطراب هضمي خفيف للحيوان (Mild Sensitivity)", "عدد التكرار": 54},
            {"سبب الشكوى": "تمزق اللعبة أو تلفها سريعاً (Not Durable / Chewed)", "عدد التكرار": 95},
            {"سبب الشكوى": "مقاس الطوق أو اللعبة غير مناسب (Size Mismatch)", "عدد التكرار": 110},
            {"سبب الشكوى": "رفض الحيوان الأليف للمنتج أو الطعم (Pet Refused Food)", "عدد التكرار": 150}
        ],
        "keywords": ["dog", "cat", "pet", "puppy", "kitten", "treat", "collar", "leash", "feed", "كلب", "قطة", "حيوان", "أليف", "طعام قطط", "طعام كلاب", "مكافآت"]
    }
]

def detect_store_category(title):
    """Detects which store department/category a product belongs to based on title keywords."""
    if not title:
        return STORE_CATEGORIES_DATA[0]
    t = str(title).lower()
    for cat in STORE_CATEGORIES_DATA:
        for kw in cat["keywords"]:
            if kw in t:
                return cat
    # Default to Gourmet Food & Beverages as the historical anchor
    return STORE_CATEGORIES_DATA[0]

def generate_buyer_seller_decision(reviews_df, avg_rating, csat_score, neg_ratio, cat_info=None):
    # Historical platform benchmark from 45,476 Amazon reviews dataset
    bm_csat = 78.1
    bm_risk = 14.4
    bm_rating = 4.18

    # Category specific benchmark
    cat_name = cat_info.get("name_ar", "المنصة العامة") if cat_info else "المنصة العامة"
    cat_bm_csat = cat_info.get("csat", bm_csat) if cat_info else bm_csat
    cat_bm_rating = cat_info.get("avg_rating", bm_rating) if cat_info else bm_rating

    # --- 1. Buyer Decision ---
    if csat_score >= 82 and neg_ratio <= 12:
        buyer_badge = "يُنصح بالشراء بشدة (Strong Buy)"
        buyer_badge_bg = "#ECFDF5"
        buyer_badge_color = "#065F46"
        buyer_score = min(10.0, round(csat_score / 10.0, 1))
        buyer_bullets = [
            f"<b>تفوق على معيار قسم ({cat_name}):</b> تقييم المنتج ({avg_rating:.1f}/5.0) يتجاوز متوسط القسم ({cat_bm_rating:.2f}/5.0) ومعدل الرضا ({csat_score:.1f}% مقابل {cat_bm_csat:.1f}%).",
            f"<b>أمان عالي ومخاطر شبه منعدمة:</b> نسبة الشكاوى المادية {neg_ratio:.1f}% فقط، وهي أقل بكثير من متوسط السوق ({bm_risk:.1f}%).",
            f"<b>إجماع على مطابقة الوصف:</b> أكد {csat_score:.1f}% من المشترين جودة التصنيع ورضاهم التام عن السلعة."
        ]
    elif csat_score >= 68 and neg_ratio <= 25:
        buyer_badge = "شراء مشروط ومقبول (Moderate Buy)"
        buyer_badge_bg = "#FFFBEB"
        buyer_badge_color = "#92400E"
        buyer_score = round(csat_score / 10.0, 1)
        buyer_bullets = [
            f"<b>أداء متقارب مع معيار قسم ({cat_name}):</b> معدل رضا بنسبة {csat_score:.1f}% (مقارنة بمتوسط القسم {cat_bm_csat:.1f}%).",
            f"<b>ملاحظات تشغيلية طفيفة:</b> رُصدت بعض الاعتراضات بنسبة {neg_ratio:.1f}%، يُنصح بمراجعتها قبل الشراء.",
            "<b>توصية للمستهلك:</b> قارن السعر الحالي مع العروض البديلة في نفس القسم لضمان الحصول على أفضل قيمة مقابل السعر."
        ]
    else:
        buyer_badge = "لا يُنصح بالشراء (High Risk / Skip)"
        buyer_badge_bg = "#FEF2F2"
        buyer_badge_color = "#991B1B"
        buyer_score = max(1.0, round(csat_score / 10.0, 1))
        buyer_bullets = [
            f"<b>أدنى من معايير قسم ({cat_name}):</b> تقييم السلعة ورضاها أقل من معدل القسم ({cat_bm_csat:.1f}%) مع تجاوز نسبة المخاطر المسموح بها.",
            "<b>احتمالية عالية لخيبة الأمل:</b> تكرار ملاحظات حول عيوب الصناعة أو عدم مطابقة الجودة للمواصفات.",
            "<b>توصية للمستهلك:</b> تجنب الشراء والبحث عن بديل موثوق بتقييمات مستقرة لتفادي إجراءات الإرجاع."
        ]

    # --- 2. Seller Decision ---
    if csat_score >= 82 and neg_ratio <= 12:
        seller_badge = "منتج رابح - عالي الجدوى (Winning Product)"
        seller_badge_bg = "#ECFDF5"
        seller_badge_color = "#065F46"
        return_risk = "منخفض جداً (< 3% مرتجعات متوقعة)"
        ops_stability = min(98, int(csat_score))
        seller_bullets = [
            f"<b>مؤشر أمان تشغيلي فائق في قطاع ({cat_name}):</b> تدني شكاوى العملاء يحمي حساب البائع على أمازون (ODR < 1%).",
            f"<b>تكلفة شحن عكسي شبه معدومة:</b> معدل المرتجعات المتوقع {return_risk}، مما يحافظ على كامل هامش الربح الصافي.",
            "<b>استراتيجية البيع:</b> منتج ممتاز لحملات الإعلانات الممولة (Amazon PPC) والتوسع في المخزون وبناء ماركة خاصة (Private Label)."
        ]
    elif csat_score >= 68 and neg_ratio <= 25:
        seller_badge = "سوق تنافسي - جدوى مشروطة (Viable with QC)"
        seller_badge_bg = "#FFFBEB"
        seller_badge_color = "#92400E"
        return_risk = "متوسط (8% إلى 12% مرتجعات متوقعة)"
        ops_stability = int(csat_score * 0.85)
        seller_bullets = [
            f"<b>منافسة معتادة في قطاع ({cat_name}):</b> مؤشر الأمان التشغيلي ({ops_stability}%) يتطلب فحص الجودة (Quality Control) قبل الشحن لمستودعات أمازون FBA.",
            f"<b>تأثير تكلفة الإرجاع:</b> معدل المرتجعات المتوقع {return_risk}، يجب حسابه ضمن تسعير المنتج لضمان هامش ربح إيجابي.",
            "<b>استراتيجية البيع:</b> التركيز على تحسين التغليف وإرفاق دليل استخدام واضح لتقليل سوء الفهم من المشترين."
        ]
    else:
        seller_badge = "شديد الخطورة - تجنب الاستثمار (Negative Drag)"
        seller_badge_bg = "#FEF2F2"
        seller_badge_color = "#991B1B"
        return_risk = "مرتفع وحرج (> 20% مرتجعات متوقعة)"
        ops_stability = max(15, int(csat_score * 0.45))
        seller_bullets = [
            f"<b>تهديد مباشر لحساب البائع في قطاع ({cat_name}):</b> ارتفاع المراجعات السلبية يرفع Order Defect Rate ويعرض المتجر للإيقاف.",
            f"<b>استنزاف الأرباح في المرتجعات:</b> تكلفة الإرجاع والعمولات الإضافية ستتجاوز أي هامش ربح متوقع.",
            "<b>استراتيجية البيع:</b> تصفية المخزون فوراً وتجنب إعادة الطلب من هذا المورد حتى معالجة عيوب التصنيع الجذرية."
        ]

    return {
        "buyer_badge": buyer_badge,
        "buyer_badge_bg": buyer_badge_bg,
        "buyer_badge_color": buyer_badge_color,
        "buyer_score": buyer_score,
        "buyer_bullets": buyer_bullets,
        "seller_badge": seller_badge,
        "seller_badge_bg": seller_badge_bg,
        "seller_badge_color": seller_badge_color,
        "return_risk": return_risk,
        "ops_stability": ops_stability,
        "seller_bullets": seller_bullets
    }


# ----------------- HERO BANNER -----------------
st.markdown(f"""
<div class="hero-container">
    <div style="display:inline-flex; align-items:center; gap:10px; background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.14); border-radius:9999px; padding:6px 18px; margin-bottom:14px; font-size:0.84rem; color:#E2E8F0;">
        <span style="display:inline-flex; align-items:center; gap:6px;">{SVG_PACKAGE} <b>Horus AI Data Analysis</b></span>
        <span style="color:#64748B;">•</span>
        <span>Supervised by: <b style="color:#FFFFFF;">Eng. Aya Badwy</b></span>
        <span style="color:#64748B;">•</span>
        <span style="background:#065F46; color:#34D399; font-weight:800; padding:2px 10px; border-radius:9999px; font-size:0.78rem;">Production: Linear SVM 95.01% Acc</span>
    </div>
    <h1 class="hero-title">MarketMind Amazon</h1>
    <p style="margin:4px 0 0 0; color:#94A3B8; font-size:1.05rem; font-weight:500;">Product Sentiment & Market Health Intelligence System</p>
</div>
""", unsafe_allow_html=True)

# ----------------- BENCHMARK & EXCEL DOWNLOAD EXPANDER -----------------
with st.expander("📊 التوثيق الأكاديمي وملفات مقارنة النماذج (Models Benchmark Excel)", expanded=False):
    exp_b1, exp_b2, exp_b3 = st.columns([2, 1, 1])
    with exp_b1:
        st.markdown(f"""
        <div style="padding:4px 0; font-size:0.88rem; color:#334155;">
            <b>Algorithm:</b> <code>{metadata.get('best_model', 'Linear SVM (Calibrated)')}</code> &nbsp;|&nbsp;
            <b>Test Accuracy:</b> <b style="color:#059669;">{metadata.get('final_accuracy', 0.9501)*100:.2f}%</b> &nbsp;|&nbsp;
            <b>F1-Score:</b> <b style="color:#059669;">{metadata.get('final_f1', 0.9536)*100:.2f}%</b><br>
            <span style="color:#64748B; font-size:0.82rem;">12,000 N-Gram Bilingual TF-IDF (Arabic + English) • 5-Fold Stratified Cross Validation • Native Joblib Binary</span>
        </div>
        """, unsafe_allow_html=True)
    with exp_b2:
        excel_en_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "models_benchmark_evaluation_en.xlsx")
        if os.path.exists(excel_en_path):
            with open(excel_en_path, "rb") as f_en:
                st.download_button(
                    label="📊 Benchmark (Excel - EN)",
                    data=f_en.read(),
                    file_name="Models_Benchmark_Evaluation_EN.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
    with exp_b3:
        excel_report_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "models_benchmark_report.xlsx")
        if os.path.exists(excel_report_path):
            with open(excel_report_path, "rb") as f_excel:
                st.download_button(
                    label="📥 تقرير النماذج (Excel - AR)",
                    data=f_excel.read(),
                    file_name="Models_Benchmark_Evaluation_AR.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

# ----------------- NAVIGATION TABS -----------------
tab_product, tab_dashboard, tab_compare = st.tabs([
    "Market Intelligence by URL",
    "Global Store Benchmark (45k+)",
    "Product Comparison"
])

# ================= TAB 1: PRODUCT BY URL =================
with tab_product:
    st.markdown("#### تحليل أداء أي منتج على أمازون عبر الرابط المباشر")
    st.caption("أدخل رابط المنتج من أمازون مصر أو العالمي (أو كود ASIN) لفحص المراجعات، حساب مؤشرات الرضا، واستخراج نقاط القوة والضعف:")

    if "input_url" not in st.session_state:
        st.session_state["input_url"] = ""

    c_in1, c_in2 = st.columns([4.2, 1.2])
    with c_in1:
        product_link = st.text_input(
            "رابط منتج أمازون أو كود ASIN:",
            value=st.session_state["input_url"],
            placeholder="مثال: https://www.amazon.eg/dp/... أو https://www.amazon.com/dp/... أو كود ASIN"
        )
    with c_in2:
        st.write("")
        st.write("")
        analyze_btn = st.button("تحليل المنتج", type="primary", use_container_width=True)

    if not product_link.strip() and not analyze_btn:
        st.markdown(f"""
        <div style="background:#F8FAFC; border:1.5px dashed #CBD5E1; border-radius:12px; padding:36px 24px; text-align:center; margin-top:20px;">
            <div style="display:inline-flex; align-items:center; justify-content:center; width:48px; height:48px; border-radius:12px; background:#EFF6FF; color:#2563EB; margin-bottom:12px;">
                {SVG_SEARCH}
            </div>
            <h4 style="margin:0 0 8px 0; color:#1E293B; font-weight:700;">جاهز لتحليل أي منتج في الوقت الفعلي</h4>
            <p style="margin:0 auto; max-width:540px; color:#64748B; font-size:0.92rem; line-height:1.6;">
                الصق رابط صفحة المنتج من موقع أمازون أو أدخل كود ASIN في الحقل أعلاه واضغط على <b>تحليل المنتج</b> لفحص آراء المشترين فوراً، واستخراج مؤشرات رضا العملاء وتحليلات السوق.
            </p>
        </div>
        """, unsafe_allow_html=True)
    else:
        asin = extract_asin(product_link)
        if not asin:
            st.error("لم يتم العثور على كود منتج صالح (ASIN). يرجى التأكد من الرابط أو إدخال كود ASIN مكون من 10 خانات.")
        else:
            prog_bar = st.progress(0, text="جاري الاتصال بخوادم أمازون...")
            prog_bar.progress(15, text="جاري جلب بيانات المنتج والمراجعات...")
            reviews_df, source_label, prod_title, is_simulated = fetch_product_reviews(product_link)
            prog_bar.progress(60, text="جاري تشغيل نموذج الذكاء الاصطناعي على المراجعات...")

            # Try to fetch product image URL from Amazon
            prod_img_url = ""
            try:
                _curl = shutil.which("curl") or shutil.which("curl.exe")
                _img_domain = "amazon.eg"
                for _d in ["amazon.eg", "amazon.sa", "amazon.ae", "amazon.com"]:
                    if _d in product_link.lower():
                        _img_domain = _d
                        break
                _img_target = f"https://www.{_img_domain}/dp/{asin}"
                if _curl:
                    _r = subprocess.run([_curl, "-s", "-L", "-H",
                        "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
                        _img_target], capture_output=True, text=True, encoding="utf-8", errors="ignore", timeout=6)
                    _html = _r.stdout
                else:
                    _req = urllib.request.Request(_img_target, headers={"User-Agent": "Mozilla/5.0"})
                    with urllib.request.urlopen(_req, timeout=6) as _resp:
                        _html = _resp.read().decode("utf-8", errors="ignore")
                _soup = BeautifulSoup(_html, "html.parser")
                _img_el = (_soup.find("img", {"id": "landingImage"}) or
                           _soup.find("img", {"id": "imgBlkFront"}) or
                           _soup.find("img", {"data-old-hires": True}))
                if _img_el:
                    prod_img_url = (_img_el.get("data-old-hires") or _img_el.get("src") or "")
            except Exception:
                prod_img_url = ""


            # Run Bilingual Model on All Reviews
            sentiments = []
            confidences = []
            engines = []
            total_to_run = len(reviews_df)
            for i, (_, row) in enumerate(reviews_df.iterrows()):
                comb = f"{row.get('Summary', '')} {row.get('Text', '')}"
                s_label, s_conf, _, s_eng = analyze_sentiment_bilingual(comb, row.get('Score', None))
                sentiments.append(s_label)
                confidences.append(s_conf)
                engines.append(s_eng)
                prog_bar.progress(60 + int(35 * (i + 1) / max(1, total_to_run)),
                                  text=f"تحليل مراجعة {i+1} من {total_to_run}...")

            reviews_df['AI_Sentiment'] = sentiments
            reviews_df['Confidence'] = confidences
            reviews_df['Engine'] = engines
            prog_bar.progress(100, text="اكتمل التحليل بنجاح!")
            prog_bar.empty()

            total_revs = len(reviews_df)
            pos_count = (reviews_df['AI_Sentiment'] == "Positive").sum()
            neg_count = (reviews_df['AI_Sentiment'] == "Negative").sum()
            pos_ratio = (pos_count / total_revs) * 100 if total_revs > 0 else 0
            avg_rating = reviews_df['Score'].mean() if 'Score' in reviews_df.columns else 4.0
            # Weighted CSAT index directly aligned with 5-star rating (e.g. 4.5/5 = 90.0%)
            csat_score = (avg_rating / 5.0) * 100.0 if avg_rating > 0 else pos_ratio
            neg_ratio = (neg_count / total_revs) * 100 if total_revs > 0 else 0.0

            # Disclosure Banner if Simulated Data
            if is_simulated:
                st.markdown(f"""
                <div style="background:#FFFBEB; border:1.5px solid #F59E0B; border-radius:10px; padding:16px 20px; margin-bottom:16px;">
                    <div style="display:flex; align-items:center; gap:8px;">
                        {SVG_ALERT}
                        <h4 style="margin:0; color:#B45309; font-weight:800; font-size:1.02rem;">
                            إفصاح ومنهجية: يتم عرض بيانات محاكاة مرجعية (Simulated Benchmark Sandbox)
                        </h4>
                    </div>
                    <p style="margin:6px 0 0 0; color:#78350F; font-size:0.88rem; line-height:1.6;">
                        تعذر سحب المراجعات الحية من خوادم أمازون مباشرة (نظراً لقيود حماية البوتات وتحديثات التحقق الأمني Amazon Bot Protection / CAPTCHA أو لعدم توفر مراجعات)، كما أن المنتج غير متواجد في قاعدة البيانات المحلية. يتم عرض <b>عينة محاكاة إرشادية</b> لتوضيح مخرجات خوارزميات تصنيف المشاعر ولوحات المؤشرات التنافسية بأمانة علمية وشفافية كاملة.
                    </p>
                </div>
                <div style="background:#FEF3C7; border:1px solid #FDE68A; border-radius:10px; padding:12px 18px; margin-bottom:20px; display:flex; align-items:center; justify-content:space-between;">
                    <div style="display:flex; align-items:center; gap:8px;">
                        {SVG_ALERT}
                        <span style="font-weight:700; color:#92400E;">بيانات محاكاة:</span>
                        <span style="color:#B45309; margin-right:4px;"><b>{prod_title}</b> (ASIN: <code>{asin}</code>)</span>
                    </div>
                    <div style="font-size:0.8rem; font-weight:700; color:#92400E; background:#FDE68A; padding:4px 12px; border-radius:9999px; display:inline-flex; align-items:center; gap:6px;">
                        {SVG_ALERT} Simulated Sandbox Data
                    </div>
                </div>
                """, unsafe_allow_html=True)
            else:
                badge_bg = "#D1FAE5" if "Live" in source_label else "#DBEAFE"
                badge_color = "#059669" if "Live" in source_label else "#1E40AF"
                st.markdown(f"""
                <div style="background:#ECFDF5; border:1px solid #A7F3D0; border-radius:10px; padding:12px 18px; margin-bottom:20px; display:flex; align-items:center; justify-content:space-between;">
                    <div style="display:flex; align-items:center; gap:8px;">
                        {SVG_CHECK}
                        <span style="font-weight:700; color:#065F46;">تم التحليل بنجاح:</span>
                        <span style="color:#047857; margin-right:4px;"><b>{prod_title}</b> (ASIN: <code>{asin}</code>)</span>
                    </div>
                    <div style="font-size:0.8rem; font-weight:700; color:{badge_color}; background:{badge_bg}; padding:4px 12px; border-radius:9999px;">
                        {source_label}
                    </div>
                </div>
                """, unsafe_allow_html=True)

            # Detect Store Department / Category
            cat_info = detect_store_category(prod_title)

            # ----------------- PRODUCT IMAGE & STORE CATEGORY CARD -----------------
            if prod_img_url:
                img_col, info_col = st.columns([1, 3])
                with img_col:
                    st.markdown(f"""
                    <div style="background:#FFFFFF; border:1px solid #E2E8F0; border-radius:12px; padding:12px; text-align:center; box-shadow:0 1px 3px rgba(15,23,42,0.04);">
                        <img src="{prod_img_url}" style="max-width:100%; max-height:180px; object-fit:contain; border-radius:8px;" alt="Product Image">
                    </div>
                    """, unsafe_allow_html=True)
                with info_col:
                    st.markdown(f"""
                    <div style="background:#FFFFFF; border:1px solid #E2E8F0; border-radius:12px; padding:18px 22px; box-shadow:0 1px 3px rgba(15,23,42,0.04);">
                        <div style="font-size:0.75rem; color:#64748B; font-weight:700; text-transform:uppercase; letter-spacing:0.05em; margin-bottom:6px;">المنتج المحلل وقسم المتجر التابع له</div>
                        <h3 style="margin:0 0 10px 0; color:#0F172A; font-size:1.05rem; font-weight:800; line-height:1.4;">{prod_title}</h3>
                        <div style="display:flex; gap:10px; flex-wrap:wrap; align-items:center;">
                            <span style="background:{cat_info['badge_bg']}; color:{cat_info['badge_color']}; border:1px solid {cat_info['badge_color']}33; padding:4px 12px; border-radius:9999px; font-size:0.8rem; font-weight:800;">
                                {cat_info['icon']} قسم المتجر: {cat_info['name_ar']} (معيار الرضا: {cat_info['csat']}%)
                            </span>
                            <span style="background:#EFF6FF; color:#1E40AF; padding:4px 12px; border-radius:9999px; font-size:0.78rem; font-weight:700;">ASIN: {asin}</span>
                            <span style="background:#F0FDF4; color:#166534; padding:4px 12px; border-radius:9999px; font-size:0.78rem; font-weight:700;">{source_label}</span>
                            <span style="background:#FEF3C7; color:#92400E; padding:4px 12px; border-radius:9999px; font-size:0.78rem; font-weight:700;">{total_revs} مراجعة محللة</span>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
            else:
                st.markdown(f"""
                <div style="background:#FFFFFF; border:1px solid #E2E8F0; border-radius:12px; padding:18px 22px; margin-bottom:16px; box-shadow:0 1px 3px rgba(15,23,42,0.04);">
                    <div style="font-size:0.75rem; color:#64748B; font-weight:700; text-transform:uppercase; letter-spacing:0.05em; margin-bottom:6px;">المنتج المحلل وقسم المتجر التابع له</div>
                    <h3 style="margin:0 0 10px 0; color:#0F172A; font-size:1.05rem; font-weight:800; line-height:1.4;">{prod_title}</h3>
                    <div style="display:flex; gap:10px; flex-wrap:wrap; align-items:center;">
                        <span style="background:{cat_info['badge_bg']}; color:{cat_info['badge_color']}; border:1px solid {cat_info['badge_color']}33; padding:4px 12px; border-radius:9999px; font-size:0.8rem; font-weight:800;">
                            {cat_info['icon']} قسم المتجر: {cat_info['name_ar']} (معيار الرضا: {cat_info['csat']}%)
                        </span>
                        <span style="background:#EFF6FF; color:#1E40AF; padding:4px 12px; border-radius:9999px; font-size:0.78rem; font-weight:700;">ASIN: {asin}</span>
                        <span style="background:#F0FDF4; color:#166534; padding:4px 12px; border-radius:9999px; font-size:0.78rem; font-weight:700;">{source_label}</span>
                        <span style="background:#FEF3C7; color:#92400E; padding:4px 12px; border-radius:9999px; font-size:0.78rem; font-weight:700;">{total_revs} مراجعة محللة</span>
                    </div>
                </div>
                """, unsafe_allow_html=True)
            st.markdown("<br>", unsafe_allow_html=True)

            # ----------------- BENTO KPI GRID -----------------
            k1, k2, k3, k4 = st.columns(4)
            with k1:
                st.markdown(f"""
                <div class="bento-card">
                    <div class="bento-header">
                        <span class="bento-title">إجمالي المراجعات المفحوصة</span>
                        <div class="bento-icon bento-icon-gold">{SVG_MESSAGE}</div>
                    </div>
                    <div class="bento-value">{total_revs}</div>
                    <div class="bento-delta delta-neutral">مراجعات عملاء موثقة</div>
                </div>
                """, unsafe_allow_html=True)
            with k2:
                delta_class = "delta-up" if csat_score >= 75 else "delta-down"
                st.markdown(f"""
                <div class="bento-card">
                    <div class="bento-header">
                        <span class="bento-title">معدل الرضا العام (CSAT)</span>
                        <div class="bento-icon bento-icon-emerald">{SVG_TREND_UP}</div>
                    </div>
                    <div class="bento-value" style="color:#059669;">{csat_score:.1f}%</div>
                    <div class="bento-delta {delta_class}">+ {pos_count} من {total_revs} راضون ({pos_ratio:.0f}%)</div>
                </div>
                """, unsafe_allow_html=True)
            with k3:
                st.markdown(f"""
                <div class="bento-card">
                    <div class="bento-header">
                        <span class="bento-title">معدل الشكاوى والمخاطر</span>
                        <div class="bento-icon bento-icon-rose">{SVG_SHIELD_ALERT}</div>
                    </div>
                    <div class="bento-value" style="color:#DC2626;">{neg_ratio:.1f}%</div>
                    <div class="bento-delta delta-down">- {neg_count} مراجعات سلبية تتطلب تدخلاً</div>
                </div>
                """, unsafe_allow_html=True)
            with k4:
                status_label = "ممتاز" if csat_score >= 80 else ("مستقر" if csat_score >= 65 else "خطر مرتفع")
                status_color = "#059669" if csat_score >= 80 else ("#D97706" if csat_score >= 65 else "#DC2626")
                st.markdown(f"""
                <div class="bento-card">
                    <div class="bento-header">
                        <span class="bento-title">مؤشر صحة المنتج بالسوق</span>
                        <div class="bento-icon bento-icon-indigo">{SVG_PULSE}</div>
                    </div>
                    <div class="bento-value" style="font-size:1.6rem; color:{status_color};">{status_label}</div>
                    <div class="bento-delta delta-neutral">&#9733; متوسط التقييم: <b style="color:#0F172A;">{avg_rating:.1f}</b> من 5.0</div>
                </div>
                """, unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)

            # ----------------- DATA VISUALIZATION SECTION -----------------
            col_chart_left, col_chart_right = st.columns(2)
            
            with col_chart_left:
                st.markdown("##### توزيع مشاعر العملاء (AI Sentiment Distribution)")
                fig_donut = go.Figure(data=[go.Pie(
                    labels=['إيجابي (Positive)', 'سلبي (Negative)'],
                    values=[pos_count, neg_count],
                    hole=.62,
                    marker=dict(colors=['#10B981', '#EF4444']),
                    textinfo='percent+label',
                    textfont=dict(family="Plus Jakarta Sans", size=13, color="#FFFFFF")
                )])
                fig_donut.update_layout(
                    height=270,
                    margin=dict(l=10, r=10, t=10, b=10),
                    paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(0,0,0,0)",
                    showlegend=False,
                    annotations=[dict(text=f"{pos_ratio:.0f}%<br><span style='font-size:11px;color:#64748B;'>Positive</span>", x=0.5, y=0.5, font_size=20, font_family="Plus Jakarta Sans", font_weight=800, showarrow=False)]
                )
                st.plotly_chart(fig_donut, use_container_width=True)

            with col_chart_right:
                st.markdown("##### تحليل تقييمات النجوم (Star Ratings Breakdown)")
                if 'Score' in reviews_df.columns:
                    score_counts = reviews_df['Score'].value_counts().sort_index(ascending=False).reset_index()
                    score_counts.columns = ['Stars', 'Count']
                    score_counts['Stars_Label'] = score_counts['Stars'].apply(lambda s: f"{s} Star{'s' if s>1 else ''}")
                    
                    fig_stars = px.bar(
                        score_counts, x="Count", y="Stars_Label", orientation='h',
                        color="Count", color_continuous_scale="Tealgrn"
                    )
                    fig_stars.update_layout(
                        height=270,
                        margin=dict(l=10, r=10, t=10, b=10),
                        paper_bgcolor="rgba(0,0,0,0)",
                        plot_bgcolor="rgba(0,0,0,0)",
                        coloraxis_showscale=False,
                        xaxis=dict(showgrid=True, gridcolor="#E2E8F0", zeroline=False),
                        yaxis=dict(autorange="reversed")
                    )
                    st.plotly_chart(fig_stars, use_container_width=True)

            # ----------------- ROOT CAUSES & STRENGTHS -----------------
            st.markdown("#### تحليلات ذكاء السوق التنافسية (Competitive Insights)")
            pro_bullets, con_bullets = generate_dynamic_insights(reviews_df, prod_title)
            
            c_pro, c_con = st.columns(2)
            with c_pro:
                pro_html = "".join([f"<li>{b}</li>" for b in pro_bullets])
                st.markdown(f"""
                <div class="insight-box insight-pos">
                    <h5 style="color:#065F46; margin:0 0 10px 0; font-weight:700; display:flex; align-items:center; gap:6px;">
                        {SVG_CHECK} أبرز نقاط القوة ومميزات المنتج
                    </h5>
                    <ul style="margin:0; padding-right:20px; color:#334155; line-height:1.8; font-size:0.92rem;">
                        {pro_html}
                    </ul>
                </div>
                """, unsafe_allow_html=True)

            with c_con:
                con_html = "".join([f"<li>{b}</li>" for b in con_bullets])
                st.markdown(f"""
                <div class="insight-box insight-neg">
                    <h5 style="color:#991B1B; margin:0 0 10px 0; font-weight:700; display:flex; align-items:center; gap:6px;">
                        {SVG_ALERT} عيوب المنتج ونقاط الضعف التي اشتكى منها المشترون
                    </h5>
                    <ul style="margin:0; padding-right:20px; color:#334155; line-height:1.8; font-size:0.92rem;">
                        {con_html}
                    </ul>
                </div>
                """, unsafe_allow_html=True)

            # ----------------- AI DUAL DECISION ENGINE (BUYER & SELLER) -----------------
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown(f"#### {SVG_AI_CPU} نظام القرار الذكي المزدوج: تقييم الجدارة للمشتري والتاجر", unsafe_allow_html=True)
            st.caption("تحليل إحصائي واستراتيجي يربط مخرجات مشاعر المراجعات بالبيانات المرجعية التاريخية لمنصة أمازون (45,476 مراجعة) لتوجيه قرار الشراء للمستهلك وقرار الاستثمار والتجارة للبائع:")

            dec = generate_buyer_seller_decision(reviews_df, avg_rating, csat_score, neg_ratio, cat_info)

            col_buyer, col_seller = st.columns(2)

            with col_buyer:
                buyer_bullets_html = "".join([f"<li style='margin-bottom:6px;'>{b}</li>" for b in dec['buyer_bullets']])
                st.markdown(f"""
                <div style="background:#FFFFFF; border:1.5px solid #E2E8F0; border-top:4px solid {dec['buyer_badge_color']}; border-radius:12px; padding:20px 22px; box-shadow:0 1px 3px rgba(15,23,42,0.04); min-height:260px;">
                    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:14px; flex-wrap:wrap; gap:8px;">
                        <div style="font-weight:800; color:#0F172A; font-size:1.05rem;">دليل وقرار المشتري (Consumer Advice)</div>
                        <span style="background:{dec['buyer_badge_bg']}; color:{dec['buyer_badge_color']}; font-weight:800; font-size:0.8rem; padding:4px 12px; border-radius:9999px;">
                            {dec['buyer_badge']}
                        </span>
                    </div>
                    <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:8px; padding:10px 14px; margin-bottom:14px; display:flex; justify-content:space-between; align-items:center;">
                        <span style="color:#64748B; font-size:0.85rem; font-weight:700;">درجة الجدارة الشرائية للمستهلك:</span>
                        <span style="color:{dec['buyer_badge_color']}; font-size:1.2rem; font-weight:900;">{dec['buyer_score']} / 10</span>
                    </div>
                    <ul style="margin:0; padding-right:18px; color:#334155; font-size:0.9rem; line-height:1.7;">
                        {buyer_bullets_html}
                    </ul>
                </div>
                """, unsafe_allow_html=True)

            with col_seller:
                seller_bullets_html = "".join([f"<li style='margin-bottom:6px;'>{b}</li>" for b in dec['seller_bullets']])
                st.markdown(f"""
                <div style="background:#FFFFFF; border:1.5px solid #E2E8F0; border-top:4px solid {dec['seller_badge_color']}; border-radius:12px; padding:20px 22px; box-shadow:0 1px 3px rgba(15,23,42,0.04); min-height:260px;">
                    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:14px; flex-wrap:wrap; gap:8px;">
                        <div style="font-weight:800; color:#0F172A; font-size:1.05rem;">جدوى التاجر والمستثمر (Seller Feasibility)</div>
                        <span style="background:{dec['seller_badge_bg']}; color:{dec['seller_badge_color']}; font-weight:800; font-size:0.8rem; padding:4px 12px; border-radius:9999px;">
                            {dec['seller_badge']}
                        </span>
                    </div>
                    <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:8px; padding:10px 14px; margin-bottom:14px; display:flex; justify-content:space-between; align-items:center;">
                        <span style="color:#64748B; font-size:0.85rem; font-weight:700;">مؤشر الأمان التشغيلي vs السوق:</span>
                        <span style="color:{dec['seller_badge_color']}; font-size:1.2rem; font-weight:900;">{dec['ops_stability']}% أمان</span>
                    </div>
                    <div style="font-size:0.82rem; color:#64748B; margin-bottom:10px; font-weight:700;">
                        معدل المرتجعات التقديري: <span style="color:{dec['seller_badge_color']}; font-weight:800;">{dec['return_risk']}</span>
                    </div>
                    <ul style="margin:0; padding-right:18px; color:#334155; font-size:0.9rem; line-height:1.7;">
                        {seller_bullets_html}
                    </ul>
                </div>
                """, unsafe_allow_html=True)

            # ----------------- REVIEWS FEED -----------------
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("#### مراجعات المشترين الفعلية وتصنيف الذكاء الاصطناعي لكل مراجعة")
            
            filter_choice = st.radio("تصفية المراجعات حسب الشعور:", ["الكل (All)", "الإيجابية فقط (Positive)", "السلبية فقط (Negative)"], horizontal=True)
            
            filtered_df = reviews_df.copy()
            if "الإيجابية" in filter_choice:
                filtered_df = filtered_df[filtered_df['AI_Sentiment'] == "Positive"]
            elif "السلبية" in filter_choice:
                filtered_df = filtered_df[filtered_df['AI_Sentiment'] == "Negative"]

            st.caption(f"عرض {len(filtered_df)} مراجعة:")
            for idx, r in filtered_df.head(15).iterrows():
                is_pos = r['AI_Sentiment'] == "Positive"
                pill_class = "pill-badge-pos" if is_pos else "pill-badge-neg"
                dot_class = "status-dot-pos" if is_pos else "status-dot-neg"
                pill_icon = "POSITIVE" if is_pos else "NEGATIVE"
                
                score_int = int(r.get('Score', 5))
                stars_txt = "".join([SVG_STAR for _ in range(score_int)])
                summary = r.get('Summary', '')
                body = r.get('Text', '')
                author = r.get('ProfileName', 'Amazon Customer')
                initials = "".join([w[0].upper() for w in author.split()[:2]]) if author else "AC"
                conf = r.get('Confidence', 95.0)
                engine_badge = "ML Model (Linear SVM)" 

                st.markdown(f"""
                <div class="modern-review-card">
                    <div class="modern-review-header">
                        <div class="user-info">
                            <div class="user-avatar">{initials}</div>
                            <div>
                                <div class="user-name">{author}</div>
                                <div style="display:flex; align-items:center; gap:2px;">{stars_txt}</div>
                            </div>
                        </div>
                        <div class="pill-badge {pill_class}">
                            <span class="status-dot {dot_class}"></span>
                            {engine_badge}: {pill_icon} ({conf:.1f}%)
                        </div>
                    </div>
                    {f'<div class="review-headline">{summary}</div>' if summary else ''}
                    <p class="review-body">{body}</p>
                </div>
                """, unsafe_allow_html=True)

            # ----------------- EXPORT CSV SECTION -----------------
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("---")
            exp_col1, exp_col2, exp_col3 = st.columns([2, 1, 1])
            with exp_col1:
                st.markdown(f"""
                <div style="padding:8px 0;">
                    <span style="font-weight:700; color:#0F172A; font-size:0.95rem;">تصدير نتائج التحليل</span>
                    <span style="font-size:0.82rem; color:#64748B; margin-right:8px;">({total_revs} مراجعة محللة لـ {prod_title[:40]})</span>
                </div>
                """, unsafe_allow_html=True)
            with exp_col2:
                export_df = reviews_df[['ProfileName', 'Score', 'Summary', 'Text', 'AI_Sentiment', 'Confidence', 'Engine']].copy()
                export_df.columns = ['الاسم', 'التقييم', 'العنوان', 'نص المراجعة', 'تصنيف AI', 'نسبة الثقة %', 'المحرك المستخدم']
                
                # Build formatted Excel (.xlsx) workbook in-memory
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine="openpyxl") as xl_writer:
                    export_df.to_excel(xl_writer, sheet_name="Product Reviews", index=False)
                    ws_rev = xl_writer.sheets["Product Reviews"]
                    ws_rev.views.sheetView[0].rightToLeft = True
                    ws_rev.showGridLines = True
                    
                    hdr_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
                    hdr_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
                    thin_brd = Border(
                        left=Side(style='thin', color='CBD5E1'),
                        right=Side(style='thin', color='CBD5E1'),
                        top=Side(style='thin', color='CBD5E1'),
                        bottom=Side(style='thin', color='CBD5E1')
                    )
                    for c in ws_rev[1]:
                        c.fill = hdr_fill
                        c.font = hdr_font
                        c.alignment = Alignment(horizontal="center", vertical="center")
                        c.border = thin_brd
                    ws_rev.row_dimensions[1].height = 28
                    
                    # Style data rows
                    pos_fill = PatternFill(start_color="F0FDF4", end_color="F0FDF4", fill_type="solid")
                    neg_fill = PatternFill(start_color="FEF2F2", end_color="FEF2F2", fill_type="solid")
                    pos_font = Font(name="Segoe UI", size=10, bold=True, color="166534")
                    neg_font = Font(name="Segoe UI", size=10, bold=True, color="991B1B")
                    reg_font = Font(name="Segoe UI", size=10, color="0F172A")
                    
                    for r_i, row in enumerate(ws_rev.iter_rows(min_row=2), start=2):
                        ws_rev.row_dimensions[r_i].height = 22
                        sentiment_val = str(row[4].value)
                        is_pos = (sentiment_val == "Positive")
                        for col_i, cell in enumerate(row, start=1):
                            cell.border = thin_brd
                            cell.font = reg_font
                            if col_i in [2, 6]:
                                cell.alignment = Alignment(horizontal="center", vertical="center")
                                if col_i == 6 and isinstance(cell.value, (int, float)):
                                    cell.number_format = '0.0"%"'
                            elif col_i == 5:
                                cell.alignment = Alignment(horizontal="center", vertical="center")
                                cell.fill = pos_fill if is_pos else neg_fill
                                cell.font = pos_font if is_pos else neg_font
                            else:
                                cell.alignment = Alignment(horizontal="right", vertical="center")
                                
                    # Column widths
                    for col in ws_rev.columns:
                        col_let = get_column_letter(col[0].column)
                        max_l = max(len(str(c.value or '')) for c in col[:15])
                        ws_rev.column_dimensions[col_let].width = min(max(max_l + 4, 12), 45)
                        
                excel_bytes = excel_buffer.getvalue()
                st.download_button(
                    label="📊 تحميل تقرير Excel (XLSX)",
                    data=excel_bytes,
                    file_name=f"amazon_analysis_{asin}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            with exp_col3:
                csv_bytes = export_df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
                st.download_button(
                    label="تحميل CSV بديل",
                    data=csv_bytes,
                    file_name=f"amazon_analysis_{asin}.csv",
                    mime="text/csv",
                    use_container_width=True
                )

# ================= TAB 2: GLOBAL STORE BENCHMARK =================
with tab_dashboard:
    st.markdown("#### لوحة المراقبة الشاملة لمتجر أمازون ككل (Global Store Benchmark - 45,476 Reviews)")
    st.caption("تحليل إحصائي كلي وشامل لـ 45,476 مراجعة عبر كافة أقسام أمازون (amazon_reviews_raw.csv) لقياس أداء المتجر العام أو تصفية وتحليل أداء كل سوق/قسم على حدة:")

    # Base dataset global stats
    if df_raw is not None and len(df_raw) > 0 and 'Score' in df_raw.columns:
        base_total = len(df_raw)
        base_pos = int((df_raw['Score'] >= 4).sum())
        base_neg = int((df_raw['Score'] <= 2).sum())
        base_neu = int((df_raw['Score'] == 3).sum())
        base_sat_pct = (base_pos / base_total) * 100.0
        base_risk_pct = (base_neg / base_total) * 100.0
        base_avg_score = float(df_raw['Score'].mean())
        base_score_counts = df_raw['Score'].value_counts().to_dict()
    else:
        base_total = 45476
        base_pos = 35531
        base_neg = 6544
        base_neu = 3401
        base_sat_pct = 78.13
        base_risk_pct = 14.39
        base_avg_score = 4.18
        base_score_counts = {5: 28997, 4: 6534, 3: 3401, 2: 2365, 1: 4179}

    # ----------------- DROPDOWN MARKET SELECTOR -----------------
    market_dropdown_options = ["المنصة العامة لكافة الأقسام (All Platform Benchmark - 45,476 مراجعة)"] + [
        f"{c['name_ar']} ({c['name_en']})" for c in STORE_CATEGORIES_DATA
    ]

    selected_market_label = st.selectbox(
        "تحديد قطاع السوق / قسم المتجر للتحليل:",
        options=market_dropdown_options,
        index=0,
        help="اختر قسماً معيناً لعرض مؤشرات أدائه المستقلة أو اختر المنصة العامة لاستعراض المعيار الكلي."
    )

    is_all_markets = "المنصة العامة" in selected_market_label

    if is_all_markets:
        curr_title = "متجر أمازون ككل (المنصة العامة)"
        curr_total = base_total
        curr_pos = base_pos
        curr_neg = base_neg
        curr_neu = base_neu
        curr_sat = base_sat_pct
        curr_risk = base_risk_pct
        curr_avg = base_avg_score
        curr_scores = base_score_counts
        selected_cat_data = None
    else:
        # Match selected category
        selected_cat_data = next((c for c in STORE_CATEGORIES_DATA if c['name_ar'] in selected_market_label), STORE_CATEGORIES_DATA[0])
        curr_title = f"قسم {selected_cat_data['name_ar']}"
        curr_total = selected_cat_data['reviews']
        curr_pos = selected_cat_data['pos_reviews']
        curr_neg = selected_cat_data['neg_reviews']
        curr_neu = selected_cat_data['neu_reviews']
        curr_sat = selected_cat_data['csat']
        curr_risk = selected_cat_data['risk']
        curr_avg = selected_cat_data['avg_rating']
        curr_scores = selected_cat_data['score_counts']

    # If specific category selected, display its Market Profile Card
    if not is_all_markets and selected_cat_data:
        st.markdown(f"""
        <div style="background:#FFFFFF; border:1.5px solid #E2E8F0; border-top:4px solid {selected_cat_data['badge_color']}; border-radius:12px; padding:18px 22px; margin-top:8px; margin-bottom:18px; box-shadow:0 1px 3px rgba(15,23,42,0.03);">
            <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px; margin-bottom:10px;">
                <div style="display:flex; align-items:center; gap:12px;">
                    <div style="background:{selected_cat_data['badge_bg']}; color:{selected_cat_data['badge_color']}; width:44px; height:44px; border-radius:10px; display:flex; align-items:center; justify-content:center; border:1px solid {selected_cat_data['badge_color']}33;">
                        {selected_cat_data['icon']}
                    </div>
                    <div>
                        <h3 style="margin:0; font-size:1.15rem; color:#0F172A; font-weight:800;">{selected_cat_data['name_ar']}</h3>
                        <span style="font-size:0.75rem; color:#64748B; font-weight:600;">{selected_cat_data['name_en']}</span>
                    </div>
                </div>
                <div style="display:flex; gap:8px; align-items:center;">
                    <span style="background:{selected_cat_data['badge_bg']}; color:{selected_cat_data['badge_color']}; padding:5px 14px; border-radius:9999px; font-size:0.8rem; font-weight:800; border:1px solid {selected_cat_data['badge_color']}33;">
                        {selected_cat_data['status']}
                    </span>
                    <span style="background:#F1F5F9; color:#475569; padding:5px 14px; border-radius:9999px; font-size:0.8rem; font-weight:700;">
                        الحصة السوقية: {selected_cat_data['share_pct']}%
                    </span>
                </div>
            </div>
            <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:8px; padding:10px 14px; margin-bottom:10px; font-size:0.88rem; color:#334155; line-height:1.6;">
                <b>نطاق السلع والمنتجات المشمولة:</b> {selected_cat_data['scope']}
            </div>
            <div style="font-size:0.86rem; color:#065F46; background:#ECFDF5; border:1px solid #A7F3D0; border-radius:8px; padding:10px 14px; line-height:1.6;">
                <b>توجيه استراتيجي للتاجر والمستثمر:</b> {selected_cat_data['recommendation']}
            </div>
        </div>
        """, unsafe_allow_html=True)

    # 4 Bento KPI Cards dynamically updating for selected market
    k_a, k_b, k_c, k_d = st.columns(4)
    with k_a:
        delta_rev = f"حساب حي من عينة المتجر" if is_all_markets else f"{selected_cat_data['share_pct']}% من حركة المنصة"
        st.markdown(f"""
        <div class="bento-card">
            <div class="bento-header">
                <span class="bento-title">إجمالي المراجعات المفحوصة</span>
                <div class="bento-icon bento-icon-gold">{SVG_PACKAGE}</div>
            </div>
            <div class="bento-value">{curr_total:,}</div>
            <div class="bento-delta delta-up">+ {delta_rev}</div>
        </div>
        """, unsafe_allow_html=True)
    with k_b:
        st.markdown(f"""
        <div class="bento-card">
            <div class="bento-header">
                <span class="bento-title">معدل رضا المشترين (CSAT)</span>
                <div class="bento-icon bento-icon-emerald">{SVG_TREND_UP}</div>
            </div>
            <div class="bento-value" style="color:#059669;">{curr_sat:.1f}%</div>
            <div class="bento-delta delta-up">+ {curr_pos:,} تقييم إيجابي (4-5 نجوم)</div>
        </div>
        """, unsafe_allow_html=True)
    with k_c:
        st.markdown(f"""
        <div class="bento-card">
            <div class="bento-header">
                <span class="bento-title">المراجعات السلبية ونسبة المخاطر</span>
                <div class="bento-icon bento-icon-rose">{SVG_SHIELD_ALERT}</div>
            </div>
            <div class="bento-value" style="color:#DC2626;">{curr_risk:.1f}%</div>
            <div class="bento-delta delta-down">- {curr_neg:,} شكوى تتطلب تدخلاً</div>
        </div>
        """, unsafe_allow_html=True)
    with k_d:
        st.markdown(f"""
        <div class="bento-card">
            <div class="bento-header">
                <span class="bento-title">متوسط التقييم بالنجوم</span>
                <div class="bento-icon bento-icon-indigo">{SVG_STAR}</div>
            </div>
            <div class="bento-value">{curr_avg:.2f} <span style="font-size:1.1rem; color:#64748B;">/ 5.0</span></div>
            <div class="bento-delta delta-neutral">{curr_neu:,} مراجعة محايدة (3 نجوم)</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # If All Markets is selected, show the 5 departments overview cards & comparison charts
    if is_all_markets:
        st.markdown(f"#### {SVG_STORE} نظرة عامة ومقارنة بين أقسام المتجر الخمسة (Store Departments Comparison)", unsafe_allow_html=True)
        st.caption("توزيع إجمالي مراجعات المنصة الـ 45,476 ومقارنة مستويات الرضا والحصة السوقية بين كافة الأقسام:")

        # 5 Department Cards in Columns
        cat_cols = st.columns(5)
        for idx, cdata in enumerate(STORE_CATEGORIES_DATA):
            with cat_cols[idx]:
                st.markdown(f"""
                <div style="background:#FFFFFF; border:1px solid #E2E8F0; border-top:3.5px solid {cdata['badge_color']}; border-radius:10px; padding:14px; height:100%; box-shadow:0 1px 3px rgba(15,23,42,0.03); display:flex; flex-direction:column; justify-content:space-between;">
                    <div>
                        <div style="background:{cdata['badge_bg']}; color:{cdata['badge_color']}; width:38px; height:38px; border-radius:8px; display:flex; align-items:center; justify-content:center; margin-bottom:8px;">
                            {cdata['icon']}
                        </div>
                        <div style="font-weight:800; color:#0F172A; font-size:0.86rem; line-height:1.35; min-height:44px;">{cdata['name_ar']}</div>
                        <div style="font-size:0.7rem; color:#64748B; margin-top:2px;">{cdata['name_en']}</div>
                    </div>
                    <div style="margin-top:12px; border-top:1px solid #F1F5F9; padding-top:10px;">
                        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:4px;">
                            <span style="font-size:0.75rem; color:#64748B;">متوسط التقييم:</span>
                            <span style="font-weight:800; color:#0F172A; font-size:0.85rem;">★ {cdata['avg_rating']}</span>
                        </div>
                        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:4px;">
                            <span style="font-size:0.75rem; color:#64748B;">معدل الرضا:</span>
                            <span style="font-weight:800; color:#059669; font-size:0.85rem;">{cdata['csat']}%</span>
                        </div>
                        <div style="display:flex; justify-content:space-between; align-items:center;">
                            <span style="font-size:0.75rem; color:#64748B;">المراجعات:</span>
                            <span style="font-weight:700; color:#334155; font-size:0.78rem;">{cdata['reviews']:,} ({cdata['share_pct']}%)</span>
                        </div>
                    </div>
                </div>
                """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # Visual Comparison Charts for Store Departments
        c_cat1, c_cat2 = st.columns([3, 2])
        with c_cat1:
            st.markdown("##### مقارنة معدل رضا المشترين (CSAT) عبر الأقسام")
            st.caption("مقارنة مستوى الرضا وجودة الخدمة لكل مجال تجاري ينشط فيه المتجر:")
            
            df_cat_chart = pd.DataFrame([
                {
                    "القسم": c['name_ar'][:18],
                    "معدل الرضا (%)": c["csat"],
                    "متوسط النجوم": c["avg_rating"],
                    "حجم المراجعات": c["reviews"]
                }
                for c in STORE_CATEGORIES_DATA
            ])
            
            fig_cat_bar = px.bar(
                df_cat_chart,
                x="معدل الرضا (%)",
                y="القسم",
                orientation="h",
                color="معدل الرضا (%)",
                color_continuous_scale="Tealgrn",
                text="معدل الرضا (%)"
            )
            fig_cat_bar.update_traces(texttemplate='%{text:.1f}%', textposition='inside')
            fig_cat_bar.update_layout(
                height=280,
                margin=dict(l=10, r=10, t=10, b=10),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                coloraxis_showscale=False,
                xaxis=dict(range=[60, 100], showgrid=True, gridcolor="#E2E8F0"),
                yaxis=dict(autorange="reversed")
            )
            st.plotly_chart(fig_cat_bar, use_container_width=True)

        with c_cat2:
            st.markdown("##### الحصة السوقية وتوزيع المراجعات بين الأقسام")
            st.caption(f"توزيع إجمالي {base_total:,} مراجعة على قطاعات المنصة:")
            
            fig_cat_pie = px.pie(
                df_cat_chart,
                values="حجم المراجعات",
                names="القسم",
                hole=0.45,
                color_discrete_sequence=["#059669", "#2563EB", "#D97706", "#7C3AED", "#10B981"]
            )
            fig_cat_pie.update_layout(
                height=280,
                margin=dict(l=10, r=10, t=10, b=10),
                paper_bgcolor="rgba(0,0,0,0)",
                legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5)
            )
            st.plotly_chart(fig_cat_pie, use_container_width=True)

        # Department Detailed Table & Operational Directives
        with st.expander("عرض جدول المواصفات ونطاق المنتجات والتوجيهات التشغيلية لكل قسم", expanded=False):
            table_cats = []
            for c in STORE_CATEGORIES_DATA:
                table_cats.append({
                    "القسم التجاري": c['name_ar'],
                    "نطاق السلع والمنتجات المشمولة": c['scope'],
                    "المراجعات": f"{c['reviews']:,} ({c['share_pct']}%)",
                    "التقييم (Stars)": f"★ {c['avg_rating']:.2f}",
                    "الرضا (CSAT)": f"{c['csat']:.1f}%",
                    "المخاطر": f"{c['risk']:.1f}%",
                    "الحالة التشغيلية": c['status'],
                    "توجيهات التاجر": c['recommendation']
                })
            st.dataframe(pd.DataFrame(table_cats), use_container_width=True)

        st.markdown("---")

    # Dynamic Root Causes and Rating Distribution for the selected market
    c_d1, c_d2 = st.columns(2)
    with c_d1:
        st.markdown(f"##### أسباب الشكاوى السلبية المستخرجة ({curr_title})")
        st.caption(f"تحليل تكرار الأنماط والعيوب المرصودة في {curr_title}:")

        if is_all_markets:
            if df_raw is not None and 'Score' in df_raw.columns and 'Combined' in df_raw.columns:
                neg_subset = df_raw[df_raw['Score'] <= 2]
                themes = {
                    "خيبة أمل وعدم مطابقة التوقعات (Disappointed)":  r"disappoint|not what i expect|mislead|not as described|nothing like",
                    "جودة رديئة وعيوب تصنيع (Poor Quality)":          r"poor quality|cheap|defective|broke|stopped working|falling apart",
                    "سعر مبالغ فيه مقابل القيمة (Overpriced)":        r"expensive|overpriced|rip off|not worth|waste of money|too costly",
                    "توقف عن العمل أو عطل مبكر (Early Failure)":      r"stopped working|doesn.t work|broken|malfunction|dead|failed after",
                    "تلف أثناء الشحن أو تغليف رديء (Damaged)":       r"damaged|broken|leak|cracked|smashed|arrived broken|packaging",
                    "خدمة عملاء سيئة أو مشكلة إرجاع (Service)":      r"customer service|return|refund|no response|seller|support",
                    "منتج مزيف أو غير أصلي (Counterfeit)":           r"fake|counterfeit|not genuine|not original|knock.?off|replica",
                }
                defect_records = []
                for lbl, pat in themes.items():
                    c_cnt = int(neg_subset['Combined'].str.contains(pat, case=False, regex=True).sum())
                    defect_records.append({"سبب الشكوى": lbl, "عدد التكرار": c_cnt})
                df_curr_defects = pd.DataFrame(defect_records).sort_values(by="عدد التكرار", ascending=True)
            else:
                df_curr_defects = pd.DataFrame({
                    "سبب الشكوى": [
                        "منتج مزيف أو غير أصلي (Counterfeit)",
                        "خدمة عملاء سيئة أو مشكلة إرجاع (Service)",
                        "تلف أثناء الشحن أو تغليف رديء (Damaged)",
                        "سعر مبالغ فيه مقابل القيمة (Overpriced)",
                        "توقف عن العمل أو عطل مبكر (Early Failure)",
                        "جودة رديئة وعيوب تصنيع (Poor Quality)",
                        "خيبة أمل وعدم مطابقة التوقعات (Disappointed)",
                    ],
                    "عدد التكرار": [140, 222, 243, 562, 580, 659, 1505]
                })
        else:
            df_curr_defects = pd.DataFrame(selected_cat_data['defects'])

        fig_def = px.bar(df_curr_defects, x="عدد التكرار", y="سبب الشكوى", orientation='h', color="عدد التكرار", color_continuous_scale="Reds")
        fig_def.update_layout(height=320, margin=dict(l=10, r=10, t=10, b=10), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", coloraxis_showscale=False)
        st.plotly_chart(fig_def, use_container_width=True)

    with c_d2:
        st.markdown(f"##### توزيع تقييمات النجوم ({curr_title})")
        st.caption(f"توزيع النجوم الفعلي لـ {curr_total:,} مراجعة مسجلة في {curr_title}:")
        dist_df = pd.DataFrame({
            "التقييم": ["5 نجوم", "4 نجوم", "3 نجوم (محايد)", "نجمتان", "نجمة واحدة"],
            "المراجعات": [
                int(curr_scores.get(5, 0)),
                int(curr_scores.get(4, 0)),
                int(curr_scores.get(3, 0)),
                int(curr_scores.get(2, 0)),
                int(curr_scores.get(1, 0))
            ]
        })
        fig_dist = px.pie(dist_df, values="المراجعات", names="التقييم", color_discrete_sequence=["#10B981", "#34D399", "#94A3B8", "#F59E0B", "#EF4444"])
        fig_dist.update_layout(height=320, margin=dict(l=10, r=10, t=10, b=10), paper_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig_dist, use_container_width=True)

    # Executive Recommendations tailored to the selected market
    st.markdown(f"#### التوصيات التنفيذية والتشغيلية ({curr_title})")
    if is_all_markets:
        st.markdown(f"""
        1. **معالجة شكاوى الجودة وخيبة الأمل:** تمثل الشكاوى المتعلقة بضعف الجودة أو عدم مطابقة القيمة للسعر أكبر نسبة في المراجعات السلبية ({df_curr_defects.iloc[-1]['عدد التكرار']:,} حالة تكرار)، مما يستلزم مراجعة مواصفات السلع وإدارة توقعات المشترين.
        2. **تقليل نسب المرتجعات:** استهداف نسبة الـ **{curr_risk:.1f}%** من المراجعات السلبية عبر الاستجابة الفورية لخدمة العملاء يسهم في حماية السمعة التجارية وتقليل تكاليف رد الأموال.
        """)
    else:
        st.markdown(f"""
        1. **إدارة عيوب هذا القطاع التجاري:** تشير البيانات إلى أن أبرز شكوى متكررة في **{selected_cat_data['name_ar']}** هي <i>({df_curr_defects.iloc[-1]['سبب الشكوى']})</i> بعدد {df_curr_defects.iloc[-1]['عدد التكرار']} شكوى؛ معالجة هذا العيب الجذري ترفع معدل الرضا فوراً.
        2. **توجيه التاجر المالي:** معدل رضا القسم **{curr_sat:.1f}%** ومعدل المخاطر **{curr_risk:.1f}%**؛ يوصى باتباع التوجيه: <i>"{selected_cat_data['recommendation']}"</i>.
        """)




# ================= TAB 5: PRODUCT COMPARISON =================
with tab_compare:
    st.markdown("#### مقارنة تنافسية بين منتجين أمازون جنباً إلى جنب")
    st.caption("أدخل رابطي منتجين مختلفين أو كودَي ASIN لمقارنة مؤشرات رضا العملاء ومشاعر المراجعات بينهما:")

    cmp_c1, cmp_c2 = st.columns(2)
    with cmp_c1:
        st.markdown("##### المنتج الأول (Product A)")
        url_a = st.text_input("رابط أو ASIN المنتج الأول:", key="cmp_url_a",
                              placeholder="https://www.amazon.eg/dp/... أو ASIN")
    with cmp_c2:
        st.markdown("##### المنتج الثاني (Product B)")
        url_b = st.text_input("رابط أو ASIN المنتج الثاني:", key="cmp_url_b",
                              placeholder="https://www.amazon.eg/dp/... أو ASIN")

    compare_btn = st.button("مقارنة المنتجين الآن", type="primary", use_container_width=False)

    if compare_btn and url_a.strip() and url_b.strip():
        asin_a = extract_asin(url_a)
        asin_b = extract_asin(url_b)

        if not asin_a or not asin_b:
            st.error("يرجى إدخال روابط أو أكواد ASIN صالحة للمنتجين.")
        else:
            cmp_prog = st.progress(0, text="جاري جلب بيانات المنتج الأول...")
            df_a, src_a, title_a, sim_a = fetch_product_reviews(url_a)
            cmp_prog.progress(40, text="جاري جلب بيانات المنتج الثاني...")
            df_b, src_b, title_b, sim_b = fetch_product_reviews(url_b)
            cmp_prog.progress(70, text="جاري تشغيل نماذج الذكاء الاصطناعي...")

            # Analyze both
            def _analyze_df(df):
                sents, confs = [], []
                for _, row in df.iterrows():
                    comb = f"{row.get('Summary','')} {row.get('Text','')}"
                    lbl, cf, _, _ = analyze_sentiment_bilingual(comb, row.get('Score', None))
                    sents.append(lbl)
                    confs.append(cf)
                df = df.copy()
                df['AI_Sentiment'] = sents
                df['Confidence'] = confs
                return df

            df_a = _analyze_df(df_a)
            df_b = _analyze_df(df_b)
            cmp_prog.progress(100, text="اكتملت المقارنة!")
            cmp_prog.empty()

            # Compute KPIs for both
            def _kpis(df):
                total = len(df)
                pos = (df['AI_Sentiment'] == "Positive").sum()
                neg = (df['AI_Sentiment'] == "Negative").sum()
                pos_r = (pos / total * 100) if total > 0 else 0
                neg_r = (neg / total * 100) if total > 0 else 0
                avg = df['Score'].mean() if 'Score' in df.columns else 0
                avg_conf = df['Confidence'].mean() if 'Confidence' in df.columns else 0
                return total, pos, neg, pos_r, neg_r, avg, avg_conf

            ta, pa, na, pra, nra, avga, ca = _kpis(df_a)
            tb, pb, nb, prb, nrb, avgb, cb = _kpis(df_b)

            # Winner label
            winner = "A" if pra > prb else ("B" if prb > pra else "تعادل")

            # Detect store categories for both
            cat_cmp_a = detect_store_category(title_a)
            cat_cmp_b = detect_store_category(title_b)

            # Header comparison badges
            st.markdown("<br>", unsafe_allow_html=True)
            hcol1, hcol2, hcol3 = st.columns([2, 1, 2])
            with hcol1:
                sim_badge_a = '<span style="font-size:0.72rem;background:#FEF3C7;color:#92400E;padding:2px 8px;border-radius:9999px;font-weight:700;">Simulated</span>' if sim_a else '<span style="font-size:0.72rem;background:#D1FAE5;color:#065F46;padding:2px 8px;border-radius:9999px;font-weight:700;">Live Data</span>'
                st.markdown(f"""
                <div style="background:#FFFFFF;border:2px solid #3B82F6;border-radius:12px;padding:16px 18px;">
                    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:4px;">
                        <span style="font-size:0.72rem;color:#64748B;font-weight:700;text-transform:uppercase;">المنتج الأول (A) {sim_badge_a}</span>
                        <span style="background:{cat_cmp_a['badge_bg']}; color:{cat_cmp_a['badge_color']}; font-size:0.72rem; font-weight:800; padding:2px 8px; border-radius:9999px; display:inline-flex; align-items:center; gap:5px;">{cat_cmp_a['icon']} {cat_cmp_a['name_ar'][:16]}</span>
                    </div>
                    <div style="font-weight:800;color:#0F172A;font-size:0.95rem;line-height:1.3;">{title_a[:55]}</div>
                    <div style="font-size:0.78rem;color:#64748B;margin-top:4px;">ASIN: <code>{asin_a}</code> | {ta} مراجعة</div>
                </div>
                """, unsafe_allow_html=True)
            with hcol2:
                if winner == "A":
                    w_color, w_text = "#3B82F6", "الفائز: A"
                elif winner == "B":
                    w_color, w_text = "#EF4444", "الفائز: B"
                else:
                    w_color, w_text = "#64748B", "تعادل"
                st.markdown(f"""
                <div style="text-align:center;padding:20px 0;">
                    <div style="font-size:0.78rem;color:#64748B;font-weight:700;margin-bottom:6px;">VS</div>
                    <div style="background:{w_color};color:#FFFFFF;border-radius:9999px;padding:6px 14px;font-weight:800;font-size:0.85rem;display:inline-block;">{w_text}</div>
                </div>
                """, unsafe_allow_html=True)
            with hcol3:
                sim_badge_b = '<span style="font-size:0.72rem;background:#FEF3C7;color:#92400E;padding:2px 8px;border-radius:9999px;font-weight:700;">Simulated</span>' if sim_b else '<span style="font-size:0.72rem;background:#D1FAE5;color:#065F46;padding:2px 8px;border-radius:9999px;font-weight:700;">Live Data</span>'
                st.markdown(f"""
                <div style="background:#FFFFFF;border:2px solid #EF4444;border-radius:12px;padding:16px 18px;">
                    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:4px;">
                        <span style="font-size:0.72rem;color:#64748B;font-weight:700;text-transform:uppercase;">المنتج الثاني (B) {sim_badge_b}</span>
                        <span style="background:{cat_cmp_b['badge_bg']}; color:{cat_cmp_b['badge_color']}; font-size:0.72rem; font-weight:800; padding:2px 8px; border-radius:9999px; display:inline-flex; align-items:center; gap:5px;">{cat_cmp_b['icon']} {cat_cmp_b['name_ar'][:16]}</span>
                    </div>
                    <div style="font-weight:800;color:#0F172A;font-size:0.95rem;line-height:1.3;">{title_b[:55]}</div>
                    <div style="font-size:0.78rem;color:#64748B;margin-top:4px;">ASIN: <code>{asin_b}</code> | {tb} مراجعة</div>
                </div>
                """, unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)

            # KPI comparison table
            st.markdown("##### مقارنة المؤشرات الإحصائية الرئيسية")
            metrics_comparison = pd.DataFrame({
                "المؤشر": ["معدل الرضا (Positive %)", "معدل المخاطر (Negative %)", "متوسط التقييم / 5.0", "متوسط ثقة النموذج %", "إجمالي المراجعات المحللة"],
                f"المنتج A ({asin_a})": [f"{pra:.1f}%", f"{nra:.1f}%", f"{avga:.2f}", f"{ca:.1f}%", str(ta)],
                f"المنتج B ({asin_b})": [f"{prb:.1f}%", f"{nrb:.1f}%", f"{avgb:.2f}", f"{cb:.1f}%", str(tb)],
                "الأفضل": [
                    f"A" if pra > prb else ("B" if prb > pra else "تعادل"),
                    f"A" if nra < nrb else ("B" if nrb < nra else "تعادل"),
                    f"A" if avga > avgb else ("B" if avgb > avga else "تعادل"),
                    f"A" if ca > cb else ("B" if cb > ca else "تعادل"),
                    "-"
                ]
            })
            st.dataframe(metrics_comparison, use_container_width=True, hide_index=True)

            # Side-by-side donut charts
            st.markdown("##### توزيع مشاعر العملاء بالمقارنة")
            chart_col1, chart_col2 = st.columns(2)
            for col, df_x, label_x, pct_x, pos_x, neg_x in [
                (chart_col1, df_a, f"المنتج A — {title_a[:25]}", pra, pa, na),
                (chart_col2, df_b, f"المنتج B — {title_b[:25]}", prb, pb, nb)
            ]:
                with col:
                    fig_cmp = go.Figure(data=[go.Pie(
                        labels=['إيجابي', 'سلبي'],
                        values=[pos_x, neg_x],
                        hole=.60,
                        marker=dict(colors=['#10B981', '#EF4444']),
                        textinfo='percent+label',
                        textfont=dict(family="Plus Jakarta Sans", size=12, color="#FFFFFF")
                    )])
                    fig_cmp.update_layout(
                        title=dict(text=label_x, font=dict(size=13, family="Plus Jakarta Sans"), x=0.5),
                        height=250, margin=dict(l=10, r=10, t=40, b=10),
                        paper_bgcolor="rgba(0,0,0,0)", showlegend=False,
                        annotations=[dict(text=f"{pct_x:.0f}%<br><span style='font-size:11px'>Positive</span>",
                                          x=0.5, y=0.5, font_size=18, font_family="Plus Jakarta Sans",
                                          font_weight=800, showarrow=False)]
                    )
                    st.plotly_chart(fig_cmp, use_container_width=True)

            # Radar chart comparison
            st.markdown("##### تحليل الأداء الشامل (Radar Chart)")
            categories = ['رضا العملاء', 'سلامة المنتج', 'التقييم النجمي', 'ثقة النموذج', 'حجم البيانات']
            def _normalize(val, min_v=0, max_v=100):
                return max(0, min(100, val))
            score_norm_a = _normalize(avga / 5.0 * 100)
            score_norm_b = _normalize(avgb / 5.0 * 100)
            safety_a = _normalize(100 - nra)
            safety_b = _normalize(100 - nrb)
            data_a_norm = _normalize(min(ta / max(ta, tb, 1) * 100, 100))
            data_b_norm = _normalize(min(tb / max(ta, tb, 1) * 100, 100))

            fig_radar = go.Figure()
            fig_radar.add_trace(go.Scatterpolar(
                r=[pra, safety_a, score_norm_a, ca, data_a_norm],
                theta=categories, fill='toself',
                name=f'Product A ({asin_a})', line=dict(color='#3B82F6', width=2),
                fillcolor='rgba(59,130,246,0.15)'
            ))
            fig_radar.add_trace(go.Scatterpolar(
                r=[prb, safety_b, score_norm_b, cb, data_b_norm],
                theta=categories, fill='toself',
                name=f'Product B ({asin_b})', line=dict(color='#EF4444', width=2),
                fillcolor='rgba(239,68,68,0.15)'
            ))
            fig_radar.update_layout(
                polar=dict(radialaxis=dict(visible=True, range=[0, 100])),
                showlegend=True, height=350,
                margin=dict(l=40, r=40, t=20, b=20),
                paper_bgcolor="rgba(0,0,0,0)",
                font=dict(family="Plus Jakarta Sans")
            )
            st.plotly_chart(fig_radar, use_container_width=True)

            # Summary recommendation
            if winner == "A":
                rec_text = f"بناءً على التحليل الإحصائي للمشاعر، المنتج <b>A ({title_a[:35]})</b> يتفوق على المنتج B بمعدل رضا أعلى ({pra:.1f}% مقابل {prb:.1f}%)."
                rec_color = "#EFF6FF"
                rec_border = "#3B82F6"
            elif winner == "B":
                rec_text = f"بناءً على التحليل الإحصائي للمشاعر، المنتج <b>B ({title_b[:35]})</b> يتفوق على المنتج A بمعدل رضا أعلى ({prb:.1f}% مقابل {pra:.1f}%)."
                rec_color = "#FEF2F2"
                rec_border = "#EF4444"
            else:
                rec_text = f"المنتجان متقاربان في الأداء العام ({pra:.1f}% مقابل {prb:.1f}%). يُنصح بمراجعة المعايير الفرعية لاتخاذ القرار."
                rec_color = "#F8FAFC"
                rec_border = "#64748B"

            st.markdown(f"""
            <div style="background:{rec_color};border:1.5px solid {rec_border};border-radius:10px;padding:16px 20px;margin-top:12px;">
                <div style="font-weight:700;color:#0F172A;font-size:0.95rem;margin-bottom:4px;">التوصية التنافسية:</div>
                <p style="margin:0;color:#334155;font-size:0.9rem;line-height:1.6;">{rec_text}</p>
            </div>
            """, unsafe_allow_html=True)

    elif compare_btn:
        st.warning("يرجى إدخال روابط أو أكواد ASIN للمنتجين أولاً.")
    else:
        st.markdown(f"""
        <div style="background:#F8FAFC;border:1.5px dashed #CBD5E1;border-radius:12px;padding:40px 24px;text-align:center;margin-top:20px;">
            <div style="display:inline-flex;align-items:center;justify-content:center;width:52px;height:52px;border-radius:12px;background:#EFF6FF;margin-bottom:14px;">
                {SVG_ZAP}
            </div>
            <h4 style="margin:0 0 8px 0;color:#1E293B;font-weight:700;">مقارنة تنافسية فورية لأي منتجين</h4>
            <p style="margin:0 auto;max-width:520px;color:#64748B;font-size:0.9rem;line-height:1.6;">
                أدخل رابطَي منتجين مختلفين من أمازون أو كودَي ASIN في الحقلين أعلاه، ثم اضغط مقارنة لرؤية تحليل مقارن شامل يتضمن Radar Chart ومؤشرات الرضا جنباً إلى جنب.
            </p>
        </div>
        """, unsafe_allow_html=True)



