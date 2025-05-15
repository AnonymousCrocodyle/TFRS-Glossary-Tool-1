import os
import csv
from docx import Document
import spacy
from sentence_transformers import SentenceTransformer, util
from xml.sax.saxutils import escape

# --- AYARLAR ---
FOLDER = r"C:\Users\selman.ozkan\Desktop\Muhasebe Standartları Dairesi\TFRS 18 python deneme"
TR_FILE = "TFRS 18 Finansal Raporlamada Sunum ve Açıklama.docx"
EN_FILE = "IFRS 18 Presentation and Disclosure in Financial Statements.docx"
MIN_SCORE = 0.75

print("🔄 NLP modelleri yükleniyor...")
tr_nlp = spacy.load("tr_core_news_sm")
en_nlp = spacy.load("en_core_web_sm")
sbert = SentenceTransformer('paraphrase-multilingual-MiniLM-L12-v2')

def load_paragraphs(filepath):
    doc = Document(filepath)
    return [p.text.strip() for p in doc.paragraphs if p.text.strip()]

print("📄 Belgeler okunuyor...")
tr_pars = load_paragraphs(os.path.join(FOLDER, TR_FILE))
en_pars = load_paragraphs(os.path.join(FOLDER, EN_FILE))

def extract_terms(paragraphs, nlp_model):
    terms = set()
    for text in paragraphs:
        doc = nlp_model(text)
        for chunk in doc.noun_chunks:
            term = chunk.text.strip()
            if len(term.split()) >= 2 and 3 <= len(term) <= 50:
                terms.add(term)
    return sorted(terms)

print("🧠 Terimler çıkarılıyor...")
tr_terms = extract_terms(tr_pars, tr_nlp)
en_terms = extract_terms(en_pars, en_nlp)

print(f"🔎 {len(tr_terms)} Türkçe terim, {len(en_terms)} İngilizce terim bulundu.")

print("🔗 Anlam eşleştirme başlıyor, sabırlı ol...")
tr_embeddings = sbert.encode(tr_terms, convert_to_tensor=True)
en_embeddings = sbert.encode(en_terms, convert_to_tensor=True)

cosine_scores = util.cos_sim(tr_embeddings, en_embeddings)
matches = []

for i, tr_term in enumerate(tr_terms):
    best_score = 0
    best_en = ""
    for j, en_term in enumerate(en_terms):
        score = float(cosine_scores[i][j])
        if score > best_score:
            best_score = score
            best_en = en_term
    if best_score >= MIN_SCORE:
        matches.append((tr_term, best_en, round(best_score, 3)))

print("\n📋 Eşleşmeleri gözden geçir: (e=evet, h=hayır)")
approved = []

for i, (tr, en, score) in enumerate(matches):
    print(f"\n{i+1}. TR: {tr}\n   EN: {en}\n   Skor: {score}")
    cevap = input("   Kabul edilsin mi? (e/h): ").strip().lower()
    if cevap == "e":
        approved.append((tr, en))

csv_path = os.path.join(FOLDER, "TFRS18_Glossary.csv")
with open(csv_path, "w", encoding="utf-8", newline="") as f:
    writer = csv.writer(f)
    writer.writerow(["Türkçe Terim", "İngilizce Terim"])
    writer.writerows(approved)
print(f"\n✅ CSV dosyası oluşturuldu: {csv_path}")

tmx_path = os.path.join(FOLDER, "TFRS18_Glossary.tmx")
with open(tmx_path, "w", encoding="utf-8") as f:
    f.write('<?xml version="1.0" encoding="UTF-8"?>\n')
    f.write('<tmx version="1.4">\n')
    f.write('  <header creationtool="ChatGPT" creationtoolversion="4.0" srclang="tr-TR" adminlang="en-US"/>\n')
    f.write('  <body>\n')
    for tr, en in approved:
        f.write('    <tu>\n')
        f.write(f'      <tuv xml:lang="tr-TR"><seg>{escape(tr)}</seg></tuv>\n')
        f.write(f'      <tuv xml:lang="en-US"><seg>{escape(en)}</seg></tuv>\n')
        f.write('    </tu>\n')
    f.write('  </body>\n</tmx>')
print(f"✅ TMX dosyası oluşturuldu: {tmx_path}")

input("\n✅ İşlem tamamlandı. Çıkmak için Enter'a bas.")
