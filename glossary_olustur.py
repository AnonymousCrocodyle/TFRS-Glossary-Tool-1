import os
import csv
import traceback
from docx import Document
import spacy
from sentence_transformers import SentenceTransformer, util
from xml.sax.saxutils import escape

LOG_FILE = "error_log.txt"

def log_error(msg):
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(msg + "\n")

def main():
    try:
        FOLDER = r"C:\Users\selman.ozkan\Desktop\Muhasebe Standartları Dairesi\TFRS 18 python deneme"
        TR_FILE = "TFRS 18 Finansal Raporlamada Sunum ve Açıklama.docx"
        EN_FILE = "IFRS 18 Presentation and Disclosure in Financial Statements.docx"
        MIN_SCORE = 0.75

        print("Modeller yükleniyor...")
        log_error("Modeller yükleniyor...")
        tr_nlp = spacy.load("tr_core_news_sm")
        en_nlp = spacy.load("en_core_web_sm")
        sbert = SentenceTransformer('paraphrase-multilingual-MiniLM-L12-v2')
        log_error("Modeller yüklendi.")

        def load_paragraphs(path):
            doc = Document(path)
            return [p.text.strip() for p in doc.paragraphs if p.text.strip()]

        print("Belgeler okunuyor...")
        log_error("Belgeler okunuyor...")
        tr_pars = load_paragraphs(os.path.join(FOLDER, TR_FILE))
        en_pars = load_paragraphs(os.path.join(FOLDER, EN_FILE))

        def extract_terms(paragraphs, nlp):
            terms = set()
            for text in paragraphs:
                doc = nlp(text)
                for chunk in doc.noun_chunks:
                    term = chunk.text.strip()
                    if len(term.split()) >= 2 and 3 <= len(term) <= 50:
                        terms.add(term)
            return sorted(terms)

        print("Terimler çıkarılıyor...")
        log_error("Terimler çıkarılıyor...")
        tr_terms = extract_terms(tr_pars, tr_nlp)
        en_terms = extract_terms(en_pars, en_nlp)
        log_error(f"Türkçe terim sayısı: {len(tr_terms)}")
        log_error(f"İngilizce terim sayısı: {len(en_terms)}")

        print(f"{len(tr_terms)} Türkçe terim, {len(en_terms)} İngilizce terim bulundu.")
        print("Eşleştirme başlıyor...")
        log_error("Eşleştirme başlıyor...")

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

        print("\nEşleşmeleri onaylayın (e=evet, h=hayır):")
        approved = []
        for i, (tr, en, score) in enumerate(matches):
            print(f"{i+1}. TR: {tr}\n   EN: {en}\n   Skor: {score}")
            cevap = input("Kabul edilsin mi? (e/h): ").strip().lower()
            if cevap == "e":
                approved.append((tr, en))

        csv_path = os.path.join(FOLDER, "TFRS18_Glossary.csv")
        with open(csv_path, "w", encoding="utf-8", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["Türkçe Terim", "İngilizce Terim"])
            writer.writerows(approved)
        print(f"CSV oluşturuldu: {csv_path}")
        log_error("CSV oluşturuldu.")

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
        print(f"TMX oluşturuldu: {tmx_path}")
        log_error("TMX oluşturuldu.")

    except Exception as e:
        print("\nHata oluştu:", e)
        log_error("Hata: " + str(e))
        log_error(traceback.format_exc())

    finally:
        print("\nİşlem tamamlandı. Detaylar error_log.txt dosyasında.")
        input("Çıkmak için Enter'a bas...")

if __name__ == "__main__":
    main()
