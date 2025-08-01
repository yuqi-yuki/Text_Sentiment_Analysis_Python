# 📚 PDF Document Analyzer with Streamlit & NLP

This project is a Streamlit-based PDF analyzer that extracts key insights from financial or strategic PDF documents. It applies keyword classification, sentiment analysis, and coefficient-based scoring to identify the most relevant pages and sentences, and exports both a summary Excel report and a trimmed PDF with the top insights.

---

## 🚀 Key Features

- 📖 Extracts text from PDF documents page by page
- 🧠 Applies NLP keyword matching with sentiment analysis (VADER)
- 📊 Assigns weights to keywords and ranks important content
- 📝 Generates Excel report with:
  - Executive summary
  - Sentiment overview
  - Most important categories and sentences
- 📄 Outputs trimmed PDF with top 3 most relevant pages
- 💻 Clean Streamlit interface for simple document upload and result download

---

## 🧠 How It Works

1. **PDF Upload** via Streamlit UI
2. **Text Cleaning**: Removes unnecessary symbols and noise
3. **Sentence Splitting**: Using regex and text normalization
4. **Keyword Matching**: Based on a custom dictionary (`dicto`)
5. **Sentiment Scoring**: VADER sentiment for each matching sentence
6. **Weighted Ranking**:
   - Keywords assigned importance coefficients (`coef`)
   - Sentences and pages scored accordingly
7. **Export**:
   - Excel: structured summary and scores
   - PDF: top 3 most important pages

