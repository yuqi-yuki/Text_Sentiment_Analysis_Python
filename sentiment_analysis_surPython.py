


#### IMPORT PACKAGES ####

# Import some basic packages
import pandas as pd
import streamlit as st
from datetime import date

# Import packages to deal with PDF files
import PyPDF2 as py
from PyPDF2 import PdfFileReader, PdfFileWriter
import logging

# Import text processing and NLP packages
import re
import nltk
from nltk.sentiment import SentimentIntensityAnalyzer
nltk.download('vader_lexicon')

# Import other packages
from io import BytesIO # Stream of in-memory bytes
import base64 # Encoding and decoding data using the Base64 encoding scheme

import warnings
warnings.filterwarnings("ignore")


## Create a logger used for logging events related to the PyPDF2 library
logger = logging.getLogger("PyPDF2")
logger.setLevel(logging.ERROR)


        ## Main PDF reader

        # Create an empty list
        result = []

        # Clean the PDF files
        for i in range(1, int(pdfReader.numPages)): # /!\ In the latest version of PyPDF2 (3.0.0), replace "pdfReader.numPages" by "len(pdfReader.pages)"
            pageObj = pdfReader.getPage(i) # /!\ In the latest version of PyPDF2 (3.0.0), replace "pdfReader.getPage(i)" by "pdfReader.pages[i]"
            text = pageObj.extractText() # /!\ In the latest version of PyPDF2 (3.0.0), replace "extractText()" by "extract_text()"

            # Skip the disclaimer page
            if any(word in text.lower() for word in ["disclaimer", "safe harbor", "forward-looking statements", "forward looking statements", "forward -looking statements"]):
                continue

            text = text.replace("\x87", ".")
            text = text.replace("vs.", "versus")
            text = text.replace("", ".")
            text = text.replace("■", ".")
            text = text.replace("・", ".")
            text = text.replace("•", ".")
            text = text.replace("2/39", ".")
            text = text.replace("Ｆ", ".")
            text = text.replace("「", ".")
            text = text.replace("|", ".")
            text = text.replace("」", ".")
            text = text.replace("【", ".")
            text = text.replace("】", ".")
            text = text.replace("\uf071", ".")
            text = text.replace("\uf070", ".")
            text = text.replace('\n', ".")
            text = re.sub(r'([a-z](?=[A-Z])|[A-Z](?=[A-Z][a-z]))', r'\1 ', text)
            text = text.lower()
            text = re.split(r'(?<!\d)\.(?!\d)', text)

            # Extract main information
            for sentence in text:
                for int_thing, list_kwds in dicto.items():
                    for word in list_kwds:
                        if str(list_kwds[word]) != "nan" and str(list_kwds[word]) != " ":
                            if str(list_kwds[word]) == "war":
                                if re.search(r'\b{}\b'.format(re.escape(str(list_kwds[word]))), sentence):
                                    kwds = str(list_kwds[word])
                                    category = str(int_thing)
                                    sentiment = sentiment_analyzer.polarity_scores(sentence)
                                    result.append([i+1, category, kwds, sentence, sentiment])
                            else:
                                if str(list_kwds[word]) in sentence:
                                    kwds = str(list_kwds[word])
                                    category = str(int_thing)
                                    sentiment = sentiment_analyzer.polarity_scores(sentence)
                                    result.append([i+1, category, kwds, sentence, sentiment])


        # Title reader with cleaning steps
        pageObj = pdfReader.getPage(0) # /!\ In the latest version of PyPDF2 (3.0.0), replace "pdfReader.getPage(0)" by "pdfReader.pages[0]"
        text_title = pageObj.extractText() # /!\ In the latest version of PyPDF2 (3.0.0), replace "extractText()" by "extract_text()"
        text_title = text_title.replace("\x87", ".")
        text_title = text_title.replace("vs.", "versus")
        text_title = text_title.replace("", ".")
        text_title = text_title.replace("■", ".")
        text_title = text_title.replace("・", ".")
        text_title = text_title.replace("•", ".")
        text_title = text_title.replace("2/39", ".")
        text_title = text_title.replace("Ｆ", ".")
        text_title = text_title.replace("| " , "")
        text_title = text_title.replace("「", ".")
        text_title = text_title.replace("」", ".")
        text_title = text_title.replace("【", ".")
        text_title = text_title.replace("】", ".")
        text_title = text_title.replace("\uf071", ".") 
        text_title = text_title.replace("\uf070", ".") 
        text_title = text_title.replace('\n',".") # Sometimes sentences in PDF are split with this so we can´t really use it
        text_title = re.sub(r'([a-z](?=[A-Z])|[A-Z](?=[A-Z][a-z]))', r'\1 ', text_title) # To enter a space between two words by uppercase 
        text_title = re.split(r'(?<!\d)\.(?!\d)', text_title) # To split text with periods avoiding decimals

        text_title_string = ' '.join(text_title)

                        
        ## Important variables for the Excel

        # Create dataframe from results
        df_results = pd.DataFrame(result, columns=["Page", "Categories", "Keywords","Sentence", "Sentiments"])


        # Common categories
        common_cat = df_results['Categories'].value_counts().to_frame()
        top3_cat = common_cat.head(3)
        top3_list = list(top3_cat.index.values)
        list_cat = "The most mentioned categories are : " + str(top3_list)


        # General sentiment
        df_results['Score'] = df_results['Sentiments'].apply(lambda score_dict: score_dict['compound'])
   
        df_results['Sentiment'] = df_results['Score'].apply(lambda c: 'POSITIVE' if c > 0 else ('NEGATIVE' if c < 0 else 'NEUTRAL'))

        df_results2 = df_results.drop(['Sentiments'], axis=1)

        df_results2['Score'] = pd.to_numeric(df_results2['Score'])
        total_sentiment = round(df_results2['Score'].sum(),3)

        if -1<=total_sentiment<=1 :
            opinion = " - NEUTRAL - "
        elif 1< total_sentiment <10: 
            opinion = " - GOOD - "
        elif total_sentiment >= 10 : 
            opinion = " - VERY GOOD - "
        elif -10<= total_sentiment <-1: 
            opinion = " - BAD - "   
        else:
            opinion = " - VERY BAD - "
            
        sentiment_str = "This document is rather " + str(opinion)


        # Coefficients and page score
        coef.columns = ['Keywords', 'Coefficient']
        df_results2 = pd.merge(df_results2, coef, on='Keywords')
        df_coef_kwds = df_results2

        page_df = df_coef_kwds.groupby('Page', as_index=False).agg({'Coefficient': 'sum'})
        page_df = page_df.sort_values('Coefficient', ascending=False).head(3)
        top_3_pages = page_df['Page'].values.tolist()
        top_3_coef = page_df['Coefficient'].values.tolist()

        top_3_pages_string = "the most important aspects can be found on pages : "\
                            + str(top_3_pages)


        # Get the most positive and negative pages
        df_pos = df_results2[df_results2['Score'] > 0].groupby('Page', as_index=False).agg({'Score': 'mean'})
        df_neg = df_results2[df_results2['Score'] < 0].groupby('Page', as_index=False).agg({'Score': 'mean'})

        # Get all pages with the highest/lowest score
        max_score = df_pos['Score'].max()
        min_score = df_neg['Score'].min()
        pos_pages = df_pos[df_pos['Score'] == max_score]['Page'].astype(str).values
        neg_pages = df_neg[df_neg['Score'] == min_score]['Page'].astype(str).values

        # Concatenate the page names into a single string
        pos_page_text = f"The most positive pages (mean score {round(max_score, 4)}) are pages " + ", ".join(pos_pages) # Some pages may have the same score, so join the pages with the same max average score
        neg_page_text = f"The most negative pages (mean score {round(min_score, 4)}) are pages " + ", ".join(neg_pages) # Some pages may have the same score, so join the pages with the same min average score


        # Today's date and title
        today = date.today()
        date = "DATE : " + str(today)

        str_title = "Document Reviewed : " + text_title_string

        title = {'TITLE': [str_title, date]}

        df_title = pd.DataFrame(title)


        # Executive summary
        data = { 'EXECUTIVE SUMMARY':[
                sentiment_str, 
                list_cat,
                top_3_pages_string,
                pos_page_text,
                neg_page_text
                ]}  

        df_data = pd.DataFrame(data)


    # Define a custom aggregation function to concatenate the values separated by a comma
        def concat_values(series):
            return ', '.join(str(x) for x in series)
        
        agg_dict = {
            'Coefficient': 'sum',
            'Categories': concat_values,
            'Page': concat_values,
            'Keywords': concat_values,
            'Score': 'first',
            'Sentiment': 'first'     
        }

        # Grouped dataframe and organize it
        df_results3 = df_results2
        
        df_results3['Score'] = df_results3['Score'].astype(str)

        merged_df = df_results3.groupby('Sentence').agg(agg_dict).reset_index()

        organized_results = merged_df.sort_values('Coefficient', ascending=False)
        new_order = ["Page", "Categories", "Keywords", "Sentence", "Score", "Sentiment", "Coefficient"]
        df_final = organized_results.reindex(columns=new_order)
        

        ## Set up Excel

        # Organize the Excel output
        df_title.to_excel(writer, sheet_name='Sheet1', startcol=2, index=False)
        df_data.to_excel(writer, sheet_name='Sheet1', startcol=2, startrow=4, index=False)
        df_final.to_excel(writer,sheet_name='Sheet1', startrow=11)
        
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        worksheet.set_column(0, 0, 0)
        worksheet.set_column(1, 1, 5)
        worksheet.set_column(2, 2, 20)
        worksheet.set_column(3, 3, 30)
        worksheet.set_column(4, 4, 70)
        worksheet.set_column(5, 5, 10)
        worksheet.set_column(6, 6, 10)
        worksheet.set_column(7, 7, 11)

        worksheet.set_row(0, 23)
        worksheet.set_row(4, 23)

        # Define a format for the cells
        cell_format_title = workbook.add_format({'bg_color': '#CAE1FF'})

        # Apply the format to a range of cells
        worksheet.conditional_format('C1:C1', {'type': 'cell', 'criteria': 'greater than', 'value': 0, 'format': cell_format_title})
        worksheet.conditional_format('C5:C5', {'type': 'cell', 'criteria': 'greater than', 'value': 0, 'format': cell_format_title})
        worksheet.conditional_format('B12:H12', {'type': 'cell', 'criteria': 'greater than', 'value': 0, 'format': cell_format_title})

        # Save Excel
        writer.save()
        

        ## Extract and concatenate the top 3 important PDF pages

        # Create a PDF file writer for the the top 3 pages
        writer_pdf = PdfFileWriter()

        # Add all 3 pages in single PDF
        for i in top_3_pages:
            pdf = PdfFileReader(pdf_file)
            page = pdf.getPage(i-1)
            writer_pdf.addPage(page)

        # Save the PDF to a bytes buffer
        pdf_bytes = BytesIO()
        writer_pdf.write(pdf_bytes)


        ## Download button for the top 3 important PDF pages and for the Excel output
        __, col2, col3, __ = st.columns([1, 3, 2, 1])

	    # Top 3 important PDF pages
        with col2:
            button_label_pdf = "📥 Download Top 3 PDF pages"
            button_download_pdf = st.download_button(
                label=button_label_pdf,
                data=pdf_bytes.getvalue(),
                file_name=pdf_file.name.replace(".pdf", "_top_3_pages.pdf"), # Use pdf_file.name to generate the output file name
                mime="application/pdf"
                )
