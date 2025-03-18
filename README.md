# assignment_thappar_scoreme_102119047

# PDF Table Extractor  

## 📌 Overview  
PDF Table Extractor is a Python-based tool designed to extract tables from PDF documents and convert them into structured Excel sheets. Unlike traditional solutions, this tool does not rely on Tabula or Camelot and can efficiently process various table formats, including standard, irregular, and borderless tables.  

## 🚀 Key Features  
✅ Extracts tables from PDFs using `pdfplumber` and `PyPDF2`  
✅ Handles irregular table structures with custom extraction techniques  
✅ Detects and corrects rotated pages before extraction  
✅ Cleans and organizes extracted table data  
✅ Saves tables in Excel (`.xlsx`) format  
✅ Offers a user-friendly **Streamlit UI** for seamless interaction  

## 🛠 Installation  
Ensure you have the required dependencies installed before running the tool:  
```bash
pip install pdfplumber PyPDF2 pandas openpyxl streamlit base64
```


### 2️⃣ Running as a Streamlit Web App  
To launch the **Streamlit UI**, execute:  
```bash
streamlit run pdf_extractor.py
```
Then:  
1. Open the provided link in your web browser.  
2. Upload a PDF file.  
3. Extract and download tables in Excel format.  



This tool provides an efficient and easy-to-use solution for extracting structured data from PDFs without relying on external table recognition libraries like Tabula or Camelot. 🚀
