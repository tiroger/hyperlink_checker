import streamlit as st
import docx
from docx.opc.constants import RELATIONSHIP_TYPE as RT
import requests
import tempfile
import PyPDF2
import re
import os
from PIL import Image
from io import BytesIO

def extract_hyperlinks_docx(docx_file):
    doc = docx.Document(docx_file)
    rels = doc.part.rels
    hyperlinks = []
    for rel in rels:
        if rels[rel].reltype == RT.HYPERLINK:
            hyperlinks.append(rels[rel]._target)
    return hyperlinks

def extract_hyperlinks_pdf(pdf_file):
    reader = PyPDF2.PdfReader(pdf_file)
    hyperlinks = []
    for page in reader.pages:
        content = page.extract_text()
        urls = re.findall('http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+', content)
        hyperlinks.extend(urls)
    return hyperlinks

def capture_screenshot(url, api_key):
    endpoint = "https://shot.screenshotapi.net/screenshot"
    params = {
        "token": api_key,
        "url": url,
        "full_page": "true",
        "output": "image",
        "file_type": "png",
        "wait_for_event": "load"
    }
    
    response = requests.get(endpoint, params=params)
    if response.status_code == 200:
        return response.content
    else:
        return None

def analyze_links_with_screenshots(links, api_key):
    results = []
    for link in links:
        try:
            response = requests.get(link, timeout=10)
            status_code = response.status_code
            content_returned = len(response.content) > 0

            screenshot_data = capture_screenshot(link, api_key)

            results.append({
                "link": link,
                "status_code": status_code,
                "content_returned": content_returned,
                "screenshot": screenshot_data
            })
        except Exception as e:
            results.append({
                "link": link,
                "error": str(e)
            })

    return results

st.title("Document Hyperlink Analyzer")

# Add custom CSS for the error message
st.markdown("""
<style>
.big-font {
    font-size:20px !important;
    font-weight: bold;
}
</style>
""", unsafe_allow_html=True)

# Use st.secrets to securely store your API key
api_key = st.secrets["SCREENSHOT_API_KEY"]

uploaded_file = st.file_uploader("Choose a DOCX or PDF file", type=["docx", "pdf"])

if uploaded_file is not None:
    file_extension = uploaded_file.name.split('.')[-1].lower()
    
    with tempfile.NamedTemporaryFile(delete=False, suffix=f".{file_extension}") as tmp_file:
        tmp_file.write(uploaded_file.getvalue())
        tmp_file_path = tmp_file.name

    if file_extension == 'docx':
        links = extract_hyperlinks_docx(tmp_file_path)
    elif file_extension == 'pdf':
        links = extract_hyperlinks_pdf(tmp_file_path)
    else:
        st.error("Unsupported file type. Please upload a DOCX or PDF file.")
        links = []

    os.unlink(tmp_file_path)

    st.write(f"Found {len(links)} hyperlinks in the document.")

    if st.button("Analyze Links"):
        with st.spinner("Analyzing links and capturing screenshots..."):
            results = analyze_links_with_screenshots(links, api_key)

        if results:
            error_count = sum(1 for result in results if 'error' in result)
            
            # # Create a more prominent error message
            # st.error(f"🚨 Links analyzed: {len(results)} | Links with errors: {error_count} 🚨")
            
            # Add a custom formatted message
            st.markdown(f'<p class="big-font">📊 Analysis Summary:</p>', unsafe_allow_html=True)
            st.markdown(f'<p class="big-font">Total Links: {len(results)} | Successful: {len(results) - error_count} | Errors: {error_count}</p>', unsafe_allow_html=True)
            
            # Calculate the error percentage
            error_percentage = (error_count / len(results)) * 100
            st.progress(error_percentage / 100)
            st.write(f"Error Percentage: {error_percentage:.2f}%")

        for result in results:
            st.write(result['link'])
            if 'error' in result:
                st.error(f"Error: {result['error']}")
                st.write("-"*25)
            else:
                st.write(f"Status Code: {result['status_code']}")
                st.write(f"Content Returned: {result['content_returned']}")
                if result['screenshot']:
                    try:
                        image = Image.open(BytesIO(result['screenshot']))
                        st.image(image, caption="Screenshot", use_column_width=True)
                        st.write("-"*25)
                    except Exception as e:
                        st.error(f"Failed to display screenshot: {str(e)}")
                        st.write("-"*25)
                else:
                    st.write("Screenshot capture failed")
                    st.write("-"*25)
                    
st.info("Note: This application interacts with external websites. Please use responsibly and in accordance with the terms of service of the websites being analyzed.")

st.sidebar.markdown("""
# Document Content Analyzer

## Description

The Document Content Analyzer is a tool designed to extract and analyze content from Microsoft Word (.docx) and PDF documents. It extracts hyperlinks automatically from both file types, analyzes each link for accessibility and content, and provides a comprehensive report with status codes and visual feedback.

**Note** Functionality for PDF files is limited and only works for explicit links (i.e. links that are actual hyperlinks and not text).

**Key features:**
- Extract hyperlinks from DOCX and PDF files
- Analyze each link for accessibility and content
- Capture screenshots of linked web pages
- Provide a comprehensive report with status codes and visual feedback

## How to Use

1. **Upload Your Document**
   - Click on the "Choose a DOCX or PDF file" button in the main area.
   - Select either a .docx or .pdf file from your computer.

2. **Review Extracted Links**
   - The application will automatically extract hyperlinks from the uploaded document.
   - The number of extracted links will be displayed.

3. **Analyze Links**
   - Click the "Analyze Links" button to start the analysis process.
   - The app will check each link and capture screenshots of the web pages.
   - This may take a few minutes, depending on the number of links and the websites' load times.

4. **Review Results**
   - For each link, you'll see:
     - The full URL
     - HTTP status code (e.g., 200 for success, 404 for not found)
     - Whether content was successfully returned
     - A screenshot of the web page

5. **Interpret the Results**
   - Green status codes (200-299) indicate successful connections.
   - Red status codes (400-599) indicate errors or unavailable pages.
   - Check screenshots to verify the content of each linked page.

6. **Note on External Interactions**
   - Remember that this app interacts with external websites.
   - Use responsibly and in accordance with the terms of service of the websites being analyzed.
""")

# Summary Toast notification for the number of links that returned errors
if 'results' in locals() and results:
    error_count = sum(1 for result in results if 'error' in result)
    st.toast(f"Links analyzed: {len(results)} | Links with errors: {error_count}", icon="🚨")