# Document Hyperlink Analyzer

## Description

The Document Hyperlink Analyzer is a Streamlit-based web application designed to extract and analyze hyperlinks from Microsoft Word (.docx) and PDF documents. This tool provides a comprehensive analysis of the extracted links, including their accessibility, content availability, and visual representation of the linked web pages.

## Features

- Extract hyperlinks from DOCX and PDF files
- Analyze each extracted link for accessibility
- Capture screenshots of linked web pages
- Provide a detailed report with HTTP status codes and visual feedback
- User-friendly interface built with Streamlit

## Requirements

- Python 3.7+
- Streamlit
- python-docx
- PyPDF2
- requests
- selenium
- webdriver_manager

## Installation

1. Clone this repository:
   ```
   git clone https://github.com/yourusername/document-hyperlink-analyzer.git
   cd document-hyperlink-analyzer
   ```

2. Create a virtual environment (optional but recommended):
   ```
   python -m venv venv
   source venv/bin/activate  # On Windows, use `venv\Scripts\activate`
   ```

3. Install the required packages:
   ```
   pip install -r requirements.txt
   ```

## Usage

1. Run the Streamlit app:
   ```
   streamlit run app.py
   ```

2. Open your web browser and navigate to the URL provided by Streamlit (usually `http://localhost:8501`).

3. Use the file uploader to select a DOCX or PDF file for analysis.

4. Click the "Analyze Links" button to start the hyperlink extraction and analysis process.

5. Review the results, which include:
   - The number of hyperlinks found
   - For each link:
     - The full URL
     - HTTP status code
     - Whether content was successfully returned
     - A screenshot of the web page

## Note on External Interactions

This application interacts with external websites to analyze the extracted hyperlinks. Please use responsibly and in accordance with the terms of service of the websites being analyzed.

## Contributing

Contributions to improve the Document Hyperlink Analyzer are welcome. Please feel free to submit pull requests or create issues for bugs and feature requests.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- This project uses Streamlit for the web interface.
- Selenium WebDriver is used for capturing screenshots.
- python-docx and PyPDF2 are used for document parsing.