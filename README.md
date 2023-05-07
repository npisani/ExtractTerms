# Extract Sensitive Terms

## Purpose

This script is designed to help users identify and manage sensitive terms in Microsoft Word documents. By extracting unique terms from the document, the user can review and mark terms as sensitive. The script then saves the marked sensitive terms to a file for future reference.

## How the script works

1. The script reads a Microsoft Word document (`.docx` format) and extracts unique terms from the text, including paragraphs, tables, and text boxes.
2. The script checks for an existing "Sensitive Terms List.txt" file in the same directory as the "Output Files" folder and preselects any matching terms.
3. A graphical user interface (GUI) displays the unique terms with checkboxes for the user to mark as sensitive.
4. The user can save the marked terms as a new or updated "Sensitive Terms List.txt" file in a folder named "Output Files" located in the script's directory.

## Installation Instructions

1. Install Python 3.7 or later from https://www.python.org/downloads/.
2. Download or clone the repository containing the script.
3. Open a terminal or command prompt and navigate to the script's directory.
4. Run the following command to install the required packages: `pip install -r requirements.txt`.
5. (Optional) Install the NLTK data needed for tokenization by running the following commands in Python:

```python
import nltk
nltk.download('punkt')
```

## How to use the script

1. Open a terminal or command prompt and navigate to the script's directory.
2. Run the script using the command: `python extractterms.py`.
3. A file dialog will appear. Select the Microsoft Word document you want to extract terms from.
4. The script will display a list of unique terms found in the document. Check the boxes next to the terms you want to mark as sensitive.
5. Click "SAVE" to save the marked terms to a "Sensitive Terms List.txt" file in the "Output Files" folder. Click "CANCEL" to close the window without saving.
6. To update the sensitive terms list for another document, run the script again and select a new document. The previously marked terms will be preselected.