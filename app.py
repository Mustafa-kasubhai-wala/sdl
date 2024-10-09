from flask import Flask, render_template, request, redirect, url_for, send_file
import pdfplumber
import pandas as pd
import os
import pdfplumber
import re


app = Flask(__name__)




# Function to extract text between headings (based on boldness and font size) with line breaks
def extract_text_between_headings(pdf_path, heading):
    extracted_text = ""
    found_heading = False
    heading_pattern = re.compile(re.escape(heading), re.IGNORECASE)
    current_font_size = None  # Track the font size of the first heading


    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            # Extract words with their font information (fontname and size)
            words = page.extract_words(extra_attrs=["fontname", "size", "top"])


            bold_text = []
            previous_top = None  # Track the position (top) of the last word to detect line changes


            for i, word in enumerate(words):
                is_bold = "Bold" in word["fontname"]
                font_size = word["size"]


                # Detect heading based on boldness and font size (assuming headings are bold or larger)
                if is_bold or (current_font_size is None or font_size > current_font_size):
                    bold_text.append(word["text"])


                    # Check if next word is not bold or the end of words list is reached
                    if (i + 1 >= len(words) or ("Bold" not in words[i + 1]["fontname"] and words[i + 1]["size"] != font_size)):
                        bold_heading = " ".join(bold_text).strip()


                        # If the bold/large text matches the desired heading
                        if not found_heading and heading_pattern.search(bold_heading):
                            found_heading = True
                            current_font_size = font_size  # Record the font size of this heading
                            bold_text = []  # Reset bold text collection
                        # If another heading is found after the first one, stop
                        elif found_heading and (is_bold or font_size == current_font_size):
                            return extracted_text.strip()


                        bold_text = []  # Reset bold text collection for next heading


                # If the first heading was found, append text until the next heading
                if found_heading and not is_bold and font_size <= current_font_size:
                    # Add a new line if the current word is from a different line (using 'top' position)
                    if previous_top is not None and word["top"] != previous_top:
                        extracted_text += "\n"
                    extracted_text += word["text"] + " "
                    previous_top = word["top"]  # Update the last position for line break detection


    return extracted_text.strip()








def extract_tables_from_pdf(pdf_path, heading):
    tables = []
    found_heading = False


    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if heading in text:
                found_heading = True
                for table in page.extract_tables():
                    df = pd.DataFrame(table[1:], columns=table[0])
                    tables.append(df)
                break
   
    if not found_heading:
        return None
   
    return tables


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    if request.method == 'POST':
        action = request.form.get('action')
        heading = request.form.get('heading')
        pdf_file = request.files['file']


        if not pdf_file or not heading:
            return "Please provide a PDF file and a heading."


        pdf_path = os.path.join("uploads", pdf_file.filename)
        pdf_file.save(pdf_path)


        if action == "extract_text":
            text = extract_text_between_headings(pdf_path, heading)
            if not text:
                return "No text found with the given heading."
            return render_template('result.html', text=text)
       
        elif action == "extract_table":
            tables = extract_tables_from_pdf(pdf_path, heading)
            if not tables:
                return "No table found with the given heading."
           
           
            output_path = os.path.join("outputs", "extracted_table.xlsx")
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                for i, table in enumerate(tables):
                    table.to_excel(writer, sheet_name=f'Sheet{i+1}', index=False)


            return send_file(output_path, as_attachment=True)


    return redirect(url_for('index'))


if __name__ == "__main__":
    if not os.path.exists('uploads'):
        os.makedirs('uploads')
    if not os.path.exists('outputs'):
        os.makedirs('outputs')
    app.run(debug=True)

