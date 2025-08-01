import os
from flask import Flask, request, render_template, send_file, jsonify
from PyPDF2 import PdfMerger
from docx import Document
import subprocess
from datetime import datetime
import pytz



def fill_index(path , out_path, data, data1):
    doc = Document(path)
    # Iterate through all paragraphs and shapes (including textboxes)
    for shape in doc.element.xpath(".//w:txbxContent//w:t"):
        # Check and replace the text if it matches any key in the dictionary
        text = shape.text
        if text in data:
            shape.text = data[text]
        elif text in data1:
            shape.text = data1[text]
    doc.save(out_path)



def combine_word_to_pdf(word_files, output_pdf):
    """
    Combine multiple Word files into a single PDF.

    :param word_files: List of paths to Word files to combine.
    :param output_pdf: Path to save the combined PDF.
    """
    temp_pdf_files = []

    # Step 1: Convert each Word file to a PDF
    for word_file in word_files:
        temp = word_file
        word_file = word_file.rstrip(".docx")
        temp_pdf = f"{word_file}.pdf"
        subprocess.run([
            "libreoffice",
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            "/home/Hetindex/mysite",
            temp
        ], check=True)

        os.remove(temp)    # Remove the word file
        temp_pdf_files.append(temp_pdf)


    # Step 2: Merge all PDFs into a single PDF
    merger = PdfMerger()
    for pdf in temp_pdf_files:
        merger.append(pdf)

    # Save the merged PDF
    merger.write(output_pdf)
    merger.close()

    # Step 3: Clean up temporary PDF files
    for temp_pdf in temp_pdf_files:
        os.remove(temp_pdf)



app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/home.html')
def home():
    return render_template('home.html')


@app.route('/home.html', methods=['POST'])
def submit():
    data = {
        'subject' : request.form.get('subject'),
        'sem':request.form.get('sem'),
        'name':request.form.get('name'),
        'pen':request.form.get('pen'),
        'class':request.form.get('class'),
        'batch':request.form.get('batch'),
        'term':request.form.get('term'),
        'faculty':request.form.get('faculty'),
    }

    number = int(request.form.get('number'))
    if number < 1:
        return jsonify({"status": "error", "Desi Language":"Topa Term Work ma 0 thi vadare number nak", "message": "Term Work must be greater then 0"}), 400
    data1 = []
    for i in range(1,number+1):
        aim={
            'n':request.form.get(f'no{i}'),
            'a':request.form.get(f'aim{i}')
        }
        data1.append(aim)

    word_files = []
    for i in range(1,number+1):
        if request.form.get('department') == 'Computer':
            path = '/home/Hetindex/mysite/index.docx'
        elif request.form.get('department') == 'IT':
            path = '/home/Hetindex/mysite/index_it.docx'
        else:
            path = '/home/Hetindex/mysite/index.docx'

        out_path = f'/home/Hetindex/mysite/index{i}.docx'
        fill_index(path,out_path,data,data1[i-1])
        word_files.append(out_path)


    # Example Usage
    output_pdf = "/home/Hetindex/mysite/term_work.pdf"  # Desired output PDF file name
    combine_word_to_pdf(word_files, output_pdf)

    # Open the file in append mode ('a')
    with open("/home/Hetindex/mysite/list_of_users.txt", "a") as file:
        # Append content to the file
        sem = request.form.get('sem')
        name = request.form.get('name')
        pen = request.form.get('pen')
        dept = request.form.get('department')

        # Define IST timezone
        ist = pytz.timezone('Asia/Kolkata')

        # Get current IST time
        current_time_ist = datetime.now(ist).strftime("%Y-%m-%d %H:%M:%S")

        s =  pen + " " + name + " from " + dept + " Sem " + sem + " on " + current_time_ist + "\n"
        print(s)
        file.write(s)

    print("Content appended successfully!")

    # Send the file for download
    return send_file(output_pdf, as_attachment=True)


    return render_template('index.html')



# if __name__ == '__main__':
#     app.run(debug=True)