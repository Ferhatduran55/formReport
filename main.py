import pandas as pd
from fpdf import FPDF
import json
import string

class PDF(FPDF):
    def __init__(self):
        super().__init__()
        self.question_number = 1

    def add_question(self, text, content_height):
        if self.get_y() + content_height > self.page_break_trigger:
            self.add_page()
        self.cell(0, 10, txt=f"{self.question_number}. {text}", ln=True)
        self.question_number += 1

def generate_star_rating(value, max_value):
    filled_stars = "★" * value
    empty_stars = "☆" * (max_value - value)
    return filled_stars + empty_stars

def column_letter_to_index(letter):
    return string.ascii_uppercase.index(letter.upper())

def load_questions(file_path="questions.json"):
    with open(file_path, "r", encoding="utf-8") as f:
        return json.load(f)

def initialize_pdf():
    pdf = PDF()
    pdf.add_page()
    try:
        pdf.add_font('DejaVu', '', 'DejaVuSans.ttf', uni=True)
        pdf.set_font('DejaVu', size=12)
    except Exception:
        pdf.set_font('Arial', size=12)
    return pdf

def process_multiple_choice(pdf, q, answers):
    content_height = 10 + (len(q["choices"]) * 10)
    pdf.add_question(q['text'], content_height)
    for choice in q["choices"]:
        checked = "☑" if choice in answers else "☐"
        pdf.cell(0, 10, txt=f"{checked} {choice}", ln=True)
    if q.get("other"):
        other_answers = [ans for ans in answers if ans not in q["choices"]]
        if other_answers:
            for other in other_answers:
                pdf.cell(0, 10, txt=f"☑ Diğer: {other}", ln=True)
        else:
            pdf.cell(0, 10, txt="☐ Diğer: ____________________", ln=True)

def process_one_choice(pdf, q, answer):
    content_height = 10 + (len(q["choices"]) * 10)
    pdf.add_question(q['text'], content_height)
    checked_choice = False
    for choice in q["choices"]:
        if choice == answer:
            checked = "●"
            checked_choice = True
        else:
            checked = "○"
        pdf.cell(0, 10, txt=f"{checked} {choice}", ln=True)
    if q.get("other"):
        if not checked_choice and answer and not any(choice == answer for choice in q["choices"]):
            pdf.cell(0, 10, txt=f"● Diğer: {answer}", ln=True)
        else:
            pdf.cell(0, 10, txt="○ Diğer: ____________________", ln=True)

def process_integer_range(pdf, q, answer):
    try:
        val = int(answer) if answer else 0
    except ValueError:
        val = 0
    max_val = q["range"][1]
    star_rating = generate_star_rating(val, max_val)
    content_height = 15
    pdf.add_question(q['text'], content_height)
    pdf.cell(0, 10, txt=star_rating, ln=True)

def process_other_types(pdf, q, answer):
    content_height = 20
    pdf.add_question(q['text'], content_height)
    pdf.cell(0, 10, txt=f"Cevap: {answer if answer else '(Boş)'}", ln=True)

def process_question(pdf, q, row, df):
    col_idx = column_letter_to_index(q["column"])
    if col_idx >= len(df.columns):
        return

    req = q.get("requiring")
    if req:
        required_column = req.get("column")
        if required_column not in string.ascii_uppercase:
            return
        req_col_idx = column_letter_to_index(required_column)
        if req_col_idx >= len(df.columns):
            return
        required_value = req.get("value")
        req_answer = row.iloc[req_col_idx]
        if req_answer != required_value:
            return

    try:
        answer = row.iloc[col_idx]
    except Exception:
        answer = ""

    if q["type"] == "multiple_choice":
        answers = [ans.strip() for ans in str(answer).split(",")] if answer else []
        process_multiple_choice(pdf, q, answers)
    elif q["type"] == "one_choice":
        process_one_choice(pdf, q, answer)
    elif q["type"] == "integer_range":
        process_integer_range(pdf, q, answer)
    else:
        process_other_types(pdf, q, answer)
    pdf.ln()

def generate_pdfs_from_excel(excel_file):
    df = pd.read_excel(excel_file)
    questions = load_questions()
    for index, row in df.iterrows():
        person_number = index + 1
        pdf = initialize_pdf()
        pdf.cell(0, 10, txt=f"Kişi: {person_number}, Tarih: {row['Zaman damgası'].strftime('%Y-%m-%d %H:%M:%S')}", ln=True)
        for q in questions:
            if q["column"] not in string.ascii_uppercase:
                continue
            process_question(pdf, q, row, df)
        output_filename = f"{person_number}_cevaplar.pdf"
        try:
            pdf.output(output_filename)
        except Exception:
            pass

generate_pdfs_from_excel("data.xlsx")