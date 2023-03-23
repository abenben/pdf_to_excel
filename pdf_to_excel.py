import io
import os
import pandas as pd
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
import tabula


def extract_text_by_page(pdf_path, page):
    with open(pdf_path, 'rb') as fh:
        for page_num, page in enumerate(PDFPage.get_pages(fh, caching=True, check_extractable=True)):
            if page_num + 1 == page:
                resource_manager = PDFResourceManager()
                fake_file_handle = io.StringIO()
                converter = TextConverter(resource_manager, fake_file_handle, laparams=LAParams())
                page_interpreter = PDFPageInterpreter(resource_manager, converter)
                page_interpreter.process_page(page)
                text = fake_file_handle.getvalue()
                yield text
                converter.close()
                fake_file_handle.close()


def convert_pdf_to_excel(pdf_path, page, output_path):
    temp_csv_path = "temp.csv"
    pages = [page]
    tabula.convert_into(pdf_path, temp_csv_path, output_format="csv", pages=pages)
    df = pd.read_csv(temp_csv_path)
    df.to_excel(output_path, index=False)
    os.remove(temp_csv_path)


pdf_path = "input/ver2.pdf"
output_path = "output/ver2_5p.xlsx"
page = 5

text = ""
for page_text in extract_text_by_page(pdf_path, page):
    text += page_text

with open("page{}.txt".format(page), "w", encoding="utf-8") as text_file:
    text_file.write(text)

convert_pdf_to_excel(pdf_path, page, output_path)
