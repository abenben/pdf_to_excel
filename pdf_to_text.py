from typing import Iterator
import io
import os
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage


def extract_text_from_pdf(pdf_path: str) -> str:
    """
    PDFファイルからテキストを抽出します。

    Args:
        pdf_path (str): PDFファイルへのパス。

    Returns:
        str: 抽出されたテキスト。
    """
    resource_manager = PDFResourceManager()
    fake_file_handle = io.StringIO()
    converter = TextConverter(resource_manager, fake_file_handle, laparams=LAParams())
    page_interpreter = PDFPageInterpreter(resource_manager, converter)

    with open(pdf_path, 'rb') as fh:
        for page in PDFPage.get_pages(fh, caching=True, check_extractable=True):
            page_interpreter.process_page(page)

    text = fake_file_handle.getvalue()
    converter.close()
    fake_file_handle.close()

    return text


def save_text_to_file(text: str, output_path: str) -> None:
    """
    テキストをファイルに保存します。

    Args:
        text (str): 保存するテキスト。
        output_path (str): テキストファイルの出力先パス。
    """
    with open(output_path, "w", encoding="utf-8") as text_file:
        text_file.write(text)


def main():
    pdf_path = "input/pdf/whitepaper.pdf"
    output_path = "output/whitepaper.txt"

    text = extract_text_from_pdf(pdf_path)
    save_text_to_file(text, output_path)


if __name__ == "__main__":
    main()
