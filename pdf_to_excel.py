from typing import Iterator

import io
import os
import pandas as pd
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
import tabula


def extract_text_by_page(pdf_path: str, page: int) -> Iterator[str]:
    """
    PDFファイル内の指定したページからテキストを抽出します。

    Args:
        pdf_path (str): PDFファイルへのパス。
        page (int): テキストを抽出するページ番号。

    Yields:
        str: 抽出されたテキスト。
    """
    # PDFファイルを開く
    with open(pdf_path, 'rb') as fh:
        # ページを順に処理
        for page_num, page_data in enumerate(PDFPage.get_pages(fh, caching=True, check_extractable=True)):
            # 指定したページ番号に到達したらテキスト抽出
            if page_num + 1 == page:
                resource_manager = PDFResourceManager()
                fake_file_handle = io.StringIO()
                converter = TextConverter(resource_manager, fake_file_handle, laparams=LAParams())
                page_interpreter = PDFPageInterpreter(resource_manager, converter)
                page_interpreter.process_page(page_data)
                text = fake_file_handle.getvalue()
                yield text
                converter.close()
                fake_file_handle.close()


def convert_pdf_to_excel(pdf_path: str, page: int, output_path: str) -> None:
    """
    PDFファイル内の指定したページの表をExcelファイルに変換します。

    Args:
        pdf_path (str): PDFファイルへのパス。
        page (int): 表が含まれるページ番号。
        output_path (str): Excelファイルの出力先パス。
    """
    temp_csv_path = "temp.csv"
    # PDFの表をCSVに変換
    tabula.convert_into(pdf_path, temp_csv_path, output_format="csv", pages=[page])
    # CSVファイルを読み込み
    df = pd.read_csv(temp_csv_path)
    # DataFrameをExcelファイルに変換して出力
    df.to_excel(output_path, index=False)
    # 一時的に作成されたCSVファイルを削除
    os.remove(temp_csv_path)


def main():
    pdf_path = "input/ver2/ver2_0.pdf"
    output_path = "output/ver2_5p.xlsx"
    page = 5

    text = ""
    # テキスト抽出
    for page_text in extract_text_by_page(pdf_path, page):
        text += page_text

    # テキストをファイルに書き込み
    with open(f"page{page}.txt", "w", encoding="utf-8") as text_file:
        text_file.write(text)

    # PDFの表をExcelファイルに変換
    convert_pdf_to_excel(pdf_path, page, output_path)


if __name__ == "__main__":
    main()
