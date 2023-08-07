from barcode import Code128
from barcode.writer import ImageWriter
from pathlib import Path
import xlsxwriter
import os
import re

RESULT_DIR = Path('results')
EXCEL_DIR = RESULT_DIR.parent / "excel"


def wrapper(func):
    def save_image_to_excel(*args):
        EXCEL_DIR.mkdir(exist_ok=True)
        wb = xlsxwriter.Workbook(EXCEL_DIR / "barcode.xlsx")
        cell_format = wb.add_format()
        cell_format.set_align('center')
        cell_format.set_align('vcenter')
        ws = wb.add_worksheet(name="RESULTS")
        func(*args)
        ws.set_column("A:A", 30)

        for i, file in enumerate(sorted(RESULT_DIR.glob('*.png'), key=os.path.getmtime), start=1):
            ws.set_row(i-1, 80)
            ws.write(f"A{i}", file.stem, cell_format)
            ws.insert_image(f"B{i}", file, {"x_scale": 0.5, "y_scale": 0.5})
        wb.close()

        for file in RESULT_DIR.glob("*.png"):
            file.unlink()
    return save_image_to_excel


@wrapper
def create_barcode_image(end_number, serial_number):
    digits = re.compile(r'(\d+)$')
    end_number = int(end_number)
    numbers = digits.search(serial_number).group(1)
    number_len = len(numbers)
    start_number = int(numbers)
    RESULT_DIR.mkdir(exist_ok=True)

    for _ in range(start_number, start_number + end_number):
        y = serial_number[:-number_len] + str(start_number)
        my_code = Code128(y.upper(), writer=ImageWriter())
        my_code.save(RESULT_DIR / f"{y}", {'write_text': False})
        start_number += 1


if __name__ == "__main__":
    create_barcode_image()
