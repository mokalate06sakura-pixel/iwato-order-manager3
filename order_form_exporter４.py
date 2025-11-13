import io
import zipfile
import pandas as pd
import xlsxwriter
from datetime import datetime


HEADER_STYLE = {
    'font_name': 'MS Gothic',
    'font_size': 18,
    'bold': True,
    'align': 'center',
    'valign': 'vcenter',
    'border': 1,
    'text_wrap': True
}

BODY_STYLE = {
    'font_name': 'MS Gothic',
    'font_size': 11,
    'valign': 'vcenter',
    'border': 1,
    'text_wrap': True
}

COL_WIDTHS = [15.18, 60.09] + [15.18]*10


def format_order_sheet(writer, sheetname, df, supplier, facility):
    ws = writer.book.add_worksheet(sheetname)

    # Set column widths and zoom
    for col_idx, width in enumerate(COL_WIDTHS):
        ws.set_column(col_idx, col_idx, width)
    ws.set_zoom(90)

    # Title and header
    title = f"注文書（{facility}）"
    ws.write("B1", title, writer.book.add_format({
        'font_name': 'MS Gothic', 'font_size': 26, 'bold': True,
        'align': 'left', 'valign': 'vcenter'}))

    ws.write("K2", "(有) ハートミール", writer.book.add_format({
        'font_name': 'MS Gothic', 'font_size': 24, 'bold': True,
        'align': 'right', 'valign': 'vcenter'}))

    ws.write("A3", f"{supplier} 御中", writer.book.add_format({
        'font_name': 'MS Gothic', 'font_size': 28, 'bold': True,
        'align': 'left', 'valign': 'bottom'}))

    # Write headers
    for col_idx, col in enumerate(df.columns):
        ws.write(5, col_idx, col, writer.book.add_format(HEADER_STYLE))

    # Write data
    prev_date = None
    for row_idx, row in enumerate(df.itertuples(index=False), start=6):
        for col_idx, val in enumerate(row):
            display_val = val
            if df.columns[col_idx] == "使用日":
                if prev_date == val:
                    display_val = ""
                else:
                    prev_date = val
            ws.write(row_idx, col_idx, display_val, writer.book.add_format(BODY_STYLE))
        ws.set_row(row_idx, 30)


def generate_order_zip(df, facility_label):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        for supplier, group in df.groupby("仕入先"):
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                format_order_sheet(writer, "注文書", group.drop("仕入先", axis=1), supplier, facility_label)
            filename = f"注文書_{supplier}_{facility_label}.xlsx"
            zipf.writestr(filename, buffer.getvalue())
    zip_buffer.seek(0)
    return zip_buffer.getvalue()


# Usage example (outside Streamlit):
# df = pd.read_excel("your_input.xlsx")
# df = df.sort_values(by=["使用日", "食品名"])
# zip_bytes = generate_order_zip(df, "ユーハウスいわと")
# with open("output.zip", "wb") as f:
#     f.write(zip_bytes)
