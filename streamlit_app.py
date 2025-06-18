import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.drawing.image import Image
import os
import tempfile
import time

def process_excel(before_file_path, template_file_path, cell_positions_dict, split_column, split_method, skiprows, sheet_name_col):
    df = pd.read_excel(before_file_path, skiprows=skiprows)

    for col in df.columns:
        if df[col].dtype == 'float64':
            df[col] = df[col].fillna(0).astype('Int64')

    if split_column and split_method:
        if split_column in df.columns:
            if split_method == "Remove Numbers":
                df[split_column] = df[split_column].astype(str).str.replace(r'^\d+\s', '', regex=True)

    try:
        wb = openpyxl.load_workbook(template_file_path)
        template_sheet = wb.active
    except Exception as e:
        st.error(f"Error opening template file: {e}")
        return None

    if "Sheet1" in wb.sheetnames:
        wb.remove(wb["Sheet1"])

    template_images = []
    for image in template_sheet._images:
        temp_img_path = os.path.join(tempfile.gettempdir(), f"temp_img_{len(template_images)}.png")
        with open(temp_img_path, "wb") as img_file:
            img_file.write(image._data())
        template_images.append((temp_img_path, image.anchor))

    for _, row in df.iterrows():
        sheet_name = str(row[sheet_name_col])[:31] if sheet_name_col in df.columns else "Default"
        
        if sheet_name not in wb.sheetnames:
            new_sheet = wb.copy_worksheet(template_sheet)
            new_sheet.title = sheet_name
        else:
            new_sheet = wb[sheet_name]

        for col_name, cell_positions in cell_positions_dict.items():
            value = row[col_name] if col_name in df.columns else ""
            for cell_pos in cell_positions:
                new_sheet[cell_pos] = value

        for img_path, img_anchor in template_images:
            img = Image(img_path)
            new_sheet.add_image(img, img_anchor)

    output_path = os.path.join(tempfile.gettempdir(), "processed_excel.xlsx")
    try:
        wb.save(output_path)
    except Exception as e:
        st.error(f"Error saving the output file: {e}")
        return None

    return output_path

st.title("\U0001F4CA ใบปะหน้าปะล่ะ")
st.markdown(
    "หมายเหตุ : หากต้องการใส่ข้อมูลหลายเซลล์ ให้ใส่เครื่องหมายจุลภาค( , )ขั้นระหว่างเซลล์"
)

before_file = st.file_uploader("Upload ไฟล์ข้อมูล", type=["xlsx"])
template_file = st.file_uploader("Upload ไฟล์ Template", type=["xlsx"])
skiprows = st.number_input("ระบุจำนวนแถวที่ต้องการข้าม (skiprows)", min_value=0, value=0, step=1)

if before_file:
    df = pd.read_excel(before_file, skiprows=skiprows)
    column_options = list(df.columns)
else:
    column_options = []

selected_columns = st.multiselect("เลือกคอลัมน์ที่ต้องการดึงข้อมูล", column_options)
sheet_name_col = st.selectbox("เลือกคอลัมน์สำหรับตั้งชื่อชีต", ["(ไม่เลือก)"] + column_options)

cell_positions_dict = {}
for col in selected_columns:
    cell_positions = st.text_input(f"ระบุตำแหน่งเซลล์สำหรับ {col} ", key=col)
    if cell_positions:
        cell_positions_dict[col] = [pos.strip() for pos in cell_positions.split(",")]

split_column = st.selectbox("เลือกคอลัมน์ที่ต้องการ split", ["(ไม่เลือก)"] + column_options)
split_method = st.selectbox("เลือกวิธีการ split", ["Remove Numbers", "Other Method"])

if st.button("Generate Excel File"):
    if before_file and template_file and selected_columns and cell_positions_dict:
        temp_dir = tempfile.gettempdir()
        before_file_path = os.path.join(temp_dir, "before.xlsx")
        template_file_path = os.path.join(temp_dir, "template.xlsx")
        
        with open(before_file_path, "wb") as f:
            f.write(before_file.getbuffer())
        with open(template_file_path, "wb") as f:
            f.write(template_file.getbuffer())
        
        time.sleep(1)
        
        output_file = process_excel(before_file_path, template_file_path, selected_columns, cell_positions_dict, split_column, split_method, skiprows, sheet_name_col)
        if output_file:
            with open(output_file, "rb") as file:
                st.download_button("Download Processed Excel", file, file_name="processed_excel.xlsx")
    else:
        st.error("กรุณาอัปโหลดไฟล์ เลือกคอลัมน์ และกรอกตำแหน่งเซลล์!")
