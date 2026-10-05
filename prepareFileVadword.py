import time
import pandas as pd
import streamlit as st

st.title("Offensive Words to Migration File Converter")

# 1. Input field: Upload file
uploaded_file = st.file_uploader(
    "Upload file (เช่น Offensive words by Pu Ops_19Jun2026.xlsx)",
    type=["xlsx"],
)

# 2. Input field: Sprint
sprint_input = st.text_input("Sprint (เช่น SP20)", value="SP20")

if uploaded_file is not None:
  if st.button("Convert File"):
    try:
      # อ่านข้อมูลจากไฟล์ Excel ที่อัปโหลด
      raw_df = pd.read_excel(uploaded_file, sheet_name="Sheet1", header=None)

      # ดึงคอลัมน์ Keywords (index 1) และ exclude_keyword (index 3)
      keywords = raw_df.iloc[2:, 1].values
      exclude_keywords = raw_df.iloc[2:, 3].values

      # จัดรูปแบบ exclude_keyword โดยแทนที่ช่องว่างด้วยเซมิโคลอน (;)
      processed_exclude = []
      for val in exclude_keywords:
        if pd.isna(val):
          processed_exclude.append(pd.NA)
        else:
          processed_exclude.append(str(val).replace(" ", ";"))

      mig_df = pd.DataFrame(
          {"Keywords": keywords, "exclude_keyword": processed_exclude}
      )
      mig_df = mig_df.dropna(subset=["Keywords"]).reset_index(drop=True)

      # อ่านชีต Fang จากไฟล์ต้นฉบับ (ถ้ามี)
      try:
        fang_df = pd.read_excel(uploaded_file, sheet_name="Fang")
      except Exception:
        fang_df = pd.DataFrame()

      # สร้างชื่อไฟล์ output ตามรูปแบบ Migration_{sprint}-{unixtimestamp}
      timestamp = int(time.time())
      sprint_clean = sprint_input.strip() or "SP01"
      output_filename = f"Migration_{sprint_clean}-{timestamp}.xlsx"

      # บันทึกไฟล์ Excel
      with pd.ExcelWriter(output_filename, engine="openpyxl") as writer:
        mig_df.to_excel(writer, sheet_name="Sheet1", index=False)
        if not fang_df.empty:
          fang_df.to_excel(writer, sheet_name="Fang", index=False)

      st.success(f"แปลงไฟล์สำเร็จ: {output_filename}")

      with open(output_filename, "rb") as f:
        st.download_button(
            label="Download Converted File",
            data=f,
            file_name=output_filename,
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )
    except Exception as e:
      st.error(f"เกิดข้อผิดพลาดในการประมวลผล: {e}")
