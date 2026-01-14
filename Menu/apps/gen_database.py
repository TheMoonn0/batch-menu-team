# -*- coding: utf-8 -*-
import pandas as pd
import warnings
import streamlit as st
from io import BytesIO

# --- ปิด Warning กวนใจ ---
pd.options.mode.chained_assignment = None
warnings.simplefilter(action="ignore", category=FutureWarning)


def merge_excel_to_parquet_bytes(uploaded_file) -> tuple[bytes, dict]:
    """
    อ่าน Excel ทุกชีท -> ทำความสะอาด -> concat -> export parquet เป็น bytes
    คืนค่า: (parquet_bytes, stats_dict)
    """
    all_sheets = pd.read_excel(uploaded_file, sheet_name=None, dtype=str)

    all_data_frames = []
    per_sheet_rows = {}

    for sheet_name, df in all_sheets.items():
        # 1) ลบคอลัมน์ Unnamed
        df = df.loc[:, ~df.columns.astype(str).str.contains(r"^Unnamed")].copy()

        # 2) ลบแถวว่างทั้งหมด
        df = df.dropna(how="all")

        per_sheet_rows[sheet_name] = len(df)

        if len(df) > 0:
            all_data_frames.append(df)

    if not all_data_frames:
        raise ValueError("ไม่พบข้อมูลหลังทำความสะอาด (ทุกชีทว่างหรือถูกลบหมด)")

    merged_df = pd.concat(all_data_frames, ignore_index=True)

    # export parquet เป็น bytes
    buffer = BytesIO()
    merged_df.to_parquet(buffer, index=False, engine="pyarrow")
    parquet_bytes = buffer.getvalue()

    stats = {
        "sheet_count": len(all_sheets),
        "merged_rows": len(merged_df),
        "per_sheet_rows": per_sheet_rows,
        "parquet_size_mb": len(parquet_bytes) / (1024 * 1024),
    }
    return parquet_bytes, stats


def render():
    """Render GEN File Database (Excel → Parquet Merger) page"""
    
    st.write("อัปโหลดไฟล์ Excel แล้วระบบจะรวมทุกชีทเป็นไฟล์ Parquet (ลบคอลัมน์ Unnamed และลบแถวว่าง)")

    uploaded = st.file_uploader("อัปโหลดไฟล์ Excel (.xlsx)", type=["xlsx"])

    output_name = st.text_input("ชื่อไฟล์ Parquet ที่จะดาวน์โหลด", value="merged_data.parquet")

    if uploaded:
        st.info(f"ไฟล์ที่อัปโหลด: {uploaded.name}")

        col1, col2 = st.columns([1, 1])
        with col1:
            run_btn = st.button("▶️ เริ่มรวมและแปลง", use_container_width=True)
        with col2:
            preview = st.checkbox("แสดงตัวอย่างข้อมูล (Preview)", value=False)

        if run_btn:
            with st.spinner("กำลังอ่านไฟล์ / รวมข้อมูล / บันทึกเป็น Parquet..."):
                try:
                    parquet_bytes, stats = merge_excel_to_parquet_bytes(uploaded)

                    st.success("เสร็จสมบูรณ์!")

                    st.subheader("สรุปผล")
                    st.write(f"- จำนวนชีททั้งหมด: **{stats['sheet_count']}**")
                    st.write(f"- จำนวนแถวรวมหลังทำความสะอาด: **{stats['merged_rows']}**")
                    st.write(f"- ขนาดไฟล์ Parquet: **{stats['parquet_size_mb']:.2f} MB**")

                    with st.expander("รายละเอียดจำนวนแถวต่อชีท"):
                        st.write(stats["per_sheet_rows"])

                    # download
                    st.download_button(
                        label="⬇️ ดาวน์โหลดไฟล์ Parquet",
                        data=parquet_bytes,
                        file_name=output_name if output_name.strip() else "merged_data.parquet",
                        mime="application/octet-stream",
                        use_container_width=True,
                    )

                    if preview:
                        df_preview = pd.read_parquet(BytesIO(parquet_bytes), engine="pyarrow")
                        st.subheader("Preview (หัวตาราง 200 แถวแรก)")
                        st.dataframe(df_preview.head(200), use_container_width=True)

                except Exception as e:
                    st.error(f"เกิดข้อผิดพลาด: {e}")
                    st.exception(e)
    else:
        st.caption("หมายเหตุ: ไฟล์ที่มีหลายชีทจะถูกรวมเป็นตารางเดียวโดยต่อแถว (append) ทุกชีทเข้าด้วยกัน")
