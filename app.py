import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from datetime import date
import io

def data_process(raw_file, mapping_file):
    df = pd.read_excel(raw_file)
    mapping_df = pd.read_excel(mapping_file)

    correction_dict = dict(zip(mapping_df['wrong_address'], mapping_df['right_address']))
    valid_addresses = set(mapping_df['right_address'])

    def correct_address(addr):
        if addr in valid_addresses:
            return addr, "correct", ""
        elif addr in correction_dict:
            return correction_dict[addr], "corrected", ""
        else:
            return addr, "unknown", "未出现在映射表中，请人工确认"

    df[['corrected_address', 'status', 'remark']] = df['address'].apply(
        lambda x: pd.Series(correct_address(x))
    )

    # 写入内存中的 Excel 文件
    output = io.BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)

    # 标红 unknown
    wb = load_workbook(output)
    ws = wb.active

    red_fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")

    status_col = None
    for idx, cell in enumerate(ws[1], start=1):
        if cell.value == "status":
            status_col = idx
            break

    for row in ws.iter_rows(min_row=2):
        if row[status_col - 1].value == "unknown":
            for cell in row:
                cell.fill = red_fill

    # 重新保存到 BytesIO 并返回
    final_output = io.BytesIO()
    wb.save(final_output)
    final_output.seek(0)
    return final_output

# === Streamlit UI ===
st.title("📍 地址纠错工具")

raw_file = st.file_uploader("上传原始地址文件（需包含 'address' 列）", type=["xlsx"])
mapping_file = st.file_uploader("上传地址映射表（包含 'wrong_address', 'right_address'）", type=["xlsx"])

if raw_file and mapping_file:
    if st.button("🚀 开始处理"):
        result_file = data_process(raw_file, mapping_file)
        today_str = date.today().isoformat()
        st.success("处理完成！点击下方按钮下载结果👇")
        st.download_button(
            label="📥 下载结果文件",
            data=result_file,
            file_name=f"corrected_{today_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
