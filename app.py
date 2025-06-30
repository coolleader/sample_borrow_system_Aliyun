import streamlit as st
import pandas as pd
import oss2
import io
from datetime import datetime
import openpyxl
import random

# 阿里云 OSS 配置（从 secrets 读）
ACCESS_KEY_ID = st.secrets["oss_access_key_id"]
ACCESS_KEY_SECRET = st.secrets["oss_access_key_secret"]
ENDPOINT = st.secrets["oss_endpoint"]
BUCKET_NAME = st.secrets["oss_bucket_name"]

# 建立连接
auth = oss2.Auth(ACCESS_KEY_ID, ACCESS_KEY_SECRET)
bucket = oss2.Bucket(auth, ENDPOINT, BUCKET_NAME)

# 文件名
OSS_FILE = "sample_inventory.xlsx"

# 读取 Excel（每次加随机数防止意外缓存）
def load_data():
    st.info(f"🔄 正在从 OSS 加载最新文件... (ID: {random.randint(1, 999999)})")
    try:
        result = bucket.get_object(OSS_FILE)
        df = pd.read_excel(io.BytesIO(result.read()), dtype=str)
    except:
        df = pd.DataFrame(columns=[
            '型号', '序列号', '料号', '样品快递号', '状态',
            '送出时间', '送出客户', '送出附件',
            '收货时间', '收货快递号', '归还附件'
        ])
    return df.fillna('')

# 保存 Excel
def save_data(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="样品数据")
        ws = writer.sheets["样品数据"]
        for col in ['B', 'C', 'D', 'J']:
            for cell in ws[col]:
                cell.number_format = '@'
    output.seek(0)
    bucket.put_object(OSS_FILE, output.getvalue())

# ================================
st.title("📦 样品送存管理系统")

menu = ["样品登记", "送出样品", "归还样品", "当前状态", "删除样品"]
choice = st.radio("选择操作", menu)

# ================================
if choice == "样品登记":
    st.header("📄 样品登记")
    df = load_data()

    sample_type = st.text_input("型号").strip()
    sample_id = st.text_input("序列号").strip()
    sample_material = st.text_input("料号").strip()
    sample_deliver_id = st.text_input("样品快递号").strip()

    if st.button("登记"):
        if sample_id and sample_id not in df['序列号'].astype(str).values:
            new_row = pd.DataFrame([{
                '型号': sample_type,
                '序列号': sample_id,
                '料号': sample_material,
                '样品快递号': sample_deliver_id,
                '状态': '在库',
                '送出时间': '', '送出客户': '', '送出附件': '',
                '收货时间': '', '收货快递号': '', '归还附件': ''
            }])
            df = pd.concat([df, new_row], ignore_index=True)
            save_data(df)
            st.success("✅ 样品已登记")
        else:
            st.warning("⚠️ 序列号为空或已存在")

# ================================
elif choice == "送出样品":
    st.header("📤 送出样品")
    df = load_data()

    sid = st.text_input("序列号").strip()
    client = st.text_input("送出客户").strip()
    send_attach = st.text_input("送出附件").strip()

    if st.button("确认送出"):
        if sid in df['序列号'].astype(str).values:
            idx = df[df['序列号'].astype(str) == sid].index[0]
            if df.at[idx, '状态'] == '在库':
                df.at[idx, '状态'] = '送出'
                df.at[idx, '送出时间'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                df.at[idx, '送出客户'] = client
                df.at[idx, '送出附件'] = send_attach
                df.at[idx, '收货时间'] = ''
                df.at[idx, '收货快递号'] = ''
                df.at[idx, '归还附件'] = ''
                save_data(df)
                st.success("✅ 样品送出成功")
            else:
                st.warning("⚠️ 样品不是在库状态")
        else:
            st.warning("⚠️ 样品不存在")

# ================================
elif choice == "归还样品":
    st.header("📥 归还样品")
    df = load_data()

    sid = st.text_input("序列号").strip()
    deliver_id = st.text_input("收货快递号").strip()
    return_attach = st.text_input("归还附件").strip()

    if st.button("确认归还"):
        if sid in df['序列号'].astype(str).values:
            idx = df[df['序列号'].astype(str) == sid].index[0]
            if df.at[idx, '状态'] == '送出':
                df.at[idx, '状态'] = '在库'
                df.at[idx, '收货时间'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                df.at[idx, '收货快递号'] = deliver_id
                df.at[idx, '归还附件'] = return_attach
                save_data(df)
                st.success("✅ 样品已归还")
            else:
                st.warning("⚠️ 样品未送出")
        else:
            st.warning("⚠️ 样品不存在")

# ================================
elif choice == "当前状态":
    st.header("📊 当前样品状态")

    # 手动刷新按钮
    if st.button("🔁 手动刷新 OSS 文件"):
        st.session_state['force_reload'] = True

    # 如果按了按钮，就强制跑一次 load_data()
    if st.session_state.get('force_reload', False):
        df = load_data()
        st.session_state['force_reload'] = False
    else:
        df = load_data()

    st.dataframe(df.astype(str), use_container_width=True)

    # 下载 Excel
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='样品数据')
        ws = writer.sheets['样品数据']
        for col in ['B', 'C', 'D', 'J']:
            for cell in ws[col]:
                cell.number_format = '@'
    excel_buffer.seek(0)
    st.download_button(
        label="📥 下载样品表 (Excel)",
        data=excel_buffer,
        file_name="样品表.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ================================
elif choice == "删除样品":
    st.header("❌ 删除样品")
    df = load_data()

    sid = st.text_input("要删除的序列号").strip()
    confirm = st.checkbox("确认删除该样品")

    if st.button("删除"):
        if sid in df['序列号'].astype(str).values:
            if confirm:
                df = df[df['序列号'].astype(str) != sid]
                save_data(df)
                st.success("✅ 样品已删除")
            else:
                st.warning("⚠️ 请勾选确认删除")
        else:
            st.warning("⚠️ 样品不存在")
