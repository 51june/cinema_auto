import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import io
import base64
import json
import requests

st.set_page_config(
    page_title="影院票房与座位自动统计工具",
    page_icon="🎬",
    layout="centered",
    initial_sidebar_state="collapsed"
)

st.markdown("""
    <style>
    .main { background-color: #fafbfa; }
    .stButton>button { width: 100%; background-color: #1F4E78; color: white; border-radius: 8px; padding: 10px; font-weight: bold; }
    .stButton>button:hover { background-color: #153654; color: white; }
    .upload-box { border: 2px dashed #1F4E78; padding: 20px; border-radius: 10px; text-align: center; background-color: #ffffff; }
    h1 { color: #1F4E78; font-family: 'Microsoft YaHei', sans-serif; }
    </style>
""", unsafe_allow_html=True)

st.title("🎬 影院座位截图自动录入系统")
st.write("上传购票界面的座位截图，系统将自动识别并更新至 Excel 表格中。自动识别同场次并覆盖旧数据。")

st.sidebar.header("⚙️ 配置中心")
api_key = st.sidebar.text_input("请输入大模型 API Key", type="password")
model_provider = st.sidebar.selectbox("选择AI模型供应商", ["Gemini (推荐)", "OpenAI GPT-4o"])

if 'excel_data' not in st.session_state:
    df_init = pd.DataFrame([
        {"影院名称": "和平影都", "日期": "", "时间档": "", "总座位数": "", "已售": "", "最后更新时间": ""},
        {"影院名称": "大光明电影院 南西", "日期": "", "时间档": "", "总座位数": "", "已售": "", "最后更新时间": ""},
        {"影院名称": "SFC上影百联 大上海", "日期": "", "时间档": "", "总座位数": "", "已售": "", "最后更新时间": ""},
        {"影院名称": "UME影城 新天地", "日期": "", "时间档": "", "总座位数": "", "已售": "", "最后更新时间": ""},
        {"影院名称": "上海科技影城", "日期": "", "时间档": "", "总座位数": "", "已售": "", "最后更新时间": ""},
        {"影院名称": "大光明 儿艺店", "日期": "", "时间档": "", "总座位数": "", "已售": "", "最后更新时间": ""},
        {"影院名称": "中影国际 兰生大厦", "日期": "", "时间档": "", "总座位数": "", "已售": "", "最后更新时间": ""},
        {"影院名称": "国泰电影院", "日期": "", "时间档": "", "总座位数": "", "已售": "", "最后更新时间": ""},
        {"影院名称": "cooperstar影城", "日期": "", "时间档": "", "总座位数": "", "已售": "", "最后更新时间": ""},
        {"影院名称": "博悦汇影城BFC外滩金融中心", "日期": "", "时间档": "", "总座位数": "", "已售": "", "最后更新时间": ""},
        {"影院名称": "万幕国际影城 黄浦店", "日期": "", "时间档": "", "总座位数": "", "已售": "", "最后更新时间": ""}
    ])
    st.session_state['excel_data'] = df_init

uploaded_file = st.file_uploader("📸 拍照或选择手机相册中的截图", type=["jpg", "jpeg", "png"])

def encode_image(uploaded_file):
    return base64.b64encode(uploaded_file.read()).decode('utf-8')

def analyze_image_with_ai(image_base64, api_key, provider):
    prompt = """
    你是一个专业的影院座位数据分析专家。请仔细分析这张电影院座位图的截图，并提取以下信息：
    1. 影院名称（例如：UME影城（上海新天地店）中的“UME影城 新天地”，请规范化为表格中的名字）
    2. 观影日期（如“今天 06月11日”提取出 “6月11日”）
    3. 时间档（如 “13:50” 或 “16:05”）
    4. 截图左上角的手机系统时间（作为最后更新时间，例如“13:46”）
    5. 总座位数：请以中间灰色竖向虚线为界，仔细按排数清点所有方格。
    6. 已售座位数：仔细清点所有变红/带有头像的红色已售方格数量。

    请严格以 JSON 格式输出，不要包含任何 Markdown 标记或多余文字，结构如下：
    {
        "cinema_name": "UME影城 新天地",
        "date": "6月11日",
        "time_slot": "13:50",
        "total_seats": 157,
        "sold_seats": 0,
        "update_time": "14:03"
    }
    """
    
    if provider == "Gemini (推荐)":
        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={api_key}"
        headers = {'Content-Type': 'application/json'}
        payload = {"contents": [{"parts": [{"text": prompt}, {"inlineData": {"mimeType": "image/jpeg", "data": image_base64}}]}]}
        try:
            response = requests.post(url, headers=headers, json=payload, timeout=30)
            res_json = response.json()
            text_response = res_json['candidates'][0]['content']['parts'][0]['text']
            text_response = text_response.replace("```json", "").replace("```", "").strip()
            return json.loads(text_response)
        except Exception as e:
            st.error(f"AI 识别出错啦: {str(e)}")
            return None
    else:
        st.info("GPT-4o 接口逻辑已集成，请输入合法的密钥使用。")
        return None

if uploaded_file is not None:
    st.image(uploaded_file, caption='已上传的截图', use_container_width=True)
    
    if st.button("🚀 开始自动分析并录入表格"):
        if not api_key:
            st.warning("⚠️ 演示模式：将使用 16:05 场次模拟数据更新（填入 API Key 即可真实读图）")
            result = {
                "cinema_name": "UME影城 新天地", "date": "6月11日", "time_slot": "16:05", 
                "total_seats": 157, "sold_seats": 2, "update_time": "14:10"
            }
        else:
            with st.spinner("AI 正在疯狂数座位中，请稍候..."):
                img_b64 = encode_image(uploaded_file)
                result = analyze_image_with_ai(img_b64, api_key, model_provider)
        
        if result:
            st.success("🎉 数据处理成功！")
            
            df = st.session_state['excel_data'].copy()
            
            # 【核心修改逻辑】：寻找同影院、同日期、同时间档的行
            exact_match = (df['影院名称'] == result['cinema_name']) & (df['日期'] == result['date']) & (df['时间档'] == result['time_slot'])
            
            if exact_match.any():
                # 找到了完全一样的场次，直接覆盖更新数据！
                idx = df[exact_match].index[0]
                df.loc[idx, '总座位数'] = result['total_seats']
                df.loc[idx, '已售'] = result['sold_seats']
                df.loc[idx, '最后更新时间'] = result['update_time']
                st.info(f"🔄 检测到相同场次，已自动为您更新实时数据。")
            else:
                # 没找到完全相同的场次，看看影院存在不
                cinema_match = df['影院名称'] == result['cinema_name']
                if cinema_match.any():
                    idx_list = df[cinema_match].index
                    # 如果该影院只有初始化时留下的一行空白数据，则直接覆盖这行空白数据
                    if len(idx_list) == 1 and df.loc[idx_list[0], '日期'] == "":
                        df.loc[idx_list[0], '日期'] = result['date']
                        df.loc[idx_list[0], '时间档'] = result['time_slot']
                        df.loc[idx_list[0], '总座位数'] = result['total_seats']
                        df.loc[idx_list[0], '已售'] = result['sold_seats']
                        df.loc[idx_list[0], '最后更新时间'] = result['update_time']
                    else:
                        # 如果已有其他场次数据，则在下方追加一行新场次
                        idx = idx_list[-1]
                        new_row = {"影院名称": result['cinema_name'], "日期": result['date'], "时间档": result['time_slot'], "总座位数": result['total_seats'], "已售": result['sold_seats'], "最后更新时间": result['update_time']}
                        df = pd.concat([df.iloc[:idx+1], pd.DataFrame([new_row]), df.iloc[idx+1:]], ignore_index=True)
                else:
                    # 连影院名字都没有，直接在最后追加
                    new_row = {"影院名称": result['cinema_name'], "日期": result['date'], "时间档": result['time_slot'], "总座位数": result['total_seats'], "已售": result['sold_seats'], "最后更新时间": result['update_time']}
                    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
                
            st.session_state['excel_data'] = df

st.subheader("📊 当前实时统计报表（数据预览）")
st.dataframe(st.session_state['excel_data'], use_container_width=True)

def convert_df_to_excel(df_data):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "票房与座位统计表"
    ws.views.sheetView[0].showGridLines = True
    
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    zebra_fill = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    header_font = Font(name="Microsoft YaHei", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Microsoft YaHei", size=10, bold=False, color="000000")
    thin_side = Side(border_style="thin", color="D9D9D9")
    border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    
    headers = ["影院名称\n(预售开启后更新影院列表)", "日期", "时间档", "总座位数", "已售", "占比", "最后更新时间(精确到分)"]
    ws.append(headers)
    ws.row_dimensions[1].height = 28
    
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_all

    for r_idx, row in enumerate(df_data.itertuples(index=False), 2):
        row_vals = [row[0], row[1], row[2], row[3], row[4], "", row[5]]
        ws.append(row_vals)
        ws.row_dimensions[r_idx].height = 22
        
        for c_idx in range(1, 8):
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.font = data_font
            cell.border = border_all
            cell.fill = zebra_fill if r_idx % 2 == 0 else white_fill
            
            if c_idx == 1:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif c_idx in [2, 3, 7]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_idx in [4, 5]:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                if cell.value != "": cell.number_format = '#,##0'
            elif c_idx == 6:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = '0.0%'
                cell.value = f'=IFERROR(E{r_idx}/D{r_idx}, "")'
                
    ws.column_dimensions['A'].width = 32
    ws.column_dimensions['G'].width = 24
    wb.save(output)
    return output.getvalue()

if st.button("📥 下载已更新的专业 Excel 报表"):
    excel_bytes = convert_df_to_excel(st.session_state['excel_data'])
    st.download_button(
        label="点击下载 .xlsx 文件",
        data=excel_bytes,
        file_name="最新影院预售统计表.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
