import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import io
import base64
import json
import requests
import re  # 【新增】：引入正则匹配库，用来强行切断时间

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
model_provider = st.sidebar.selectbox("选择AI模型供应商", ["智谱清言 GLM-4V (国内直连推荐)", "Gemini (需科学上网)"])

if 'excel_data' not in st.session_state:
    df_init = pd.DataFrame(columns=["影院名称", "日期", "时间档", "总座位数", "已售", "最后更新时间"])
    st.session_state['excel_data'] = df_init

uploaded_file = st.file_uploader("📸 拍照或选择手机相册中的截图", type=["jpg", "jpeg", "png"])

def encode_image(uploaded_file):
    return base64.b64encode(uploaded_file.read()).decode('utf-8')

def analyze_image_with_ai(image_base64, api_key, provider):
    # 修改了数数方式的提示，使其更清晰
    prompt = """
    你是一个专业的影院座位数据分析专家。请务必仔细看图，真实计算，绝不能捏造数据！

    你的核心任务是从座位图中准确数出灰色和红色的方格数量。请采用极其严谨的、从左到右、从上到下的顺序进行清点，以确保不会漏数或多数。

    请提取以下信息：
    1. 影院名称（例如：UME影城（上海新天地店）规范化为“UME影城 新天地”）
    2. 观影日期（如“今天 06月11日”提取出 “6月11日”）
    3. 时间档：请提取电影的开始时间（例如：13:50）。
    4. 截图左上角的手机系统时间（作为最后更新时间，例如“13:46”）
    5. 总座位数：必须重新清点！请以中间灰色竖向虚线为界，极其仔细地清点整个座位图中的所有灰色方格。请确保按照排和列，从左到右逐个清点。
    6. 已售座位数：请极其仔细地清点整个座位图中的所有变红/带有头像的红色已售方格。请确保按照排和列，从左到右逐个清点。

    请严格以 JSON 格式输出，不要包含任何 Markdown 标记或多余文字。
    注意：下面示例中的数字 0 仅为格式参考，你必须输出你真实数出来的总座位数和已售数！
    {
        "cinema_name": "示例影城",
        "date": "1月1日",
        "time_slot": "12:00",
        "total_seats": 0,
        "sold_seats": 0,
        "update_time": "12:00"
    }
    """
    
    if provider == "智谱清言 GLM-4V (国内直连推荐)":
        url = "https://open.bigmodel.cn/api/paas/v4/chat/completions"
        headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json"
        }
        payload = {
            "model": "glm-4v",
            "messages": [
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                        {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{image_base64}"}}
                    ]
                }
            ]
        }
        try:
            response = requests.post(url, headers=headers, json=payload, timeout=30)
            res_json = response.json()
            if 'error' in res_json:
                st.error(f"API 错误: {res_json['error']['message']}")
                return None
            text_response = res_json['choices'][0]['message']['content']
            text_response = text_response.replace("```json", "").replace("```", "").strip()
            return json.loads(text_response)
        except Exception as e:
            st.error(f"AI 识别出错啦: {str(e)}")
            return None
            
    elif provider == "Gemini (需科学上网)":
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

if uploaded_file is not None:
    st.image(uploaded_file, caption='已上传的截图', use_container_width=True)
    
    if st.button("🚀 开始自动分析并录入表格"):
        if not api_key:
            st.warning("⚠️ 无法真实分析：请输入 API Key 后再点击本按钮。")
            result = None
        else:
            with st.spinner("AI 正在严谨地查数座位中，请稍候..."):
                img_b64 = encode_image(uploaded_file)
                result = analyze_image_with_ai(img_b64, api_key, model_provider)
        
        if result:
            # 【代码已更新】：强制用 Python 正则表达式截取时间档，确保万无一失
            raw_time = str(result.get('time_slot', ''))
            time_match = re.search(r'\d{1,2}:\d{2}', raw_time)
            if time_match:
                result['time_slot'] = time_match.group()
                
            st.success("🎉 数据处理成功！")
            
            df = st.session_state['excel_data'].copy()
            
            exact_match = (df['影院名称'] == result['cinema_name']) & (df['日期'] == result['date']) & (df['时间档'] == result['time_slot'])
            
            if exact_match.any():
                idx = df[exact_match].index[0]
                df.loc[idx, '总座位数'] = result['total_seats']
                df.loc[idx, '已售'] = result['sold_seats']
                df.loc[idx, '最后更新时间'] = result['update_time']
                st.info(f"🔄 检测到相同场次，已自动为您覆盖更新实时数据。")
            else:
                new_row = {
                    "影院名称": result['cinema_name'], 
                    "日期": result['date'], 
                    "时间档": result['time_slot'], 
                    "总座位数": result['total_seats'], 
                    "已售": result['sold_seats'], 
                    "最后更新时间": result['update_time']
                }
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
    
    headers = ["影院名称(预售开启后更新影院列表)", "日期", "时间档", "总座位数", "已售", "占比", "最后更新时间(精确到分)"]
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
