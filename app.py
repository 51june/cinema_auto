import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import io
import base64
import json
import requests
import re

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

def clean_url_string(url_str):
    if not url_str:
        return ""
    return re.sub(r'[^\x21-\x7E]', '', url_str)

def analyze_image_with_ai(image_base64, api_key, provider):
    # 【重磅更新】：采用 Few-Shot 范例教学法，用真实的推导逻辑死死卡住 AI 的输出格式
    prompt = """
    你是一个极其精准的影院座位结构分析机器人。请严格按照以下说明执行，绝对不要输出任何 markdown 表格，绝对不要输出多余解释！

    【核心范围约束】：
    在图中定位以下两个文本：
    - 起始线（上边界）：包含“XX厅”字样（如“2号沙发VIP厅”）的行。
    - 截止线（下边界）：包含电影制式语言（如“国语2D”或“16:15-17:48”）的行。
    你只能清点这两行夹在中间的方格矩阵。

    【严谨计算推导法】：
    1. 找到上边界和下边界。
    2. 数出这两行之间一共有多少横排。
    3. 数出每一排有多少个座位方格（包含所有灰色、黄色、红色的格子）。
    4. 总座位数 = 排数 × 每排的格子数。
    5. 已售数 = 仔细清点其中变红或带有小猫/小猫头像的红色格子。

    【请参考以下成功推理范例，你必须严格模仿这个输出格式】：
    示例输入截图特征：4排，每排7个方格，全空无红格。
    示例最终输出（你必须且只能输出类似这样的纯 JSON，严禁夹带任何废话或 Markdown 符号）：
    {
        "cinema_name": "AI CINEMA新天地东台里影城",
        "date": "6月19日",
        "time_slot": "16:15",
        "total_seats": 28,
        "sold_seats": 0,
        "update_time": "15:56"
    }

    现在请处理你看到的这张图片，严格按照上面的 JSON 键名和格式吐出数据：
    """
    
    text_response = ""
    
    if provider == "智谱清言 GLM-4V (国内直连推荐)":
        url = clean_url_string("https://open.bigmodel.cn/api/paas/v4/chat/completions")
        clean_api_key = str(api_key).strip().replace(" ", "")
        
        headers = {
            "Authorization": f"Bearer {clean_api_key}",
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
            response = requests.post(url, headers=headers, json=payload, timeout=30, proxies={"http": None, "https": None})
            res_json = response.json()
            if 'error' in res_json:
                st.error(f"API 错误反馈: {res_json['error']['message']}")
                return None
            text_response = res_json['choices'][0]['message']['content']
        except Exception as e:
            st.error(f"请求智谱 API 接口失败: {str(e)}")
            return None
            
    elif provider == "Gemini (需科学上网)":
        clean_key = str(api_key).strip().replace(" ", "")
        url = clean_url_string(f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={clean_key}")
        
        headers = {'Content-Type': 'application/json'}
        payload = {"contents": [{"parts": [{"text": prompt}, {"inlineData": {"mimeType": "image/jpeg", "data": image_base64}}]}]}
        try:
            response = requests.post(url, headers=headers, json=payload, timeout=30, proxies={"http": None, "https": None})
            res_json = response.json()
            if 'candidates' not in res_json:
                st.error(f"Gemini 未能成功返回有效内容。完整返回: {res_json}")
                return None
            text_response = res_json['candidates'][0]['content']['parts'][0]['text']
        except Exception as e:
            st.error(f"请求 Gemini API 接口失败: {str(e)}")
            return None

    if text_response:
        try:
            # 强效清洗多余外壳
            clean_text = text_response.replace("```json", "").replace("```", "").strip()
            json_match = re.search(r'\{.*\}', clean_text, re.DOTALL)
            if json_match:
                clean_text = json_match.group()
            return json.loads(clean_text)
        except Exception as json_err:
            st.error(f"❌ 文本转换为表格失败！大模型未能按规定格式返回标准数据。")
            st.warning(f"💡 AI 本次的实际回复内容为：\n\n{text_response}")
            return None
            
    return None

if uploaded_file is not None:
    st.image(uploaded_file, caption='已上传的截图', use_container_width=True)
    
    if st.button("🚀 开始自动分析并录入表格"):
        if not api_key:
            st.warning("⚠️ 无法真实分析：请输入 API Key 后再点击本按钮。")
            result = None
        else:
            with st.spinner("AI 正在使用范例推理算法定位数数中..."):
                img_b64 = encode_image(uploaded_file)
                result = analyze_image_with_ai(img_b64, api_key, model_provider)
        
        if result:
            cinema_name = result.get('cinema_name', '未知影院')
            date_val = result.get('date', '未知日期')
            raw_time = str(result.get('time_slot', ''))
            total_seats = result.get('total_seats', 0)
            sold_seats = result.get('sold_seats', 0)
            update_time = result.get('update_time', '未识别')
            
            time_match = re.search(r'\d{1,2}:\d{2}', raw_time)
            if time_match:
                time_slot = time_match.group()
            else:
                time_slot = raw_time if raw_time else "未知时间"
                
            st.success("🎉 数据处理成功！")
            
            df = st.session_state['excel_data'].copy()
            exact_match = (df['影院名称'] == cinema_name) & (df['日期'] == date_val) & (df['时间档'] == time_slot)
            
            if exact_match.any():
                idx = df[exact_match].index[0]
                df.loc[idx, '总座位数'] = total_seats
                df.loc[idx, '已售'] = sold_seats
                df.loc[idx, '最后更新时间'] = update_time
                st.info(f"🔄 检测到相同场次，已自动为您覆盖更新实时数据。")
            else:
                new_row = {
                    "影院名称": cinema_name, 
                    "日期": date_val, 
                    "时间档": time_slot, 
                    "总座位数": total_seats, 
                    "已售": sold_seats, 
                    "最后更新时间": update_time
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
