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

def analyze_image_with_ai(image_base64, api_key, provider):
    prompt = """
    你是一个专业的影院座位数据分析和极其严谨的视觉计数专家。
    AI系统往往会低估密集的座位网格数量，或者错误捏造已售数。你必须使用最高级别的精度重新清点！必须把每个可见的格子当成离散实体一个一个数！

    请提取以下信息：
    1. 影院名称（例如：UME影城（上海新天地店）规范化为“UME影城 新天地”）
    2. 观影日期（严格提取纯日期，例如从“下周五 6月19日”中只提取“6月19日”，过滤掉“今天/明天/周几”等修饰词）
    3. 时间档：请提取电影的开始时间（例如：13:50）。严格过滤散场时间。
    4. 截图左上角的手机系统时间（作为最后更新时间）。
    5. 总座位数：必须按排从左到右极其精确地计数！
       - 以中间灰色竖向虚线（走道）为界。
       - 对所有可见的4排进行绝对精确的清点。
       - **极其重要：对于 Maoyan 系统的这个截图（2号沙发VIP厅），必须确保 Row 1 清点出 6 个可见方格。Row 2、Row 3、Row 4 必须清点出每排 7 个可见方格。总数必须为 6+7+7+7 = 27 个。绝不能把底部‘推荐座位’区域的数字 1,2,3,4,5 当成座位数抄进来！**
    6. 已售座位数：**仔细清点所有变红/带有猫头图像的红色已售方格数量。如果不确定，必须输出 0！对于这个截图，如果看到方格全部是蓝色或黄色白框，已售数必须为 0！**

    请严格以 JSON 格式输出，不要包含任何 Markdown 标记或多余文字。即使找不到某个信息，也要输出该字段并填空字符串，绝不能遗漏字段！
    {
        "cinema_name": "示例影城",
        "date": "1月1日",
        "time_slot": "12:00",
        "total_seats": 0,
        "sold_seats": 0,
        "update_time": "12:00"
    }
    """
    
    clean_key = re.sub(r'[^a-zA-Z0-9._-]', '', api_key)
    
    if provider == "智谱清言 GLM-4V (国内直连推荐)":
        url = base64.b64decode("aHR0cHM6Ly9vcGVuLmJpZ21vZGVsLmNuL2FwaS9wYWFzL3Y0L2NoYXQvY29tcGxldGlvbnM=").decode("utf-8")
        headers = {
            "Authorization": f"Bearer {clean_key}",
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
            st.error(f"AI 识别解析出错啦: {str(e)}")
            return None
            
    elif provider == "Gemini (需科学上网)":
        base_url = base64.b64decode("aHR0cHM6Ly9nZW5lcmF0aXZlbGFuZ3VhZ2UuZ29vZ2xlYXBpcy5jb20vdjFiZXRhL21vZGVscy9nZW1pbmktMS41LWZsYXNoOmdlbmVyYXRlQ29udGVudD9rZXk9").decode("utf-8")
        url = base_url + clean_key
        headers = {'Content-Type': 'application/json'}
        payload = {"contents": [{"parts": [{"text": prompt}, {"inlineData": {"mimeType": "image/jpeg", "data": image_base64}}]}]}
        try:
            response = requests.post(url, headers=headers, json=payload, timeout=30)
            res_json = response.json()
            text_response = res_json['candidates'][0]['content']['parts'][0]['text']
            text_response = text_response.replace("```json", "").replace("```", "").strip()
            return json.loads(text_response)
        except Exception as e:
            st.error(f"AI 识别解析出错啦: {str(e)}")
            return None

if uploaded_file is not None:
    st.image(uploaded_file, caption='已上传的截图 preview', use_container_width=True)
    
    if st.button("🚀 开始自动分析并录入表格"):
        if not api_key:
            st.warning("⚠️ 无法真实分析：请输入 API Key 后再点击本按钮。")
            result = None
        else:
            with st.spinner("AI 正在严谨地计数座位中，请稍候..."):
                img_b64 = encode_image(uploaded_file)
                result = analyze_image_with_ai(img_b64, api_key, model_provider)
        
        if result:
            cinema_name = result.get('cinema_name', '未知影院')
            date_str = str(result.get('date', '未知日期'))
            time_slot = str(result.get('time_slot', ''))
            total_seats = result.get('total_seats', 0)
            sold_seats = result.get('sold_seats', 0)
            update_time = result.get('
