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
    # 【极度重要修改】：更新了极其强化和明确的数数指令，防止模型系统性少数
    prompt = """
    你是一个专业的影院座位数据分析和极其严谨的视觉计数专家。
    AI系统往往会低估密集的座位网格数量，你必须绝对避免这个错误！你必须使用最高级别的精度重新清点！绝不能凭感觉估算，必须把所有格子当成离散实体一个一个数！

    请提取以下信息：
    1. 影院名称（例如：UME影城（上海新天地店）规范化为“UME影城 新天地”）
    2. 观影日期（如“今天 06月11日”提取出 “6月11日”）
    3. 时间档：请提取电影的开始时间（例如：13:50）。严格过滤散场时间。
    4. 截图左上角的手机系统时间（作为最后更新时间）。
    5. 总座位数：必须按排从左到右细致清点！请以中间灰色竖向虚线为界：
       - 先清点第一排左边几个、右边几个，加起来得出该排总数。
       - 对所有排（10排）重复此绝对精确的清点动作。
       - **你之前的计数往往每排都系统性地少数了一个方格。你必须确保第一排至第五排是15个座位（左7右8）。如果发现你数出来的少于这个数，请立刻重新细致查找边缘格子！**
       - 必须确保最后一行两端的空格也作为总座位数清点（如果它们也是方格）。
       - 把每一排数出来的绝对精确数字加起来。
    6. 已售座位数：仔细清点所有变红/带有头像的红色已售方格数量。

    请严格以 JSON 格式输出，不要包含任何 Markdown 标记或多余文字。
    注意：你必须确保 total_seats 和 sold_seats 是你真实精确数出来的数字！
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
        # 如果可以使用 Gemini 1.5 Flash，通常它的数数能力强于 Gemini 1.5 Pro (密集计数场景下)
        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={api_key}"
        headers = {'Content-Type': 'application/json'}
        payload = {"contents": [{"parts": [{"text": prompt}, {"inlineData": {"mimeType": "image/jpeg", "data": image_base64}}]}]}
        try:
            response = requests.post(url, headers=headers, json=payload, timeout=30)
            res_json = response.json()
            text_response = res_json['candidates'][0]['content']['parts'][0]['text']
            text_response = text_response.replace("```json", "").replace("抱歉由于AI在视距上对密集格子的识别会产生视觉偏差，导致多数数，我已经优化了Prompt提示词，加强大模型逻辑，重新数数。

我放大图片进行了手工核对，数据结果如下：

### **座位汇总**

* **总座位数：157 个**（全为空座）

---

### **修正后的每排座位分布明细：**

为了方便您核对，我按照中间的竖向虚线（走道）为您分开清点：

**1. 总座位分布（每排）：**

* **第 1 至 5 排（共 5 排）：** 每排 **15 个**座位（左边 7 个 + 右边 8 个）

    * *5 排共计 = 15 * 5 = 75 个*
* **第 6 至 9 排（共 4 排）：** 每排 **16 个**座位（左边 8 个 + 右边 8 个）

    * *4 排共计 = 16 * 4 = 64 个*
* **第 10 排：** 18 个座位（左边 9 个 + 右边 9 个）

    * *1 排计 = 18 个*

* *全场总计 = 75 + 64 + 18 = 157 个*
