def analyze_image_with_ai(image_base64, api_key, provider):
    # 精准的边界定位提示词
    prompt = """
    你是一个专业的影院座位数据分析专家。请务必仔细看图，真实计算，绝不能捏造数据！

    【核心数数域限制 - 非常重要】：
    请你在图上先找到两个文字锚点作为边界：
    - 上边界：写有“XX厅”（例如“2号沙发VIP厅”）的那一行。
    - 下边界：写有电影版本和语言（例如“国语2D”、“英语3D”或包含时间的行）的那一行。
    
    你接下来的所有座位计数，【必须严格限制在上下边界之间】的区域！绝对不能去数上边界以上的“已售/可选”图例方块，也绝对不能数下边界以下的“推荐座位(1人/2人...)”按钮！

    【计数算法要求】：
    在这个限定区域内，请采用“行列乘法”来计算总数，避免肉眼数错：
    1. 数出这部分区域一共有几排（从上到下）。
    2. 数出每一排一共有几个座位格子（从左到右）。
    3. 总座位数 = 排数 × 每排的格子数。
    4. 已售座位数：在两行之间，仔细数出变红或带有头像的红色格子数量。

    【输出格式要求】：
    你必须直接输出符合 JSON 格式的字符串，不要包含任何 markdown 标记（如 ```json），不要包含任何前后解释的废话。
    确保字典键名完全一致：
    {
        "cinema_name": "示例影城",
        "date": "1月1日",
        "time_slot": "12:00",
        "total_seats": 0,
        "sold_seats": 0,
        "update_time": "12:00"
    }
    """
    
    text_response = "" # 初始化返回值
    
    if provider == "智谱清言 GLM-4V (国内直连推荐)":
        url = "[https://open.bigmodel.cn/api/paas/v4/chat/completions](https://open.bigmodel.cn/api/paas/v4/chat/completions)"
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
                st.error(f"API 错误反馈: {res_json['error']['message']}")
                return None
            text_response = res_json['choices'][0]['message']['content']
        except Exception as e:
            st.error(f"请求智谱 API 接口失败: {str(e)}")
            return None
            
    elif provider == "Gemini (需科学上网)":
        url = f"[https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=](https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=){api_key}"
        headers = {'Content-Type': 'application/json'}
        payload = {"contents": [{"parts": [{"text": prompt}, {"inlineData": {"mimeType": "image/jpeg", "data": image_base64}}]}]}
        try:
            response = requests.post(url, headers=headers, json=payload, timeout=30)
            res_json = response.json()
            if 'candidates' not in res_json:
                st.error(f"Gemini 未能成功返回有效内容，请检查 Key 或网络状态。完整返回: {res_json}")
                return None
            text_response = res_json['candidates'][0]['content']['parts'][0]['text']
        except Exception as e:
            st.error(f"请求 Gemini API 接口失败: {str(e)}")
            return None

    # 【核心清洗防御机制】：强行把 AI 返回的非标准文本进行提取和解析
    if text_response:
        try:
            # 1. 移除大模型可能自带的 markdown 语法外壳
            clean_text = text_response.replace("```json", "").replace("```", "").strip()
            
            # 2. 如果里面还是夹杂了废话，用正则强行把 {} 及其内部的东西抠出来
            json_match = re.search(r'\{.*\}', clean_text, re.DOTALL)
            if json_match:
                clean_text = json_match.group()
            
            return json.loads(clean_text)
        except Exception as json_err:
            # 如果解析依然失败，说明 AI 根本没吐出 JSON，把它的原话打印出来供排查原因
            st.error(f"❌ 文本转换为表格失败！AI 本次并未返回标准数据。")
            st.warning(f"💡 AI 的实际回复内容为：\n\n{text_response}")
            return None
    return None
