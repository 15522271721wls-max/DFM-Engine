import streamlit as st
from pptx import Presentation
import io
import requests
from zhipuai import ZhipuAI
import docx 
from openai import OpenAI
import json
import streamlit.components.v1 as components

# === 网页基础设置 ===
st.set_page_config(page_title="DFM生成引擎", page_icon="⚙️", layout="centered")

# ==========================================
#           🚨 专属系统防盗门 (拦截器)
# ==========================================
def check_password():
    """验证密码的回调函数"""
    if st.session_state["password_input"] == "JINYA888":
        st.session_state["authenticated"] = True
    else:
        st.session_state["authenticated"] = False
        st.error("❌ 邀请码错误或已失效，请联系系统管理员获取！")

# 初始化会话状态
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

# 如果没有通过验证，显示登录拦截屏并停止执行后续所有代码
if not st.session_state["authenticated"]:
    st.markdown("<br><br><br>", unsafe_allow_html=True)
    st.markdown("<h1 style='text-align: center; color: #1E3A8A; font-size: 3rem;'>🔒 JINYA 智绘终端</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #6B7280; font-size: 1.2rem; margin-bottom: 2rem;'>内部专属自动化演示系统，请输入授权邀请码解锁</p>", unsafe_allow_html=True)
    
    # 居中布局密码框
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        # 输入框
        st.text_input(
            "邀请码", 
            type="password", 
            key="password_input", 
            label_visibility="collapsed", 
            placeholder="请输入邀请码 (如: 123456)"
        )
        # 验证按钮
        st.button("🔓 验证并解锁系统", on_click=check_password, use_container_width=True)
    
    # 强制停止！不让未授权访客看到后面的代码和界面
    st.stop()
# ==========================================


# === 高级 CSS 魔法 ===
st.markdown("""
<style>
    /* 1. 强制全局字体统一 */
    html, body, [class*="css"] { font-family: "PingFang SC", "Microsoft YaHei", sans-serif !important; }
    
    /* 2. 拓宽中间内容区 */
    .block-container { max-width: 900px !important; padding-top: 2rem !important; }
    
    /* 3. 高级渐变色主标题 */
    .main-title { text-align: center; font-size: 2.8rem; font-weight: 800; background: linear-gradient(90deg, #1E3A8A, #3B82F6); -webkit-background-clip: text; -webkit-text-fill-color: transparent; margin-bottom: 0.5rem; padding-top: 1rem; }
    
    /* 4. 精致副标题 */
    .sub-title { text-align: center; color: #6B7280; font-size: 1.1rem; font-weight: 500; margin-bottom: 2rem; }
    
    /* 5. 统一所有主按键样式 */
    .stButton > button, .stDownloadButton > button, div[data-testid="stFormSubmitButton"] > button { background: linear-gradient(90deg, #1E3A8A, #3B82F6) !important; color: white !important; font-size: 1.1rem !important; font-weight: bold !important; border-radius: 8px !important; border: none !important; padding: 0.6rem 0 !important; box-shadow: 0 4px 6px rgba(59, 130, 246, 0.2) !important; transition: all 0.3s ease !important; }
    .stButton > button:hover, .stDownloadButton > button:hover, div[data-testid="stFormSubmitButton"] > button:hover { transform: translateY(-2px); box-shadow: 0 6px 12px rgba(59, 130, 246, 0.4) !important; }
    
    /* 6. 美化顶部提示框 */
    .stAlert { border-radius: 12px !important; border: none !important; box-shadow: 0 2px 8px rgba(0,0,0,0.05) !important; }
    
    /* 7. 彻底隐藏输入框聚焦时的英文提示 */
    div[data-testid="InputInstructions"] { display: none !important; }
    
    /* 8. 强制下拉菜单 (Selectbox) 悬停时显示小手 */
    div[data-baseweb="select"], div[data-baseweb="select"] > div, div[data-baseweb="select"] input, div[data-baseweb="select"] svg { cursor: pointer !important; }
    
    /* 9. 强制拉宽左侧边栏，确保文字完整展示 */
    section[data-testid="stSidebar"] { min-width: 360px !important; max-width: 400px !important; }
    
    /* 10. 【终极防漏靶】精准替换整个上传组件内部的英文与按钮，兼容所有系统，支持黑白模式 */
    div[data-testid="stFileUploader"] button {
        background: linear-gradient(90deg, #1E3A8A, #3B82F6) !important;
        border: none !important;
        border-radius: 8px !important;
        color: transparent !important;
        position: relative;
    }
    div[data-testid="stFileUploader"] button::after {
        content: "浏览本地文件" !important;
        position: absolute;
        color: white !important;
        left: 50%;
        top: 50%;
        transform: translate(-50%, -50%);
        font-weight: bold;
        white-space: nowrap;
    }
    div[data-testid="stFileUploader"] button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 10px rgba(59, 130, 246, 0.3) !important;
    }
    div[data-testid="stFileUploader"] [data-testid="stMarkdownContainer"] p { font-size: 0 !important; }
    div[data-testid="stFileUploader"] [data-testid="stMarkdownContainer"] p::after {
        content: "拖拽文件至此区域" !important;
        font-size: 16px !important;
        font-weight: 600 !important;
        color: var(--text-color) !important; 
    }
    div[data-testid="stFileUploader"] small { font-size: 0 !important; }
    div[data-testid="stFileUploader"] small::after {
        content: "单文件上限 200MB (格式: DOCX, MD, TXT, PPTX)" !important;
        font-size: 13px !important;
        color: var(--text-color) !important;
        opacity: 0.6;
    }
    
    /* 11. 【优雅的跑马灯特效】让必读标题多色交替缓慢变色 */
    @keyframes rainbowFlash {
        0%   { color: #FF0000; text-shadow: 0 0 10px rgba(255, 0, 0, 0.8); }
        15%  { color: #FF7F00; text-shadow: 0 0 10px rgba(255, 127, 0, 0.8); }
        30%  { color: #FFFF00; text-shadow: 0 0 10px rgba(255, 255, 0, 0.8); }
        45%  { color: #00FF00; text-shadow: 0 0 10px rgba(0, 255, 0, 0.8); }
        60%  { color: #00FFFF; text-shadow: 0 0 10px rgba(0, 255, 255, 0.8); }
        75%  { color: #0000FF; text-shadow: 0 0 10px rgba(0, 0, 255, 0.8); }
        90%  { color: #8B00FF; text-shadow: 0 0 10px rgba(139, 0, 255, 0.8); }
        100% { color: #FF0000; text-shadow: 0 0 10px rgba(255, 0, 0, 0.8); }
    }
    div[data-testid="stExpander"] summary p {
        font-weight: 900 !important;
        font-size: 1.15rem !important;
        animation: rainbowFlash 3s linear infinite; 
    }
</style>
""", unsafe_allow_html=True)

# === 后台逻辑函数 ===
def extract_text_content(uploaded_file):
    if not uploaded_file: return ""
    try:
        if uploaded_file.name.endswith('.md') or uploaded_file.name.endswith('.txt'): return uploaded_file.getvalue().decode('utf-8')
        elif uploaded_file.name.endswith('.docx'): return '\n'.join([para.text for para in docx.Document(io.BytesIO(uploaded_file.getvalue())).paragraphs])
    except: return ""
    return ""

def parse_markdown(md_text):
    if not md_text: return []
    parsed_slides = []
    for slide in md_text.split('---'):
        lines = slide.strip().split('\n')
        if not lines or lines == ['']: continue
        l_type, title, body, ai_prompt = "正文", "", [], ""
        for line in lines:
            line = line.strip()
            if line.startswith('[') and line.endswith(']') and not line.startswith('[AI配图:'): l_type = line[1:-1]
            elif line.startswith('[AI配图:') and line.endswith(']'): ai_prompt = line[7:-1].strip()
            elif line.startswith('# '): title = line[2:]
            elif line: body.append(line)
        parsed_slides.append({'type': l_type, 'title': title, 'body': '\n'.join(body), 'ai_prompt': ai_prompt})
    return parsed_slides

def get_english_translation(chinese_title):
    title_dict = {
        "背景": "PROJECT BACKGROUND", "需求": "REQUIREMENT ANALYSIS", "目标": "PROJECT OBJECTIVES", "参数": "TECHNICAL SPECIFICATIONS", "环境": "ENVIRONMENTAL REQUIREMENTS",
        "工艺流程": "PROCESS FLOW", "工艺解析": "PROCESS ANALYSIS", "时序": "TIMING SEQUENCE", "节拍": "CYCLE TIME ANALYSIS", "产能": "PRODUCTION CAPACITY",
        "整体方案": "OVERALL SCHEME", "设备布局": "EQUIPMENT LAYOUT", "空间": "SPACE LAYOUT & ERGONOMICS", "外观": "EQUIPMENT APPEARANCE",
        "核心机构": "CORE MECHANISM DESIGN", "结构设计": "STRUCTURAL DESIGN", "模组选型": "MODULE SELECTION", "夹具": "FIXTURE DESIGN", "治具": "JIG & FIXTURE DESIGN", "升降": "LIFTING MECHANISM", "传动": "TRANSMISSION SYSTEM", "上料": "LOADING SYSTEM", "下料": "UNLOADING SYSTEM", "移载": "TRANSFER MECHANISM", "定位": "POSITIONING SYSTEM",
        "电气": "ELECTRICAL CONTROL SYSTEM", "控制": "CONTROL SYSTEM ARCHITECTURE", "气动": "PNEUMATIC SYSTEM", "视觉": "MACHINE VISION SYSTEM", "检测": "INSPECTION & TESTING", "传感器": "SENSOR CONFIGURATION",
        "性能": "PERFORMANCE METRICS", "精度": "PRECISION ANALYSIS", "受力": "FORCE & STRESS ANALYSIS", "负载": "LOAD CAPACITY",
        "风险": "RISK ASSESSMENT (FMEA)", "可行性": "FEASIBILITY ANALYSIS", "难点": "TECHNICAL DIFFICULTIES", "安全": "SAFETY & PROTECTION", "防错": "POKA-YOKE DESIGN", "防腐": "ANTI-CORROSION DESIGN", "环保": "ENVIRONMENTAL PROTECTION", "维护": "MAINTENANCE & ACCESSIBILITY",
        "计划": "PROJECT SCHEDULE", "交期": "DELIVERY SCHEDULE", "清单": "BOM & SPARE PARTS", "成本": "COST ESTIMATION", "优势": "SOLUTION ADVANTAGES",
        "结论": "DESIGN CONCLUSION", "建议": "SUGGESTIONS & NEXT STEPS", "总结": "PROJECT SUMMARY"
    }
    for k, v in title_dict.items():
        if k in chinese_title: return v
    return "CHAPTER OVERVIEW"

# === 新增：API 真实连通性验证函数 ===
def validate_api_key(api_key, provider, base_url=None):
    try:
        if "智谱" in provider:
            if "." not in api_key:
                return False, "❌ 格式错误：智谱 API Key 必须包含小数点 (格式: id.secret)。"
            client = ZhipuAI(api_key=api_key)
            # 发送一个极其微小的测试请求，探测云端连通性
            client.chat.completions.create(
                model="glm-4-flash", 
                messages=[{"role": "user", "content": "1"}],
                max_tokens=1
            )
            return True, "✅ 智谱引擎连接成功，授权已就绪！"
        else:
            client = OpenAI(api_key=api_key, base_url=base_url)
            # 向 OpenAI 或中转站请求模型列表，探测连通性
            client.models.list()
            return True, "✅ 视觉引擎连接成功，授权已就绪！"
    except Exception as e:
        err_str = str(e).lower()
        if "401" in err_str or "auth" in err_str or "invalid" in err_str or "api_key" in err_str:
            return False, "❌ 凭证无效：请检查您的 API Key 是否填写正确或已被封禁！"
        elif "url" in err_str or "connection" in err_str or "timeout" in err_str or "resolve" in err_str:
            return False, "❌ 网络错误：无法连接到接口，请检查 Base URL 是否正确填写！"
        else:
            return False, f"❌ 验证失败：{e}"

def create_ppt(parsed_slides, template_bytes, api_key, ai_provider, base_url=None, model_name=None):
    if not parsed_slides or not template_bytes: return None
    prs = Presentation(io.BytesIO(template_bytes))
    
    for i, slide_data in enumerate(parsed_slides):
        l_type = slide_data['type']
        try:
            if l_type == "封面":
                slide = prs.slides.add_slide(prs.slide_layouts[0])
                slide.placeholders[10].text = (slide_data['title'] + "\n" + slide_data['body']).strip()
            elif l_type == "目录":
                slide = prs.slides.add_slide(prs.slide_layouts[1])
                slide.placeholders[12].text = slide_data['body'].strip()
            elif l_type == "过渡":
                slide = prs.slides.add_slide(prs.slide_layouts[2])
                chinese_title = slide_data['title'].strip()
                slide.placeholders[10].text = chinese_title
                slide.placeholders[12].text = get_english_translation(chinese_title)
            elif l_type == "左文右图":
                slide = prs.slides.add_slide(prs.slide_layouts[3])
                slide.placeholders[10].text = slide_data['title']
                slide.placeholders[11].text = slide_data['body']
            else: 
                slide = prs.slides.add_slide(prs.slide_layouts[4])
                slide.placeholders[10].text = slide_data['title']
                slide.placeholders[11].text = slide_data['body']
                
            if slide_data['ai_prompt'] and api_key:
                st.toast(f"🎨 正在呼叫视觉引擎绘制: {slide_data['title']} ...", icon="🤖")
                img_url = ""
                
                try:
                    if ai_provider == "官方 智谱 AI (CogView-3)":
                        client = ZhipuAI(api_key=api_key)
                        response = client.images.generations(model="cogview-3", prompt=slide_data['ai_prompt'])
                        img_url = response.data[0].url
                    
                    elif ai_provider == "官方 OpenAI (DALL-E 3)":
                        client = OpenAI(api_key=api_key)
                        response = client.images.generate(model="dall-e-3", prompt=slide_data['ai_prompt'], n=1, size="1024x1024")
                        img_url = response.data[0].url
                        
                    elif ai_provider == "通用聚合中转站 (支持 Midjourney/DALL-E等)":
                        client = OpenAI(api_key=api_key, base_url=base_url)
                        response = client.images.generate(model=model_name, prompt=slide_data['ai_prompt'], n=1, size="1024x1024")
                        img_url = response.data[0].url
                        
                    if img_url:
                        img_stream = io.BytesIO(requests.get(img_url).content)
                        slide_width, slide_height = prs.slide_width, prs.slide_height
                        slide.shapes.add_picture(img_stream, int(slide_width * 0.55), int(slide_height * 0.25), width=int(slide_width * 0.40))
                except Exception as draw_e:
                    st.error(f"第 {i+1} 页画图失败 (请检查Key是否透支或网络异常): {draw_e}")
                    
        except Exception as e:
            st.error(f"⚠️ 生成第 {i+1} 页时遇到麻烦: {e}")
            pass 
            
    try: prs.slides.add_slide(prs.slide_layouts[5])
    except: pass

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output

# ================= UI 布局与展示 =================

try: st.sidebar.image("JINYA-logo-透明背景.png", use_container_width=True)
except: st.sidebar.markdown("### 🏢 JINYA")

st.sidebar.markdown("<br>", unsafe_allow_html=True)
st.sidebar.markdown("#### 🛡️ 开发者授权配置")

with st.sidebar.form(key='api_form'):
    ai_provider = st.selectbox(
        "🌐 选择视觉大模型引擎", 
        ["官方 智谱 AI (CogView-3)", "官方 OpenAI (DALL-E 3)", "通用聚合中转站 (支持 Midjourney/DALL-E等)"]
    )
    
    base_url = None
    model_name = None
    if ai_provider == "通用聚合中转站 (支持 Midjourney/DALL-E等)":
        st.caption("🚀 您已开启聚合模式。请填入提供商的接口地址与模型名称。")
        base_url = st.text_input("接口地址 (Base URL)", placeholder="如: https://api.xxx.com/v1")
        model_name = st.text_input("调用模型名称", placeholder="如: dall-e-3 或 midjourney", value="dall-e-3")
        
    api_key_input = st.text_input("API 凭证 (API Key)", type="password", placeholder="填入以激活自动画图功能...")
    
    # === 更新：加入真实云端连通性验证 ===
    submit_btn = st.form_submit_button("✓ 保存并验证配置", use_container_width=True)
    if submit_btn:
        if api_key_input:
            with st.spinner("🔌 正在连接云端验证凭证..."):
                is_valid, msg = validate_api_key(api_key_input, ai_provider, base_url)
                if is_valid:
                    st.success(msg)
                else:
                    st.error(msg)
        else:
            st.warning("⚠️ 请先填入 API 凭证！")

st.sidebar.markdown("""
<div style="cursor: pointer; color: #6B7280; font-size: 0.85rem; padding-top: 10px;">
    注：无需生成图片时，可不填写 API 凭证。
</div>
""", unsafe_allow_html=True)


st.markdown("<div class='main-title'>DFM生成引擎</div>", unsafe_allow_html=True)
st.markdown("<div class='sub-title'>自动化设备 DFM 极速演示方案生成器，PPT 排版一键搞定！</div>", unsafe_allow_html=True)

st.info("""
** 欢迎使用！本系统将为您自动化完成以下工作：**
1.  **多格式读取**：无缝解析 `Word文档 (.docx)`、`Markdown (.md)` 或纯文本。
2.  **智能对位排版**：自动匹配公司 PPT 模板，一键套用。
3.  **专业双语翻译**：内置自动化专属词典，毫秒级生成英文标题过渡页。
4.  **全网 AI 概念配图**：不仅支持官方智谱，更可通过中转协议一键接入 Midjourney/DALL-E 绘制概念图！
""")

st.markdown("<br>", unsafe_allow_html=True)

# === AI 提示词与一键复制按钮 ===
ai_prompt_text = """# Role (角色定位)
你是一位拥有10年经验的资深非标自动化机械工程师。你需要根据我提供的项目需求，撰写一份极其专业的《详细设计方案 (DFM)》。

# Output Format Strict Rules (输出格式绝对要求)
为了确保输出能够完美导入自动化渲染引擎，你必须【严格遵守】以下 Markdown 语法规则：

1. 分页符：每一页 PPT 之间，必须且只能用 `---` (三个减号) 独占一行来进行分隔。
2. 版式标签：每一页的开头第一行，必须指定版式，格式为 `[版式名称]`。只能从以下 5 种版式中选择：
   - `[封面]`：仅用于第一页。
   - `[目录]`：仅用于第二页，列出方案大纲。
   - `[过渡]`：用于每个大章节的开头页。
   - `[左文右图]`：用于需要展示设备结构、布局的页面。
   - `[正文]`：用于纯文字说明或数据分析页面。
3. 页面标题：紧接版式标签的下一行，必须是标题，以 `# ` 开头（注意 # 后面有一个空格）。
4. 正文内容：标题下方为正文内容。正文请精简专业，使用项目符号（* 或 -）排版，绝对不要使用复杂的 Markdown 表格或多级嵌套。
5. AI 配图指令：当版式为 `[左文右图]` 时，请在正文末尾单独起一行，生成提示词：`[AI配图: 一段详细的英文画面描述，要求工业风、3D概念图风格]`。

# 示例标准模板 (请严格仿照此格式输出全文)：
[封面]
# 自动化设备详细设计方案 (DFM)
项目编号：TSDE-2026
---
[目录]
# 方案目录
* 一、项目背景与需求
* 二、整体方案及布局
---
[过渡]
# 二、整体方案及布局
---
[左文右图]
# 整体设备空间布局
* 设备尺寸预估为 2500x1800x2000 mm。
* 采用全封闭式钣金外壳。

[AI配图: 3D concept render of a modern automated manufacturing machine, industrial blue and white color, highly detailed]
---
[正文]
# 节拍分析 (CT)
* 综合 CT 预估为 20s/pcs，满足客户需求。

# Task (任务)
请阅读上述规则。现在，我将提供新项目信息，请按严格格式生成DFM方案 Markdown 文本。不需要任何废话。"""

with st.expander("！！注意！！使用本引擎前必读！！", expanded=False):
    st.markdown("如果您使用大模型（如 DeepSeek、Kimi 等）帮您起草 DFM 方案，**请点击下方按钮一键复制专用指令发送给 AI**。这样 AI 就能生成完美适配本软件规则的文档。")
    
    components.html(
        f"""
        <script>
        function copyToClipboard() {{
            const text = {json.dumps(ai_prompt_text)};
            if (navigator.clipboard && window.isSecureContext) {{
                navigator.clipboard.writeText(text).then(showSuccess).catch(fallbackCopy);
            }} else {{
                fallbackCopy();
            }}
            
            function showSuccess() {{
                const btn = document.getElementById("copy-btn");
                btn.innerHTML = "✅ 复制成功！快去发给大模型吧";
                btn.style.background = "linear-gradient(90deg, #059669, #10B981)";
                setTimeout(() => {{
                    btn.innerHTML = "📋 一键复制专属提示词";
                    btn.style.background = "linear-gradient(90deg, #1E3A8A, #3B82F6)";
                }}, 3000);
            }}
            
            function fallbackCopy() {{
                const textArea = document.createElement("textarea");
                textArea.value = text;
                textArea.style.position = "fixed";
                document.body.appendChild(textArea);
                textArea.focus();
                textArea.select();
                try {{
                    document.execCommand('copy');
                    showSuccess();
                }} catch (err) {{
                    console.error('Fallback copy failed', err);
                }}
                document.body.removeChild(textArea);
            }}
        }}
        </script>
        <button id="copy-btn" onclick="copyToClipboard()" style="
            width: 100%; 
            padding: 12px; 
            margin-bottom: 10px;
            background: linear-gradient(90deg, #1E3A8A, #3B82F6); 
            color: white; 
            border: none; 
            border-radius: 8px; 
            font-size: 16px; 
            font-weight: bold; 
            cursor: pointer; 
            box-shadow: 0 4px 6px rgba(59, 130, 246, 0.2);
            font-family: 'PingFang SC', 'Microsoft YaHei', sans-serif;
            transition: all 0.3s ease;
        ">
            📋 一键复制专属提示词
        </button>
        """,
        height=70
    )
    
    st.code(ai_prompt_text, language="markdown")

st.markdown("#### 01. 导入您的方案内容")
with st.container(border=True):
    st.markdown("#####  📂  方式一：直接上传文档文件")
    uploaded_doc = st.file_uploader("支持 .docx / .md / .txt 格式", type=["docx", "md", "txt"])
    
    st.markdown("#####  ✍️  方式二：输入纯文本")
    raw_md_text = st.text_area("如果没有文档，也可在此处直接粘贴/输入方案文本...", height=400)
    
    md_text_to_process = ""
    if uploaded_doc:
        extracted_text = extract_text_content(uploaded_doc)
        if extracted_text: md_text_to_process = extracted_text
    elif raw_md_text: md_text_to_process = raw_md_text

st.markdown("<br>", unsafe_allow_html=True)

st.markdown("#### 02. 上传 PPT 模板")
with st.container(border=True):
    uploaded_template = st.file_uploader("占位标签", type=["pptx"], label_visibility="collapsed")

st.markdown("<br>", unsafe_allow_html=True)

st.markdown("#### 03. 一键输出DFM")
generate_btn = st.button("✨ 启动智能生成引擎 ✨", use_container_width=True)

if generate_btn:
    if not md_text_to_process: st.error("⚠️ 请在步骤 01 中上传文档或粘贴文本内容！")
    elif not uploaded_template: st.error("⚠️ 请在步骤 02 中上传 PPT 模板！")
    else:
        with st.spinner('🚀 引擎全速运转中，请稍候...'):
            slides = parse_markdown(md_text_to_process)
            template_bytes = uploaded_template.read()
            final_ppt_ppt = create_ppt(slides, template_bytes, api_key_input, ai_provider, base_url, model_name)
            
            if final_ppt_ppt:
                st.success("🎉 生成完毕！您的专业 DFM 报告已就绪。")
                st.download_button(
                    label="📥 点击下载图文 PPT 方案",
                    data=final_ppt_ppt,
                    file_name="JINYA_DFM自动生成方案.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True
                )