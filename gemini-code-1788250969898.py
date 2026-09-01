import os
from datetime import datetime
import pandas as pd
import streamlit as st
from google import genai

# 页面基础配置
st.set_page_config(
    page_title="TikTok Shop 视频拆解与复盘工作台", page_icon="🎬", layout="wide"
)

# 初始化云端/安全配置的 Gemini Client
api_key = None
try:
  if "GEMINI_API_KEY" in st.secrets:
    api_key = st.secrets["GEMINI_API_KEY"]
except Exception:
  pass

if not api_key:
  api_key = os.environ.get("GEMINI_API_KEY")

# 侧边栏：团队权限与跟进人配置
st.sidebar.markdown("### 🔒 团队安全与权限工作台")
input_pwd = st.sidebar.text_input("请输入访问密码", type="password")
operator_name = st.sidebar.text_input("当前跟进人姓名/ID", value="运营小哥")

# 简单密码校验（默认密码 8888）
if input_pwd != "8888":
  st.sidebar.warning("请输入正确的团队访问密码以解锁系统 (默认: 8888)")
  st.stop()

st.sidebar.success(f"已解锁 | 当前跟进人: {operator_name}")
st.sidebar.markdown("---")

if not api_key:
  st.error(
      "未检测到后台 GEMINI_API_KEY 配置，请在 Streamlit Secrets 中配置。"
  )
  st.stop()

client = genai.Client(api_key=api_key)

# 历史数据持久化路径
LOG_FILE = "history_log.csv"


def save_log(task_type, detail_info):
  """自动将生成结果或复盘内容持久化到 CSV 文件中"""
  new_data = {
      "时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
      "跟进人": operator_name,
      "任务类型": task_type,
      "核心详情": detail_info,
  }
  df_new = pd.DataFrame([new_data])
  if os.path.exists(LOG_FILE):
    df_new.to_csv(LOG_FILE, mode="a", header=False, index=False, encoding="utf-8")
  else:
    df_new.to_csv(LOG_FILE, mode="w", header=True, index=False, encoding="utf-8")


# 主界面标签页
tab1, tab2 = st.tabs(
    ["🎬 民宿对标视频拆解与脚本生成", "📊 视频漏斗复盘与跑偏诊断"]
)

# ----------------- 标签页 1：对标视频拆解与脚本生成 -----------------
with tab1:
  st.markdown("### TikTok Shop 民宿实景爆款拆解与脚本生成")
  st.caption("固定模型: gemini-2.5-flash | 零自由对话 | 全流程标准化 SOP")

  col1, col2 = st.columns([1, 1])

  with col1:
    uploaded_file = st.file_uploader(
        "1. 上传对标视频 (.mp4，建议控制在30秒内)", type=["mp4"]
    )
    product_selling_point = st.text_area(
        "2. 我的产品核心卖点",
        placeholder=(
            "例如：一涂遮盖快递标签上的姓名和地址；可重复补充墨水；大号适合桌面，小号适合随身携带。"
        ),
    )
    scene_limit = st.selectbox(
        "3. 民宿拍摄场景限制",
        [
            "民宿客厅·沙发休闲区",
            "民宿厨房·多功能岛台",
            "民宿卧室·床头沉浸式",
            "民宿卫生间/梳妆台",
            "纯桌面特写(无真人)",
        ],
    )

  with col2:
    st.info(
        "**AI 执行标准逻辑**：\n"
        "• 提取前3秒视觉与痛点钩子（Hook）\n"
        "• 结合选定的民宿实景环境重新编排镜头\n"
        "• 输出英文口播及可直接用于拍摄的结构化分镜"
    )

  if st.button("开始分析视频并生成脚本", type="primary", use_container_width=True):
    if not uploaded_file:
      st.error("请先上传一个对标视频文件。")
    elif not product_selling_point.strip():
      st.error("请填写“我的产品核心卖点”。")
    else:
      with st.spinner("AI 正在深度解析视频画面、音频与民宿场景结构..."):
        try:
          bytes_data = uploaded_file.getvalue()
          video_file_obj = client.files.upload(
              file=bytes_data, mime_type="video/mp4"
          )

          prompt = f"""
          你是一个顶级的美国 TikTok Shop 跨境电商短视频编导。
          请分析这个对标视频，并结合我的产品卖点和民宿场景，输出高转化的拍摄脚本。
          我的产品核心卖点：{product_selling_point}
          选定的民宿拍摄场景：{scene_limit}

          请严格输出 Markdown 格式的分镜表格，包含以下列：
          1. 分镜序号
          2. 景别/机位
          3. 画面描述(结合民宿实景与道具动作)
          4. 英文口播文案/字幕
          5. 音效/节奏提示
          6. 设计目的(底层心理逻辑)
          """

          response = client.models.generate_content(
              model="gemini-2.5-flash", contents=[video_file_obj, prompt]
          )

          script_result = response.text
          st.markdown("### 生成的标准化拍摄脚本")
          st.markdown(script_result)

          # 自动保存到后端日志
          save_log(
              "脚本生成",
              f"卖点: {product_selling_point} | 场景: {scene_limit}",
          )

          # 提供 Excel 下载
          # 简单将生成文本转成 DataFrame 提供下载
          df_download = pd.DataFrame(
              [
                  {
                      "跟进人": operator_name,
                      "卖点": product_selling_point,
                      "场景": scene_limit,
                      "完整脚本内容": script_result,
                  }
              ]
          )
          csv_bytes = df_download.to_csv(index=False).encode("utf-8")
          st.download_button(
              label="📥 一键下载脚本记录 (CSV)",
              data=csv_bytes,
              file_name=f"tiktok_script_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
              mime="text/csv",
          )

        except Exception as e:
          st.error(f"处理过程中发生错误: {e}")

# ----------------- 标签页 2：数据复盘与脚本迭代 -----------------
with tab2:
  st.markdown("### 视频漏斗复盘与跑偏诊断")
  st.caption(
      "通过核心漏斗数据（前3秒、完播率、锚点点击率）精准定位视频是否跑偏"
  )

  raw_script_input = st.text_area(
      "1. 粘贴第一步生成的原始脚本或大纲",
      placeholder="在此粘贴分镜脚本或核心文案内容...",
  )

  c1, c2, c3, c4 = st.columns(4)
  with c1:
    rate_3s = st.number_input("前3秒留存率 (%)", 0.0, 100.0, 25.0, 1.0)
  with c2:
    rate_completion = st.number_input("平均完播率 (%)", 0.0, 100.0, 15.0, 1.0)
  with c3:
    rate_click = st.number_input("商品锚点点击率 (%)", 0.0, 100.0, 2.0, 0.1)
  with c4:
    rate_conversion = st.number_input("订单转化率 (%)", 0.0, 100.0, 0.5, 0.1)

  if st.button(
      "诊断视频跑偏原因并生成优化版脚本", type="primary", use_container_width=True
  ):
    with st.spinner("AI 正在诊断漏斗数据瓶颈..."):
      diagnostic_prompt = f"""
      你是一个资深的 TikTok Shop 数据运营专家。某条短视频的数据如下：
      - 前3秒留存率: {rate_3s}% (基准线通常>30%)
      - 平均完播率: {rate_completion}% (基准线通常>20%)
      - 商品锚点点击率: {rate_click}% (基准线通常>3%)
      - 订单转化率: {rate_conversion}%
      
      原始脚本参考：
      {raw_script_input}

      请帮我进行深度诊断：
      1. 判断视频是否跑偏，指出具体是哪个环节崩了（例如：前3秒低说明钩子失败，锚点点击低说明痛点激发或引导不足）。
      2. 给出针对民宿实景拍摄环境的具体整改方案。
      3. 输出一份调整优化后的全新分镜脚本。
      """

      res_diag = client.models.generate_content(
          model="gemini-2.5-flash", contents=[diagnostic_prompt]
      )

      diagnosis_text = res_diag.text
      st.markdown("### 📊 数据诊断与优化报告")
      st.markdown(diagnosis_text)

      save_log(
          "数据复盘",
          f"3秒:{rate_3s}%|完播:{rate_completion}%|点击:{rate_click}%|转化:{rate_conversion}%",
      )

# ----------------- 侧边栏：历史记录持久化展示 -----------------
st.sidebar.markdown("---")
st.sidebar.markdown("### 📁 团队云端留存记录")
if os.path.exists(LOG_FILE):
  df_history = pd.read_csv(LOG_FILE)
  st.sidebar.dataframe(df_history.tail(10), use_container_width=True)

  with open(LOG_FILE, "rb") as f:
    st.sidebar.download_button(
        "📥 下载完整团队历史备份",
        f,
        file_name="team_history_backup.csv",
        mime="text/csv",
    )
else:
  st.sidebar.info("暂无历史记录，开始生成后自动留存。")