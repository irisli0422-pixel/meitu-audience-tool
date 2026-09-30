#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
app.py
======
美图内部 · 广告人群包生成工具（Streamlit 网页）

整体流程：
    密码登录 → 填写品牌需求表单 → 点"生成人群包"
    → 调用火山方舟豆包（开启联网搜索）+ old_skill 业务规则作为 system prompt
    → 把 AI 返回的结构化人群包写成标准格式 .xlsx
    → 网页提供下载，并写入 history.json 历史记录

启动方式：
    streamlit run app.py
"""

import os
from datetime import datetime
from urllib.parse import urlparse

import streamlit as st

# 自己写的几个模块
import ark_client
import demo_sample
import excel_builder
import history_store
import skill_rules

# ============================================================
# 0. 页面基础设置
# ============================================================
st.set_page_config(
    page_title="美图内部 · 广告人群包生成工具",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ------------------------------------------------------------
# 美图品牌视觉样式（主色 #F11D48 / 浅粉 #FFE4EC / 白底 / 深灰字）
# 圆角卡片 + 简洁现代风
# ------------------------------------------------------------
MEITU_CSS = """
<style>
/* ---------- 全局：白底深灰字 ---------- */
.stApp { background-color: #FFFFFF; }
html, body, [class*="css"] { color: #333333; }

/* ---------- 顶部标题条（美图红渐变圆角卡片） ---------- */
.meitu-header {
    background: linear-gradient(120deg, #F11D48 0%, #FF5C7A 100%);
    border-radius: 18px;
    padding: 26px 32px;
    margin-bottom: 8px;
    box-shadow: 0 6px 20px rgba(241, 29, 72, 0.18);
}
.meitu-header h1 {
    color: #FFFFFF;
    font-size: 27px;
    font-weight: 700;
    margin: 0;
    letter-spacing: 0.5px;
}
.meitu-header p {
    color: rgba(255, 255, 255, 0.92);
    font-size: 14px;
    margin: 8px 0 0 0;
}

/* ---------- Streamlit 自带带边框容器：改成美图风圆角浅粉卡片 ---------- */
div[data-testid="stVerticalBlockBorderWrapper"]:has(> div > div[data-testid="stVerticalBlock"]) {
    border-radius: 16px;
}
div[data-testid="stForm"] {
    border: 1px solid #FFE4EC !important;
    border-radius: 16px !important;
    padding: 22px 24px !important;
    background: #FFFFFF;
    box-shadow: 0 2px 10px rgba(241, 29, 72, 0.05);
}

/* ---------- 通用圆角卡片 ---------- */
.meitu-card {
    background: #FFFFFF;
    border: 1px solid #FFE4EC;
    border-radius: 16px;
    padding: 20px 24px;
    margin-bottom: 16px;
    box-shadow: 0 2px 10px rgba(241, 29, 72, 0.06);
}

/* ---------- 浅粉信息卡片 ---------- */
.meitu-pink-card {
    background: #FFE4EC;
    border-radius: 14px;
    padding: 16px 20px;
    margin-bottom: 16px;
    border: none;
}

/* ---------- 板块小标题 ---------- */
.meitu-section-title {
    font-size: 17px;
    font-weight: 700;
    color: #F11D48;
    margin: 6px 0 14px 0;
    padding-left: 11px;
    border-left: 4px solid #F11D48;
    line-height: 1.3;
}

/* ---------- 历史记录单条 ---------- */
.history-item {
    background: #FFFFFF;
    border: 1px solid #FFE4EC;
    border-radius: 12px;
    padding: 13px 18px;
    margin-bottom: 9px;
}
.history-time { color: #999999; font-size: 12px; }
.history-brand { color: #F11D48; font-weight: 700; font-size: 15px; }
.history-product { color: #333333; font-size: 14px; }
.history-summary { color: #777777; font-size: 12.5px; margin-top: 5px; line-height: 1.5; }

/* ---------- 标签（药丸样式） ---------- */
.meitu-pill {
    display: inline-block;
    background: #FFE4EC;
    color: #F11D48;
    border-radius: 999px;
    padding: 3px 12px;
    font-size: 12px;
    font-weight: 600;
    margin-right: 6px;
}

/* ---------- 按钮：主按钮用美图红 ---------- */
div.stButton > button, div.stFormSubmitButton > button {
    border-radius: 10px;
    font-weight: 600;
    border: 1px solid #FFD0DC;
    background: #FFFFFF;
    color: #F11D48;
    transition: all 0.15s ease;
}
div.stButton > button:hover, div.stFormSubmitButton > button:hover {
    border-color: #F11D48;
    background: #FFF5F7;
    color: #F11D48;
}
div.stButton > button[kind="primary"],
div.stFormSubmitButton > button[kind="primary"] {
    background: #F11D48;
    color: #FFFFFF;
    border: none;
    box-shadow: 0 4px 12px rgba(241, 29, 72, 0.25);
}
div.stButton > button[kind="primary"]:hover,
div.stFormSubmitButton > button[kind="primary"]:hover {
    background: #D4123B;
    color: #FFFFFF;
}

/* ---------- 下载按钮：美图红描边 ---------- */
div.stDownloadButton > button {
    border-radius: 10px;
    background: #FFFFFF;
    color: #F11D48;
    border: 1.5px solid #F11D48;
    font-weight: 600;
}
div.stDownloadButton > button:hover {
    background: #F11D48;
    color: #FFFFFF;
}

/* ---------- 输入框圆角 ---------- */
.stTextInput input, .stTextArea textarea, .stSelectbox div[data-baseweb="select"] > div {
    border-radius: 10px !important;
}
.stTextInput input:focus, .stTextArea textarea:focus {
    border-color: #F11D48 !important;
    box-shadow: 0 0 0 1px #F11D48 !important;
}

/* ---------- 隐藏 Streamlit 自带页脚菜单，界面更干净 ---------- */
#MainMenu, footer { visibility: hidden; }

/* ---------- 进度条/滑块等用主色 ---------- */
.stProgress > div > div > div > div { background-color: #F11D48; }

/* ---------- 分隔线 ---------- */
hr { border-color: #FFE4EC; }
</style>
"""
st.markdown(MEITU_CSS, unsafe_allow_html=True)


# ============================================================
# 1. 读取配置（全部来自 .streamlit/secrets.toml，不硬编码）
# ============================================================

def load_config():
    """
    从 st.secrets 读取配置。读不到就返回空值并在页面上提示，
    避免把密钥写死在代码里。
    """
    try:
        ark = st.secrets.get("ark", {})
        return {
            "app_password": st.secrets.get("app_password", ""),
            "monthly_limit": int(st.secrets.get("monthly_call_limit", 500)),
            # 顺手去掉从控制台复制时容易带进来的空格、换行和引号
            "api_key": str(ark.get("api_key", "")).strip().strip('"').strip("'"),
            # 模型标识：可以是模型 ID（如 doubao-seed-2-1-pro-260628）或接入点 ID（ep- 开头）。
            # 顺手去掉首尾空格和末尾多余的点，避免从控制台复制时带进不可见字符。
            "model": str(ark.get("model", "")).strip().rstrip("."),
            "base_url": ark.get("base_url", "https://ark.cn-beijing.volces.com/api/v3/chat/completions"),
            "enable_web_search": bool(ark.get("enable_web_search", True)),
            "enable_deep_thinking": bool(ark.get("enable_deep_thinking", False)),
        }
    except Exception as e:
        st.error("读取 .streamlit/secrets.toml 配置失败：{}".format(e))
        return {
            "app_password": "", "monthly_limit": 500, "api_key": "",
            "model": "", "base_url": "", "enable_web_search": True,
            "enable_deep_thinking": False,
        }


CONFIG = load_config()


# 共享密码是固定的，不接受匿名、不接受别的密码。
SHARED_PASSWORD = "Meitu2026"
# 只有这个邮箱能看到「使用统计」。比较时会转成小写。
ADMIN_EMAIL = "lxy11@meitu.com"


def is_admin():
    """当前登录的是不是管理员。"""
    return current_email() == history_store.normalize_email(ADMIN_EMAIL)


def current_email():
    """当前登录邮箱。没登录时返回空字符串。"""
    return history_store.normalize_email(st.session_state.get("user_email", ""))


def current_user_name():
    """右上角显示的用户名：邮箱 @ 前面那一段。"""
    return history_store.email_prefix(current_email())


# ============================================================
# 2. 邮箱 + 共享密码登录
# ============================================================

def render_login():
    """
    未登录时显示的登录页。返回 True 表示已登录。

    规则：
        - 必须填公司邮箱，不能留空
        - 邮箱必须以 @meitu.com 结尾
        - 密码必须是共享密码 Meitu2026
        - 不验证邮箱是不是真的属于这个人，只挡住公司外部的人
    """
    email_ok = history_store.is_meitu_email(st.session_state.get("user_email", ""))
    if st.session_state.get("logged_in") and email_ok:
        return True

    # 旧会话里如果只有「已登录」没有合法邮箱，清掉，强制重新登录
    st.session_state["logged_in"] = False

    st.markdown(
        """
        <div class="meitu-header">
            <h1>美图内部 · 广告人群包生成工具</h1>
            <p>品牌广告定向人群包圈选规则自动生成 · 仅限公司内部使用</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("")

    left, center, right = st.columns([1, 1.5, 1])
    with center:
        with st.container(border=True):
            st.markdown(
                '<div class="meitu-section-title">公司邮箱登录</div>',
                unsafe_allow_html=True,
            )

            with st.form("login_form"):
                email = st.text_input(
                    "公司邮箱",
                    placeholder="name@meitu.com",
                )
                password = st.text_input(
                    "共享密码",
                    type="password",
                    placeholder="请输入共享密码",
                )
                submitted = st.form_submit_button("登 录", type="primary", use_container_width=True)

            if submitted:
                cleaned = history_store.normalize_email(email)
                if not cleaned:
                    st.error("请输入公司邮箱，不能留空")
                elif not history_store.is_meitu_email(cleaned):
                    st.error("仅限美图公司邮箱登录")
                elif password != SHARED_PASSWORD:
                    st.error("密码不正确，请重新输入")
                else:
                    st.session_state["logged_in"] = True
                    st.session_state["user_email"] = cleaned
                    st.rerun()

            st.caption("请使用 @meitu.com 邮箱和团队共享密码登录。不需要验证邮箱收件箱。")

    return False


# ============================================================
# 3. 页面顶部（已登录）
# ============================================================

def render_header():
    """顶部标题条。右上角显示当前用户名（邮箱前缀）。"""
    user_name = _escape(current_user_name() or "未登录")
    st.markdown(
        """
        <div class="meitu-header" style="display:flex;justify-content:space-between;align-items:center;gap:24px;">
            <div>
                <h1>美图内部 · 广告人群包生成工具</h1>
                <p>填写品牌广告需求 → AI 联网搜索行业洞察 → 按美图六大圈选维度生成定向人群包 Excel</p>
            </div>
            <div style="text-align:right;min-width:120px;">
                <div style="font-size:12px;opacity:0.85;">当前用户</div>
                <div style="font-size:20px;font-weight:700;letter-spacing:0.3px;">{name}</div>
            </div>
        </div>
        """.format(name=user_name),
        unsafe_allow_html=True,
    )

    # 读取业务规则加载情况和调用配额，展示在顶部状态条
    _, loaded_files = skill_rules.build_system_prompt()
    allowed, used, limit = history_store.check_quota(CONFIG["monthly_limit"])

    col1, col2, col3, col4 = st.columns([1.1, 1.1, 1.1, 1])
    with col1:
        st.metric("本月已调用", "{} / {} 次".format(used, limit))
    with col2:
        st.metric("业务规则文件", "{} 个已加载".format(len(loaded_files)))
    with col3:
        st.metric("联网搜索", "已开启" if CONFIG["enable_web_search"] else "已关闭")
    with col4:
        st.write("")
        if st.button("退出登录", use_container_width=True):
            st.session_state["logged_in"] = False
            st.session_state["user_email"] = ""
            st.rerun()

    # 配额用完时高亮提示
    if not allowed:
        st.error("本月调用次数已达上限（{} 次），下个月 1 日自动恢复。如需提额请联系管理员。".format(limit))

    # 业务规则文件缺失时提示（说明 old_skill 被挪走或改名了）
    if len(loaded_files) < 3:
        st.warning(
            "只加载到 {} 个业务规则文件，请检查 old_skill 文件夹是否完整"
            "（需要 SKILL.md、references/meitu_audience_format.md、references/case_studies.md）。".format(
                len(loaded_files)
            )
        )

    render_api_check()
    return allowed


def render_api_check():
    """
    接口连通性自检区。
    点一下就能确认 API Key / 模型端点能不能用，不用先填一整张表单去试。
    """
    # 先在本地体检一遍密钥格式，有问题就直接展开提示，不用等用户去点测试
    key_ok, key_hint = ark_client.validate_api_key_format(CONFIG["api_key"])

    with st.expander(
        "🔧 接口配置与连通性自检（第一次使用建议先点一下）",
        expanded=not key_ok,
    ):
        if not key_ok:
            st.error("**API Key 配置有问题，现在还没法真正生成。**\n\n" + key_hint)

        c1, c2 = st.columns([1, 2.4])
        with c1:
            if st.button("测试接口连通性", use_container_width=True):
                with st.spinner("正在测试火山方舟接口…"):
                    ok, message = ark_client.test_connection(
                        CONFIG["api_key"], CONFIG["model"], CONFIG["base_url"]
                    )
                st.session_state["api_check"] = {"ok": ok, "message": message}
        with c2:
            # 展示当前读到的配置（密钥只显示前后几位，避免截图泄露）
            key = CONFIG["api_key"] or ""
            masked = (key[:8] + "…" + key[-4:]) if len(key) > 14 else (key or "（未配置）")
            st.markdown(
                "- 接口地址：`{url}`\n"
                "- 模型端点：`{model}`\n"
                "- API Key：`{key}`（读取自 `.streamlit/secrets.toml`）\n"
                "- 联网搜索：{web}".format(
                    url=CONFIG["base_url"], model=CONFIG["model"] or "（未配置）",
                    key=masked, web="已开启" if CONFIG["enable_web_search"] else "已关闭",
                )
            )

        check = st.session_state.get("api_check")
        if check:
            if check["ok"]:
                st.success(check["message"])
            else:
                st.error(check["message"])


# ============================================================
# 4. 需求输入表单
# ============================================================

# 广告目标选项（结合 old_skill 的使用场景）
AD_GOALS = [
    "新品上市",
    "品牌曝光",
    "效果转化",
    "联名IP种草",
    "节点大促",
    "品牌焕新/形象升级",
    "门店引流（O2O）",
]


def render_form():
    """
    渲染需求表单。
    返回：(是否点了生成按钮, 表单内容字典)
    """
    st.markdown('<div class="meitu-section-title">① 填写品牌广告需求</div>', unsafe_allow_html=True)

    with st.form("audience_form"):
        # ---- 第一行：品牌 / 产品 / 目标 ----
        c1, c2, c3 = st.columns(3)
        with c1:
            brand_name = st.text_input(
                "品牌名称 *", placeholder="例：安佳 Anchor / SKECHERS 斯凯奇",
                help="填中文名或中英文都写上，AI 会据此联网搜索品牌信息与竞品",
            )
        with c2:
            product_name = st.text_input(
                "广告产品 / 活动名称 *", placeholder="例：安佳黄油 · 核桃拿破仑蛋糕联名",
                help="具体到单品或系列，越具体生成的圈选规则越准",
            )
        with c3:
            ad_goal = st.selectbox(
                "广告目标 *", AD_GOALS,
                help="决定人群包的优先级排序（高转化 / 高复购 / 高传播）",
            )

        # ---- 第二行：目标人群描述 ----
        audience_desc = st.text_area(
            "目标人群描述 *",
            placeholder="用大白话写就行，例：25-35岁一二线城市女性，关注高端护肤，有一定消费力，喜欢在社交平台分享种草内容",
            height=92,
            help="年龄 / 性别 / 城市 / 兴趣 / 消费力，能写多少写多少；没写的 AI 会根据品牌特性自行判断",
        )

        # ---- 第三行：合作内容 / 竞品 ----
        c4, c5 = st.columns(2)
        with c4:
            cooperation = st.text_area(
                "合作内容 / 联名IP / 核心卖点",
                placeholder="例：与 DC 超级英雄联名，主推巧克力+惊喜玩具双重体验",
                height=82,
            )
        with c5:
            competitors = st.text_area(
                "竞品品牌（留空则由 AI 补充）",
                placeholder="例：总统President、银宝Lurpak、德运Devondale（留空时 AI 会自动补国际+国产竞品）",
                height=82,
            )

        # ---- 第四行：预算 / 地区 ----
        c6, c7 = st.columns(2)
        with c6:
            budget = st.text_input(
                "投放预算", placeholder="例：300万 / 50-100万 / 暂未确定",
                help="影响人群包量级规模的参考，可不填",
            )
        with c7:
            regions = st.text_input(
                "投放地区", placeholder="例：全国一二线城市 / 华东+华南 / 北上广深",
                help="会影响 LBS 场景识别维度的 POI 选择，可不填",
            )

        # ---- 第五行：其他补充 ----
        extra = st.text_area(
            "其他补充说明",
            placeholder="例：希望重点圈选母婴家庭人群；避开学生群体；需要包含线下门店到店人群等",
            height=72,
        )

        st.markdown("")

        # 演示模式：不调用接口，用模板示例数据把流程跑通（方便演示/排查）
        demo_mode = st.checkbox(
            "演示模式（不调用 AI 接口，用模板示例数据跑通流程，不消耗调用次数）",
            value=False,
            help="还没配好真实 API Key、或者只想给同事演示一下工具长什么样时勾选。"
                 "产出的 Excel 会明确标注「演示样例」，不要直接对外交付。",
        )

        submitted = st.form_submit_button(
            "🎯  生成人群包", type="primary", use_container_width=True,
        )

    form_data = {
        "brand_name": brand_name,
        "product_name": product_name,
        "ad_goal": ad_goal,
        "audience_desc": audience_desc,
        "cooperation": cooperation,
        "competitors": competitors,
        "budget": budget,
        "regions": regions,
        "extra": extra,
    }
    return submitted, form_data, demo_mode


# ============================================================
# 5. 生成流程
# ============================================================

def do_generate(form_data, demo_mode=False):
    """
    执行完整生成流程：调用 AI → 解析 → 生成 Excel → 写历史记录。
    结果存进 st.session_state["last_result"]，供页面展示。

    demo_mode=True 时跳过接口调用，直接用 demo_sample 的模板数据。
    """
    # ---- 5.1 必填校验 ----
    missing = []
    if not form_data["brand_name"].strip():
        missing.append("品牌名称")
    if not form_data["product_name"].strip():
        missing.append("广告产品 / 活动名称")
    if not form_data["audience_desc"].strip():
        missing.append("目标人群描述")
    if missing:
        st.error("请先填写：{}".format("、".join(missing)))
        return

    if not demo_mode:
        # ---- 5.2 配置校验 ----
        if not CONFIG["api_key"] or not CONFIG["model"]:
            st.error("火山方舟 API 配置不完整，请检查 .streamlit/secrets.toml 里的 [ark] 配置。")
            return

        # 密钥格式本地体检：明显不对就别发请求了，
        # 省掉一次注定失败的调用，也不白占当月调用次数
        key_ok, key_hint = ark_client.validate_api_key_format(CONFIG["api_key"])
        if not key_ok:
            st.error("**API Key 配置有问题，本次没有发起调用**（不计入调用次数）。\n\n" + key_hint)
            st.info(
                "想先把流程跑通看看效果，可以勾选表单下方的「演示模式」再点一次生成。"
            )
            return

        # ---- 5.3 配额校验 ----
        allowed, used, limit = history_store.check_quota(CONFIG["monthly_limit"])
        if not allowed:
            st.error("本月调用次数已用完（{}/{}），无法生成。".format(used, limit))
            return

    # ---- 5.4 组装提示词（old_skill 业务规则作为 system prompt）----
    system_prompt, loaded_files = skill_rules.build_system_prompt()
    user_prompt = skill_rules.build_user_prompt(form_data)

    label = "正在生成演示样例…" if demo_mode else "正在生成人群包，请稍候…"
    status_box = st.status(label, expanded=True)
    with status_box:
        st.write("已加载业务规则：{}".format("、".join(loaded_files) if loaded_files else "无"))

        if demo_mode:
            # ============ 演示模式：不调接口 ============
            st.write("演示模式：跳过 AI 接口调用，使用模板示例数据。")
            data = demo_sample.build_demo_data(form_data)
            result = {
                "ok": True, "data": data, "mode": "演示模式（未调用接口）",
                "web_search_used": False, "content": None, "references": [],
            }
        else:
            # ============ 正常模式：调用豆包 ============
            st.write("正在调用豆包大模型并联网搜索行业人群洞察、竞品投放策略、消费趋势…")
            st.caption("联网搜索 + 生成 10 个人群包约需 1-3 分钟，请不要关闭页面。")

            # 进度提示回调：把当前尝试的调用方式显示出来
            def on_progress(message):
                st.write(message)

            # ---- 5.5 调用豆包 ----
            result = ark_client.call_doubao(
                system_prompt=system_prompt,
                user_prompt=user_prompt,
                api_key=CONFIG["api_key"],
                model=CONFIG["model"],
                base_url=CONFIG["base_url"],
                enable_web_search=CONFIG["enable_web_search"],
                deep_thinking=CONFIG["enable_deep_thinking"],
                progress_callback=on_progress,
            )

            if not result["ok"]:
                # 接口已经被请求过了，记一条失败日志，占用本月额度
                _record_api_call(form_data, "失败")
                status_box.update(label="生成失败", state="error", expanded=True)
                st.session_state["last_result"] = {
                    "ok": False,
                    "error": result["error"],
                    "auth_failed": result.get("auth_failed", False),
                }
                return

            st.write("模型返回成功，实际调用方式：{}".format(result["mode"]))

            if not result["data"]:
                _record_api_call(form_data, "失败")
                status_box.update(label="AI 返回内容格式异常", state="error", expanded=True)
                st.session_state["last_result"] = {
                    "ok": False,
                    "error": "AI 返回的内容没法解析成结构化人群包，请点一次「生成人群包」重试。",
                    "raw_content": result.get("content"),
                }
                return

            data = result["data"]

            # 如果模型自己没给参考来源，用接口返回的联网引用补上
            if not data.get("sources") and result.get("references"):
                data["sources"] = result["references"]

        # ---- 5.6 生成 Excel ----
        st.write("正在按美图标准格式生成 Excel 文件…")
        output_dir = history_store.ensure_output_dir()
        try:
            excel_path = excel_builder.build_audience_excel(
                data=data,
                output_dir=output_dir,
                form=form_data,
                extra_meta={"web_search_used": result["web_search_used"]},
            )
        except Exception as e:
            if not demo_mode:
                _record_api_call(form_data, "失败")
            status_box.update(label="Excel 生成失败", state="error", expanded=True)
            st.session_state["last_result"] = {
                "ok": False,
                "error": "Excel 文件生成出错：{}".format(e),
                "data": data,
            }
            return

        # ---- 5.7 统计并写历史记录 ----
        segments = [s for s in (data.get("segments") or []) if isinstance(s, dict) and s.get("tags")]
        segment_count = len(segments)
        tag_count = sum(len(s.get("tags") or []) for s in segments)

        history_store.add_history(
            brand_name=form_data["brand_name"],
            product_name=form_data["product_name"],
            ad_goal=form_data["ad_goal"],
            requirement_summary=form_data["audience_desc"],
            excel_path=excel_path,
            segment_count=segment_count,
            tag_count=tag_count,
            web_search_used=result["web_search_used"],
            full_form=form_data,
            email=current_email(),
        )
        # 演示模式没有调用接口，不写 usage_log，也不占 500 次额度
        if not demo_mode:
            _record_api_call(form_data, "成功")

        st.write("完成：{} 个人群包 / {} 个子标签".format(segment_count, tag_count))
        status_box.update(label="人群包生成完成", state="complete", expanded=False)

    st.session_state["last_result"] = {
        "ok": True,
        "data": data,
        "excel_path": excel_path,
        "segment_count": segment_count,
        "tag_count": tag_count,
        "web_search_used": result["web_search_used"],
        "mode": result["mode"],
        "brand_name": form_data["brand_name"],
        "demo_mode": demo_mode,
    }


# ============================================================
# 6. 生成结果展示
# ============================================================

def render_result():
    """展示最近一次生成结果：下载按钮 + 人群包预览 + 行业洞察。"""
    result = st.session_state.get("last_result")
    if not result:
        return

    st.markdown('<div class="meitu-section-title">② 生成结果</div>', unsafe_allow_html=True)

    # ---- 失败情况 ----
    if not result.get("ok"):
        st.error(result.get("error", "生成失败"))
        if result.get("raw_content"):
            with st.expander("查看 AI 原始返回内容（排查用）"):
                st.code(result["raw_content"][:6000])
        return

    data = result["data"]
    excel_path = result["excel_path"]

    # ---- 成功概览 ----
    if result.get("demo_mode"):
        st.warning(
            "⚠️ 这是**演示模式**产出的模板示例（{segs} 个人群包 / {tags} 个子标签），"
            "不是 AI 真实分析结果，请勿直接对外交付。"
            "配置真实 API Key 后取消勾选「演示模式」即可生成正式版本。".format(
                segs=result["segment_count"], tags=result["tag_count"],
            )
        )
    else:
        st.success(
            "已生成「{brand}」人群包：{segs} 个人群包 / {tags} 个子标签".format(
                brand=result["brand_name"],
                segs=result["segment_count"],
                tags=result["tag_count"],
            )
        )

    # ---- 下载按钮 ----
    if os.path.exists(excel_path):
        with open(excel_path, "rb") as f:
            st.download_button(
                label="⬇️  下载人群包 Excel（{}）".format(os.path.basename(excel_path)),
                data=f.read(),
                file_name=os.path.basename(excel_path),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
                key="download_latest",
            )
        st.caption("文件同时已保存在本机：{}".format(excel_path))

    # ---- 人群包预览（按人群包分组展开）----
    st.markdown("")
    st.markdown("**人群包预览**")
    for seg in (data.get("segments") or []):
        if not isinstance(seg, dict) or not seg.get("tags"):
            continue
        seg_name = seg.get("name", "未命名人群包")
        tags = seg.get("tags") or []
        with st.expander("{}（{} 个子标签）".format(seg_name, len(tags))):
            rows = []
            for t in tags:
                if isinstance(t, dict):
                    rows.append({
                        "人群标签": t.get("tag", ""),
                        "秀秀单标签逻辑之一": t.get("logic", ""),
                        "人群包量级": t.get("scale", ""),
                    })
            if rows:
                st.dataframe(rows, use_container_width=True, hide_index=True)

    # ---- 行业洞察 ----
    insights = data.get("industry_insights") or []
    if insights:
        with st.expander("联网搜索到的行业洞察 / 竞品投放策略 / 消费趋势"):
            for item in insights:
                st.markdown("- {}".format(item))

    # ---- 竞品名单 ----
    competitors = data.get("competitors") or []
    if competitors:
        with st.expander("竞品品牌名单"):
            st.markdown("、".join(str(c) for c in competitors))

    # ---- 参考来源 ----
    sources = data.get("sources") or []
    if sources:
        with st.expander("参考资料来源"):
            for s in sources:
                if isinstance(s, dict):
                    title = s.get("title") or "参考资料"
                    url = s.get("url") or ""
                    st.markdown("- [{}]({})".format(title, url) if url else "- {}".format(title))
                else:
                    st.markdown("- {}".format(s))


# ============================================================
# 7. 历史记录
# ============================================================

def _record_api_call(form_data, status):
    """把一次真实的火山方舟调用写进 usage_log.json。"""
    history_store.add_usage_log(
        email=current_email(),
        brand=(form_data or {}).get("brand_name"),
        product=(form_data or {}).get("product_name"),
        status=status,
    )


def render_history():
    """历史记录页：只显示当前登录邮箱自己的记录，最近 20 条。"""
    st.markdown('<div class="meitu-section-title">我的历史记录（最近 20 条）</div>', unsafe_allow_html=True)

    email = current_email()
    records = history_store.load_history(email=email)
    if not records:
        st.markdown(
            '<div class="meitu-pink-card">你还没有生成记录。请到「生成人群包」里填写需求并点击生成，'
            '这里只显示当前账号 {} 的记录。</div>'.format(
                _escape(email)
            ),
            unsafe_allow_html=True,
        )
        return

    st.caption("只显示 {} 的记录，共 {} 条，下面展示最近 20 条。".format(email, len(records)))

    for idx, rec in enumerate(records[:20]):
        col_info, col_btn = st.columns([4.2, 1])

        with col_info:
            web_pill = (
                '<span class="meitu-pill">联网</span>' if rec.get("web_search_used") else ""
            )
            count_pill = ""
            if rec.get("segment_count"):
                count_pill = '<span class="meitu-pill">{}包 / {}标签</span>'.format(
                    rec.get("segment_count"), rec.get("tag_count", 0)
                )
            goal_pill = (
                '<span class="meitu-pill">{}</span>'.format(rec.get("ad_goal"))
                if rec.get("ad_goal") else ""
            )

            st.markdown(
                """
                <div class="history-item">
                    <div class="history-time">{time}</div>
                    <div><span class="history-brand">{brand}</span>
                         &nbsp;<span class="history-product">{product}</span></div>
                    <div style="margin-top:6px;">{goal}{count}{web}</div>
                    <div class="history-summary">需求：{summary}</div>
                </div>
                """.format(
                    time=rec.get("timestamp", ""),
                    brand=_escape(rec.get("brand_name", "—")),
                    product=_escape(rec.get("product_name", "")),
                    summary=_escape(rec.get("summary", "—")),
                    goal=goal_pill, count=count_pill, web=web_pill,
                ),
                unsafe_allow_html=True,
            )

        with col_btn:
            st.write("")
            path = rec.get("excel_path", "")
            if path and os.path.exists(path):
                with open(path, "rb") as f:
                    st.download_button(
                        label="下载人群包Excel",
                        data=f.read(),
                        file_name=os.path.basename(path),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="hist_dl_{}".format(idx),
                        use_container_width=True,
                    )
            else:
                # 文件被手动删掉或挪走了
                st.button(
                    "文件已丢失", disabled=True, key="hist_missing_{}".format(idx),
                    use_container_width=True,
                )


def _escape(text):
    """历史记录是用 HTML 渲染的，用户输入要转义，避免把页面样式搞乱。"""
    return (
        str(text or "")
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


# ============================================================
# 8. 主流程
# ============================================================

def render_stats():
    """使用统计页。调用前必须先确认是管理员。"""
    if not is_admin():
        _go_home()
        return

    stats = history_store.build_usage_stats()
    limit = CONFIG["monthly_limit"]
    used = stats["month_total"]

    st.markdown('<div class="meitu-section-title">使用统计</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        st.metric("本月总调用次数", "{} / {} 次".format(used, limit))
    with c2:
        st.metric("统计月份", stats["month"])

    if used >= limit:
        st.error("本月额度已用完。")
    else:
        st.caption("额度按本月全部调用计算，成功和失败都算 1 次。演示模式不计入。")

    rows = stats["rows"]
    if not rows:
        st.markdown(
            '<div class="meitu-pink-card">还没有调用记录。有人调用过火山方舟之后，这里会出现统计。</div>',
            unsafe_allow_html=True,
        )
        return

    st.dataframe(
        rows,
        use_container_width=True,
        hide_index=True,
        column_order=["用户", "累计调用次数", "本月调用次数", "最近一次使用时间"],
    )
    st.caption("按本月调用次数从高到低排序。用户名是邮箱 @ 前面的部分。")


def _page_generate():
    """「生成人群包」页：表单和本次结果。历史记录不放在这里。"""
    quota_ok = st.session_state.get("quota_ok", True)
    submitted, form_data, demo_mode = render_form()
    if submitted:
        if quota_ok or demo_mode:
            do_generate(form_data, demo_mode=demo_mode)
        else:
            st.error("本月调用次数已达上限，无法生成。")
    render_result()


def _page_history():
    render_history()


def _page_stats():
    # 网址被直接打开时，这里再拦一次
    if not is_admin():
        _go_home()
        return
    render_stats()


def _request_path():
    """
    当前浏览器地址的路径，例如 /stats。
    多取几个来源，避免统计页还没登记进导航时，路径被吞掉。
    """
    urls = []
    try:
        if st.context.url:
            urls.append(st.context.url)
    except Exception:
        pass
    try:
        from streamlit.runtime.scriptrunner_utils.script_run_context import (
            get_script_run_ctx,
        )
        ctx = get_script_run_ctx()
        raw = getattr(getattr(ctx, "context_info", None), "url", None)
        if raw:
            urls.append(raw)
    except Exception:
        pass

    for url in urls:
        path = urlparse(url).path.rstrip("/")
        if path.endswith("/stats"):
            return path
    if urls:
        return urlparse(urls[0]).path.rstrip("/")
    return ""


def _go_home():
    """不是管理员时，离开统计页，回到生成人群包。"""
    st.switch_page(HOME_PAGE)


# 默认页的地址是网站根路径 / 。统计页地址是 /stats。
HOME_PAGE = st.Page(_page_generate, title="生成人群包", icon="🎯", default=True)
HISTORY_PAGE = st.Page(_page_history, title="我的历史记录", icon="📋", url_path="history")
STATS_PAGE = st.Page(_page_stats, title="使用统计", icon="📊", url_path="stats")


def main():
    # 未登录：只显示登录页，不出现任何标签
    if not render_login():
        return

    # 有人直接打开 /stats，但当前邮箱不是管理员：踢回首页
    if _request_path().endswith("/stats") and not is_admin():
        _go_home()
        return

    st.session_state["quota_ok"] = render_header()

    pages = [HOME_PAGE, HISTORY_PAGE]
    if is_admin():
        pages.append(STATS_PAGE)

    # 顶部标签。非管理员的列表里没有「使用统计」，所以入口不会出现。
    st.navigation(pages, position="top").run()

    st.markdown("---")
    st.caption(
        "美图内部 · 广告人群包生成工具｜当前用户 {}｜{}".format(
            current_email(), datetime.now().strftime("%Y-%m-%d")
        )
    )


if __name__ == "__main__":
    main()
