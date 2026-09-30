#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
skill_rules.py
==============
把 old_skill 文件夹里的业务规则，组装成调用豆包大模型时用的 system prompt。

设计思路：
    不把业务规则写死在代码里，而是每次运行时**实时读取** old_skill 下的
    SKILL.md、references/meitu_audience_format.md、references/case_studies.md。
    这样以后业务规则有变动，只需要改 old_skill 里的 md 文件，网页无需改代码。

对外只暴露两个东西：
    - build_system_prompt()  ：生成 system prompt（业务规则 + 输出格式约束）
    - build_user_prompt(...) ：把网页表单填的内容组装成 user prompt
"""

import os

# 当前文件所在目录，用于定位 old_skill 文件夹
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OLD_SKILL_DIR = os.path.join(BASE_DIR, "old_skill")

# 需要读取的业务规则文件（相对 old_skill 的路径）
SKILL_FILES = [
    "SKILL.md",
    "references/meitu_audience_format.md",
    "references/case_studies.md",
]


def _read_skill_file(relative_path):
    """读取单个 old_skill 文件，读不到就返回空字符串（不让网页崩掉）。"""
    full_path = os.path.join(OLD_SKILL_DIR, relative_path)
    try:
        with open(full_path, "r", encoding="utf-8") as f:
            return f.read()
    except Exception:
        return ""


def load_skill_knowledge():
    """
    把 old_skill 里所有业务规则文件拼成一整段文本。
    返回：(拼好的知识文本, 成功读到的文件名列表)
    """
    chunks = []
    loaded = []
    for rel in SKILL_FILES:
        content = _read_skill_file(rel)
        if content.strip():
            loaded.append(rel)
            chunks.append(
                "===== 业务规则文件：old_skill/{path} =====\n{body}".format(
                    path=rel, body=content.strip()
                )
            )
    return "\n\n".join(chunks), loaded


# ============================================================
# 输出 JSON 结构约束
# ------------------------------------------------------------
# 这段是给模型看的"填空模板"。字段全部对应 old_skill 里定义的
# Excel 四大板块（数据说明 / 人群包汇总 / 基础人群定向 / 定制人群包圈选规则 / 备注）
# ============================================================
OUTPUT_SCHEMA = """
{
  "brand_name": "品牌中文名",
  "brand_name_en": "品牌英文名（只含英文字母数字，用于Excel工作表名）",
  "colors": {
    "title":   "标题行背景色（品牌主色深色，6位十六进制不带#，如 1A1A1A）",
    "section": "板块标题背景色（品牌主色中色）",
    "header":  "表头背景色（品牌主色亮色）",
    "basic":   "基础定向背景色（品牌色系浅色）"
  },
  "data_notes": [
    "1. 整体逻辑说明……",
    "2. 涉及的APP：美图系场景（美图秀秀、美颜相机等）；",
    "3. 人群包采用近6个月中符合圈选规则的人群；",
    "4. 品牌背景信息（产品线/定位/核心卖点）……",
    "5. 本次核心场景：……"
  ],
  "summary_text": "人群包汇总公式，形如：【基础定向 ∪ 影像识别XX人群 ∪ 场景识别XX人群】∩【人群包A ∪ 人群包B ∪ … ∪ 种子人群包】",
  "basic_targeting": [
    {"dimension": "性别维度",     "logic": "女 / 男（注明偏向及原因），整体男女全覆盖"},
    {"dimension": "年龄维度",     "logic": "核心人群年龄段及说明"},
    {"dimension": "兴趣爱好维度", "logic": "安装应用偏好：……\\n社区兴趣偏好：……"},
    {"dimension": "中高消费维度", "logic": "中等及以上消费力（月收入区间），对……有消费意愿"}
  ],
  "segments": [
    {
      "name": "人群包名称（如 本竞品人群包）",
      "color": "该人群包行背景浅色（6位十六进制不带#）",
      "tags": [
        {
          "tag":   "子标签名称（如 影像识别-竞品品牌）",
          "logic": "圈选规则详细描述，必须严格套用 old_skill 六大维度的句式模板",
          "scale": "预估量级（如 280w）"
        }
      ]
    }
  ],
  "remarks": [
    "1. 以上人群包量级为预估量级，实际投放时以系统最终圈选结果为准；",
    "2. 并集（∪）/ 交集（∩）逻辑说明……",
    "3. 影像识别技术说明……",
    "4. LBS场景识别说明……",
    "5. 原生需求说明……",
    "6. 品牌核心信息……",
    "7. 核心人群优先级：XX人群包(高转化) → XX人群包(高复购) → XX人群包(高传播)。"
  ],
  "competitors": ["竞品品牌1", "竞品品牌2", "……（国际+国产都要覆盖）"],
  "industry_insights": [
    "联网搜索得到的行业人群洞察 / 竞品投放策略 / 消费趋势，每条一句话，5-8条"
  ],
  "sources": [
    {"title": "参考资料标题", "url": "网址"}
  ]
}
"""


def build_system_prompt():
    """
    组装 system prompt：old_skill 全部业务规则 + 硬性输出要求。
    返回：(system_prompt 文本, 成功加载的业务规则文件列表)
    """
    knowledge, loaded = load_skill_knowledge()

    system_prompt = """你是美图（Meitu）广告商业化团队的资深人群包策划专家，专门为品牌方设计「美图定制人群包圈选规则」。

下面是你必须严格遵循的内部业务规则文档（来自内部 skill 库），包含人群包维度定义、圈选规则句式、标准组合、量级范围、Excel格式规范和历史交付案例：

{knowledge}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【本次任务的硬性要求】
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

1. 先联网搜索，再设计。必须搜索并参考：
   - 该品牌 / 该品类的目标人群洞察与消费趋势
   - 竞品品牌名单（国际品牌 + 国产品牌都要覆盖）
   - 竞品近期的营销投放策略、主打卖点、联名动作
   搜索到的关键结论写进 industry_insights，来源写进 sources。

2. 人群包结构（严格按业务规则文档）：
   - 必选：「本竞品人群包」放第1个，「种子人群包」放最后1个
   - 中间设计 6-8 个品牌定制人群包，总计 10 个人群包
   - 每个人群包 4-6 个子标签，全部人群包子标签总数控制在 45-55 个
   - 每个人群包至少包含 1 个「影像识别」维度子标签 + 1-2 个其他维度
     （其他维度 = LBS场景识别 / 原生需求 / 社区互动 / 广告点击行为）
   - 「种子人群包」只用种子人群维度（历史互动 + 品牌搜索）

3. 圈选规则（logic 字段）必须逐字套用业务规则文档里的六大维度句式模板，例如：
   - 影像识别（场景）：影像识别用户近180天内图片中高频出现XXX场景：[具体元素列举]
   - 影像识别（品牌）：近180天内通过影像识别侦察到【品牌A、品牌B、品牌C】品牌形象的用户
   - LBS：近180天内通过LBS及影像识别侦察到【场所A、场所B、场所C】的用户，频次≥N次
   - 原生需求：近180天内在美图秀秀产品产生原生需求【关键词A、关键词B】配方/滤镜/贴纸的用户
   - 社区互动：近180天内在美图社区中发布或互动【话题A、话题B】相关话题及图片的用户
   - 广告点击：在美图系平台的广告点击行为数据判断出用户对相关广告感兴趣，广告点击兴趣=关键词A/关键词B，频次=30天
   - 种子人群：2024-2026年度在美图秀秀APP点击过XX品牌广告、参与过XX话题互动、使用过XX联名配方/贴纸/滤镜的人群（去重种子人群包）
   规则里的关键词、场景、品牌名、POI 名称必须具体到可执行，禁止出现 XXX 之类的占位符。

4. 时间窗口统一用「近180天」；广告点击可用 30/90/180 天；品牌搜索用「近365天」。

5. 量级（scale 字段）：普通子标签 120w-1000w，种子人群包子标签 50w-85w，写成 "380w" 这种形式。

6. 配色：从品牌 VI 提取品牌色，按「标题深色 → 板块中色 → 表头亮色 → 基础定向浅色」递进。
   每个人群包的 color 用不同的浅色，便于阅读区分。

7. 输出格式：只输出一个 JSON 对象，不要输出任何解释文字，不要用 markdown 代码块包裹。
   JSON 结构必须严格如下（注释仅说明用途，实际输出不要带注释）：
{schema}

8. 全部内容使用简体中文（品牌名可保留英文原名）。""".format(
        knowledge=knowledge if knowledge else "（未能读取到 old_skill 业务规则文件，请检查 old_skill 文件夹）",
        schema=OUTPUT_SCHEMA,
    )

    return system_prompt, loaded


def build_user_prompt(form):
    """
    把网页表单内容组装成 user prompt。

    参数 form 是一个字典，包含网页上填的各项：
        brand_name / product_name / ad_goal / audience_desc /
        competitors / budget / regions / cooperation / extra
    """
    def _or_auto(value, auto_hint="（未提供，请你根据品牌特性和联网搜索结果自行判断补充）"):
        """没填的项，明确告诉模型自己补，而不是留空。"""
        text = (value or "").strip()
        return text if text else auto_hint

    user_prompt = """请为下面这个品牌需求，设计一套完整的「美图定制人群包圈选规则」。

【品牌名称】{brand_name}
【主推产品 / 活动名称】{product_name}
【本次广告目标】{ad_goal}
【目标人群描述】{audience_desc}
【合作内容 / 联名IP / 核心卖点】{cooperation}
【竞品品牌】{competitors}
【投放预算】{budget}
【投放地区】{regions}
【其他补充说明】{extra}

请先联网搜索该品牌与品类的最新人群洞察、竞品投放策略和消费趋势，再依据内部业务规则文档设计 10 个人群包（本竞品人群包开头、种子人群包结尾），最后严格按要求的 JSON 结构输出。""".format(
        brand_name=_or_auto(form.get("brand_name")),
        product_name=_or_auto(form.get("product_name")),
        ad_goal=_or_auto(form.get("ad_goal")),
        audience_desc=_or_auto(form.get("audience_desc")),
        cooperation=_or_auto(form.get("cooperation")),
        competitors=_or_auto(form.get("competitors"), "（未提供，请根据品类自行补充国际+国产竞品名单）"),
        budget=_or_auto(form.get("budget"), "（未提供，按中等预算规模考虑人群包量级）"),
        regions=_or_auto(form.get("regions"), "（未提供，按全国投放考虑，一二线城市为主）"),
        extra=_or_auto(form.get("extra"), "（无）"),
    )
    return user_prompt
