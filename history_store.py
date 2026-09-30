#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
history_store.py
================
三件事：
    1. 历史记录读写 —— 存在本地 history.json（按登录邮箱分开）
    2. 每次真实调用火山方舟的日志 —— 存在本地 usage_log.json
    3. 每月调用次数 —— 从 usage_log.json 里按月份数出来（上限默认 500 次/月）

之所以用本地 json 而不是数据库：公司内部小工具，一个文件最省事，
出问题时用记事本就能打开看/改。
"""

import json
import os
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

HISTORY_FILE = os.path.join(BASE_DIR, "history.json")       # 历史记录
USAGE_LOG_FILE = os.path.join(BASE_DIR, "usage_log.json")   # 每次 API 调用日志
USAGE_FILE = os.path.join(BASE_DIR, "usage.json")           # 旧的月度计数（已不再作为统计来源）
OUTPUT_DIR = os.path.join(BASE_DIR, "output")               # Excel 存放目录

# 允许登录的邮箱后缀
ALLOWED_EMAIL_DOMAIN = "@meitu.com"


# ============================================================
# 通用 json 读写（带容错，文件坏了不让网页崩）
# ============================================================

def _load_json(path, default):
    """读 json 文件；文件不存在或内容损坏时返回 default。"""
    if not os.path.exists(path):
        return default
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def _save_json(path, data):
    """
    写 json 文件。先写临时文件再改名，避免写一半断电导致文件损坏。
    """
    try:
        tmp_path = path + ".tmp"
        with open(tmp_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.replace(tmp_path, path)
        return True
    except Exception:
        return False


# ============================================================
# 一、历史记录
# ============================================================

def normalize_email(email):
    """邮箱统一成小写、去掉首尾空格，方便比较。"""
    return str(email or "").strip().lower()


def email_prefix(email):
    """zhangsan@meitu.com → zhangsan。没有 @ 就原样返回。"""
    text = normalize_email(email)
    if "@" in text:
        return text.split("@", 1)[0]
    return text


def is_meitu_email(email):
    """
    是否是可用的美图公司邮箱。
    必须是「非空前缀 + @meitu.com」，不允许只有后缀，也不允许多余的 @。
    """
    text = normalize_email(email)
    domain = ALLOWED_EMAIL_DOMAIN
    if not text.endswith(domain):
        return False
    local = text[: -len(domain)]
    if not local or "@" in local or any(ch.isspace() for ch in local):
        return False
    return True


def load_history(email=None):
    """
    读取历史记录，按时间倒序（最新的在最前面）返回。

    传入 email 时只返回这个邮箱自己的记录。
    没有邮箱字段的旧记录不会出现在任何人的列表里。
    """
    records = _load_json(HISTORY_FILE, [])
    if not isinstance(records, list):
        return []
    records = [r for r in records if isinstance(r, dict)]
    if email:
        wanted = normalize_email(email)
        records = [r for r in records if normalize_email(r.get("email")) == wanted]
    records.sort(key=lambda r: str(r.get("timestamp", "")), reverse=True)
    return records


def add_history(brand_name, product_name, ad_goal, requirement_summary,
                excel_path, segment_count=0, tag_count=0,
                web_search_used=False, full_form=None, email=""):
    """
    追加一条历史记录。

    存的字段：
        timestamp           生成时间
        brand_name          品牌名称
        product_name        广告产品/活动
        ad_goal             广告目标
        requirement         用户输入的需求（完整）
        summary             需求摘要（列表里显示用，截断到 80 字）
        excel_path          Excel 文件路径
        excel_filename      Excel 文件名
        segment_count       人群包个数
        tag_count           子标签个数
        web_search_used     本次是否开了联网搜索
        form                完整表单内容（便于以后复用）
        email               生成这条记录的登录邮箱
    """
    requirement = str(requirement_summary or "").strip()
    summary = requirement.replace("\n", " ")
    if len(summary) > 80:
        summary = summary[:80] + "…"

    record = {
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "brand_name": str(brand_name or "").strip(),
        "product_name": str(product_name or "").strip(),
        "ad_goal": str(ad_goal or "").strip(),
        "requirement": requirement,
        "summary": summary,
        "excel_path": str(excel_path or ""),
        "excel_filename": os.path.basename(str(excel_path or "")),
        "segment_count": int(segment_count or 0),
        "tag_count": int(tag_count or 0),
        "web_search_used": bool(web_search_used),
        "email": normalize_email(email),
        "form": full_form or {},
    }

    records = _load_json(HISTORY_FILE, [])
    if not isinstance(records, list):
        records = []
    records.append(record)
    _save_json(HISTORY_FILE, records)
    return record


# ============================================================
# 二、每月调用次数计数
# ============================================================

def _current_month():
    """当前月份标识，如 '2026-09'。"""
    return datetime.now().strftime("%Y-%m")


# ============================================================
# 二、API 调用日志（usage_log.json）与使用统计
# ============================================================

def load_usage_log():
    """读取全部调用日志。文件不存在或损坏时返回空列表。"""
    records = _load_json(USAGE_LOG_FILE, [])
    if not isinstance(records, list):
        return []
    return [r for r in records if isinstance(r, dict)]


def add_usage_log(email, brand, product, status):
    """
    追加一条火山方舟调用日志。

    只在真正发起了 API 请求之后写。演示模式不走这里。
    status 只能是「成功」或「失败」。
    """
    status_text = "成功" if status == "成功" else "失败"
    record = {
        "时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "邮箱": normalize_email(email),
        "品牌": str(brand or "").strip(),
        "产品": str(product or "").strip(),
        "状态": status_text,
    }
    records = load_usage_log()
    records.append(record)
    _save_json(USAGE_LOG_FILE, records)
    return record


def build_usage_stats():
    """
    按邮箱前缀汇总调用情况。

    返回：
        {
          "month": "2026-09",
          "month_total": 12,          # 本月全部调用（成功+失败）
          "rows": [
              {
                "用户": "zhangsan",
                "累计调用次数": 8,
                "本月调用次数": 3,
                "最近一次使用时间": "2026-09-30 15:30:00",
              },
              ...
          ]
        }
    行按「本月调用次数」从高到低排，次数相同再按累计次数排。
    """
    month = _current_month()
    buckets = {}

    for item in load_usage_log():
        email = normalize_email(item.get("邮箱"))
        if not email:
            continue
        prefix = email_prefix(email) or email
        when = str(item.get("时间") or "")
        bucket = buckets.setdefault(prefix, {
            "用户": prefix,
            "累计调用次数": 0,
            "本月调用次数": 0,
            "最近一次使用时间": "",
        })
        bucket["累计调用次数"] += 1
        if when.startswith(month):
            bucket["本月调用次数"] += 1
        if when > bucket["最近一次使用时间"]:
            bucket["最近一次使用时间"] = when

    rows = list(buckets.values())
    rows.sort(key=lambda r: (r["本月调用次数"], r["累计调用次数"], r["最近一次使用时间"]), reverse=True)
    month_total = sum(r["本月调用次数"] for r in rows)
    return {"month": month, "month_total": month_total, "rows": rows}


def get_usage():
    """
    读取当月调用次数。
    如果记录里的月份不是本月，说明跨月了，自动清零。
    返回：{"month": "2026-09", "count": 12}
    """
    usage = _load_json(USAGE_FILE, {})
    if not isinstance(usage, dict):
        usage = {}

    month = _current_month()
    if usage.get("month") != month:
        # 跨月清零（保留上个月的数字到 history 字段里做个留痕）
        usage = {"month": month, "count": 0, "previous": usage or None}
        _save_json(USAGE_FILE, usage)

    return {"month": usage.get("month", month), "count": int(usage.get("count", 0))}


def increment_usage():
    """当月调用次数 +1，返回加完之后的次数。"""
    usage = get_usage()
    usage["count"] = int(usage.get("count", 0)) + 1
    _save_json(USAGE_FILE, {"month": usage["month"], "count": usage["count"]})
    return usage["count"]


def check_quota(limit):
    """
    检查本月配额是否还够。
    次数来自 usage_log.json 里本月的全部调用（成功和失败都算一次）。
    返回：(是否允许调用, 已用次数, 上限)
    """
    used = build_usage_stats()["month_total"]
    return (used < int(limit), used, int(limit))


def ensure_output_dir():
    """确保 output 文件夹存在，返回它的路径。"""
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    return OUTPUT_DIR
