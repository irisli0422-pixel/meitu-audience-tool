#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
登录规则、历史按邮箱隔离、使用统计的本地测试。
不调用火山方舟，不改正式的 history.json / usage_log.json。
"""

import os
import tempfile

import history_store


def main():
    # ---------- 邮箱规则 ----------
    assert history_store.is_meitu_email("zhangsan@meitu.com")
    assert history_store.is_meitu_email("  ZhangSan@Meitu.com  ")
    assert not history_store.is_meitu_email("")
    assert not history_store.is_meitu_email("@meitu.com")
    assert not history_store.is_meitu_email("zhangsan@gmail.com")
    assert not history_store.is_meitu_email("zhangsan@meitu.com.cn")
    assert history_store.email_prefix("ZhangSan@Meitu.com") == "zhangsan"
    print("邮箱规则通过")

    # 把读写指到临时目录，避免污染正式数据
    tmp = tempfile.mkdtemp(prefix="audience_stats_")
    history_store.HISTORY_FILE = os.path.join(tmp, "history.json")
    history_store.USAGE_LOG_FILE = os.path.join(tmp, "usage_log.json")

    # ---------- 历史记录按邮箱隔离 ----------
    history_store.add_history(
        "CHANEL", "礼盒", "品牌曝光", "高奢女性",
        os.path.join(tmp, "a.xlsx"), email="zhangsan@meitu.com",
    )
    history_store.add_history(
        "DIOR", "香水", "新品上市", "都市女性",
        os.path.join(tmp, "b.xlsx"), email="lisi@meitu.com",
    )
    zhang = history_store.load_history(email="zhangsan@meitu.com")
    li = history_store.load_history(email="lisi@meitu.com")
    assert len(zhang) == 1 and zhang[0]["brand_name"] == "CHANEL"
    assert len(li) == 1 and li[0]["brand_name"] == "DIOR"
    assert history_store.load_history(email="nobody@meitu.com") == []
    print("历史记录按邮箱隔离通过")

    # ---------- 调用日志与统计 ----------
    history_store.add_usage_log("zhangsan@meitu.com", "CHANEL", "礼盒", "成功")
    history_store.add_usage_log("zhangsan@meitu.com", "CHANEL", "礼盒", "失败")
    history_store.add_usage_log("lisi@meitu.com", "DIOR", "香水", "成功")

    stats = history_store.build_usage_stats()
    assert stats["month_total"] == 3
    assert stats["rows"][0]["用户"] == "zhangsan"
    assert stats["rows"][0]["累计调用次数"] == 2
    assert stats["rows"][0]["本月调用次数"] == 2
    assert stats["rows"][1]["用户"] == "lisi"
    assert stats["rows"][0]["本月调用次数"] >= stats["rows"][1]["本月调用次数"]

    allowed, used, limit = history_store.check_quota(500)
    assert allowed and used == 3 and limit == 500
    print("使用统计通过：本月 {} 次，第一名 {}".format(used, stats["rows"][0]["用户"]))
    print("全部通过")


if __name__ == "__main__":
    main()
