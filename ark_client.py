#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ark_client.py
=============
负责调用火山方舟（豆包）大模型，并开启内置联网搜索。

【关于联网搜索走哪个接口 —— 很重要】
火山方舟有两套对话接口，联网搜索的支持情况不一样：

    接口                              联网搜索
    /api/v3/chat/completions          ❌ 不支持
    /api/v3/responses                 ✅ 支持（tools: [{"type": "web_search"}]）

官方能力对照表明确写了：需要联网搜索 / 图像处理 / 知识库检索 / MCP 时，
必须使用 Responses API。所以本文件的主路径是 Responses API。

（另外还有第三条路：在控制台用「零代码应用」配好联网内容插件，拿到 bot-xxx
  再调 /api/v3/bots/chat/completions。但那需要人工在控制台建应用，
  对内部小工具来说太绕，所以不走这条。）

因此调用策略是：
    1. 优先：Responses API + 内置联网搜索  → 能联网，这是我们想要的
    2. 兜底：Chat Completions 不带工具      → 不能联网，但至少能出结果

注意：绝对不要在 chat/completions 上声明一个叫 web_search 的 function 工具。
那样服务器会接受请求，但模型只会返回一个"我要调用工具"的 tool_call，
正文是空的，看起来像成功其实是坏的。这个坑踩过一次，别再踩。
"""

import json
import re

import requests

# 单次请求超时时间（秒）。联网搜索 + 长文本生成比较慢，给 5 分钟。
REQUEST_TIMEOUT = 300

# 明确不走任何代理。
# 原因：本机系统里配了 http_proxy / https_proxy 指向本地代理 127.0.0.1:54442，
# 但那个代理并没有开着，requests 默认会读系统环境变量里的代理设置，
# 结果请求先被丢给一个不存在的代理，直接报 ProxyError。
# 火山方舟是公网地址，本来就不需要代理，所以这里强制直连。
NO_PROXY = {"http": None, "https": None}


# ============================================================
# 0. API Key 格式本地体检
# ============================================================

def validate_api_key_format(api_key):
    """
    在发请求之前，先在本地检查一下 API Key 的样子对不对。

    为什么要做这一步：
        密钥填错时服务器只会回一句英文 "The API key format is incorrect"，
        看不出到底错在哪。本地先查一遍，能直接告诉用户是"填了名称"
        还是"复制不全"，而且不用等网络请求、不消耗调用次数。

    这里只拦"肯定错"的几种情况，其他一律放过去让服务器判断——
    方舟的密钥格式以后可能变，本地规则写太死会把正常密钥误拦下来。

    返回：(是否放行, 提示说明)
          放行且没问题      → (True, None)
          放行但格式有点可疑 → (True, 提示文本)   ← 只提醒，不阻止
          明显填错          → (False, 提示文本)   ← 阻止发请求
    """
    key = (api_key or "").strip()

    if not key:
        return False, "API Key 是空的，请在 `.streamlit/secrets.toml` 的 `[ark]` 段里填上 `api_key`。"

    # ---- 明显错误一：填的是密钥"名称"而不是"值" ----
    # 方舟默认把新建的密钥命名成 api-key-年月日时分秒
    if re.fullmatch(r"api-key-\d{6,}", key):
        return False, (
            "填进去的是密钥的**名称**（`{}`），不是密钥的**值**。\n\n"
            "名称是方舟按创建时间自动生成的，不能用来调接口。"
            "请到控制台「API Key 管理」里点那条记录下面一行右边的 👁 眼睛图标，"
            "把真正的密钥值显示出来再复制。".format(key)
        )

    # ---- 明显错误二：把控制台打码的省略号一起复制进来了 ----
    if key[-1] in ".…*":
        return False, (
            "这个 API Key 末尾带了个 `{last}`，那是控制台打码显示时的省略号，"
            "说明后面还有字符没显示出来，**复制不完整**。\n\n"
            "正确做法：先点 👁 眼睛图标让密钥完整显示，再用旁边的**复制按钮**复制"
            "（不要用鼠标框选屏幕上看到的那一截）。".format(last=key[-1])
        )

    # ---- 明显错误三：短得不可能是密钥 ----
    if len(key) < 20:
        return False, (
            "这个 API Key 只有 {} 个字符，太短了，不像完整的密钥。"
            "请到控制台重新复制完整值。".format(len(key))
        )

    # ---- 以下只是"看着有点怪"，提醒一下但不阻止 ----
    # 方舟密钥主体通常是标准 UUID（8-4-4-4-12），后面可能还跟附加段
    body = key[4:] if key.lower().startswith("ark-") else key
    uuid_like = re.match(
        r"^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}(-.+)?$",
        body,
    )
    if not uuid_like:
        return True, (
            "这个 API Key 的格式和常见的方舟密钥不太一样（通常主体是 "
            "`xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx` 这种形式）。"
            "已经照原样发给服务器验证了，如果返回认证失败，就回控制台重新复制一次完整值。"
        )

    return True, None


# ============================================================
# 1. 网络会话与地址处理
# ============================================================

def _build_session():
    """
    建一个不读系统代理设置的 requests 会话。

    trust_env = False 会让 requests 忽略环境变量里的
    http_proxy / https_proxy / no_proxy 以及 ~/.netrc，
    配合请求时再显式传 proxies=NO_PROXY，双重保证直连不走代理。
    """
    session = requests.Session()
    session.trust_env = False
    return session


def _derive_urls(base_url):
    """
    根据配置里的接口地址，推算出两个接口的完整地址。

    secrets.toml 里填的是 chat/completions 的地址，
    这里把它的路径换掉，得到 responses 接口的地址，
    省得让用户在配置文件里填两个 URL。

    返回：(chat_completions 地址, responses 地址)
    """
    url = (base_url or "").strip().rstrip("/")
    if not url:
        url = "https://ark.cn-beijing.volces.com/api/v3/chat/completions"

    # 找到 /api/v3 这一段作为根，后面的路径自己拼
    marker = "/api/v3"
    idx = url.find(marker)
    if idx != -1:
        root = url[: idx + len(marker)]
    else:
        # 配置里填的可能就是根地址
        root = url

    return root + "/chat/completions", root + "/responses"


# ============================================================
# 2. 主调用函数
# ============================================================

def call_doubao(system_prompt, user_prompt, api_key, model, base_url,
                enable_web_search=True, temperature=0.7, max_tokens=16000,
                deep_thinking=False, progress_callback=None):
    """
    调用豆包模型，返回结构化结果。

    调用顺序（前面的成功就不试后面的）：
        1. Responses API + 内置联网搜索（enable_web_search=True 时）
        2. Chat Completions 不带工具（兜底，不联网）

    参数：
        system_prompt      : old_skill 业务规则组成的系统提示词
        user_prompt        : 表单内容组成的用户提示词
        api_key / model / base_url : 火山方舟配置（来自 secrets.toml）
        enable_web_search  : 是否开启联网搜索
        progress_callback  : 可选，回调函数，用来在网页上显示当前进度

    返回字典：
        ok              : 是否成功
        content         : 模型返回的正文文本
        data            : 解析后的人群包 JSON（dict，失败时 None）
        mode            : 实际生效的调用方式说明
        web_search_used : 是否真的开了联网搜索
        search_count    : 模型实际发起了几次联网搜索
        references      : 联网搜索引用的网页列表
        error           : 给用户看的错误提示（成功时 None）
        error_detail    : 原始报错详情（排查用）
        attempt_log     : 每种调用方式的尝试结果（排查用）
    """
    chat_url, responses_url = _derive_urls(base_url)
    session = _build_session()
    headers = {
        "Content-Type": "application/json",
        "Authorization": "Bearer {}".format(api_key),
    }

    # 要依次尝试的调用方式
    attempts = []
    if enable_web_search:
        attempts.append(("Responses API + 内置联网搜索", "responses", True))
    attempts.append((
        "Chat Completions（未联网，仅用模型自身知识）" if enable_web_search
        else "Chat Completions（配置里关闭了联网搜索）",
        "chat", False,
    ))

    errors = []       # 失败原因，全失败时一起返回
    attempt_log = []  # 每一步的结果，方便排查

    for mode_name, kind, use_search in attempts:
        if progress_callback:
            progress_callback("正在尝试调用方式：{}".format(mode_name))

        if kind == "responses":
            url = responses_url
            payload = _build_responses_payload(
                system_prompt, user_prompt, model, temperature, max_tokens,
                use_search, deep_thinking,
            )
        else:
            url = chat_url
            payload = _build_chat_payload(
                system_prompt, user_prompt, model, temperature, max_tokens,
                deep_thinking,
            )

        try:
            response = session.post(
                url, headers=headers, json=payload,
                timeout=REQUEST_TIMEOUT, proxies=NO_PROXY,
            )
        except requests.exceptions.ProxyError as e:
            msg = "[{}] 仍被代理拦截：{}。请检查本机是否有全局代理软件强制接管网络。".format(mode_name, e)
            errors.append(msg); attempt_log.append(msg)
            continue
        except requests.exceptions.Timeout:
            msg = "[{}] 请求超时（超过 {} 秒）".format(mode_name, REQUEST_TIMEOUT)
            errors.append(msg); attempt_log.append(msg)
            continue
        except requests.exceptions.RequestException as e:
            msg = "[{}] 网络请求失败：{}".format(mode_name, e)
            errors.append(msg); attempt_log.append(msg)
            continue

        # HTTP 状态码不是 200
        if response.status_code != 200:
            snippet = (response.text or "(无返回内容)")[:500]
            msg = "[{}] HTTP {}：{}".format(mode_name, response.status_code, snippet)
            errors.append(msg); attempt_log.append(msg)
            # 401/403 是密钥本身的问题，换接口也没用，直接结束
            if response.status_code in (401, 403):
                break
            continue

        try:
            result = response.json()
        except ValueError:
            msg = "[{}] 返回内容不是合法 JSON：{}".format(mode_name, (response.text or "")[:300])
            errors.append(msg); attempt_log.append(msg)
            continue

        # HTTP 200 但 body 里有 error
        if isinstance(result, dict) and result.get("error"):
            msg = "[{}] 接口报错：{}".format(
                mode_name, json.dumps(result["error"], ensure_ascii=False)[:400]
            )
            errors.append(msg); attempt_log.append(msg)
            continue

        # 按接口类型分别解析正文
        if kind == "responses":
            content = _extract_content_responses(result)
            references = _extract_references_responses(result)
            search_count = _count_search_calls(result)
        else:
            content = _extract_content_chat(result)
            references = []
            search_count = 0

        if not content or not content.strip():
            # 单独识别"输出被截断"这种情况，报错要说人话。
            # 典型原因：深度思考把 token 额度吃光了，正文还没开始写就被切断。
            incomplete = (result.get("incomplete_details") or {}) if isinstance(result, dict) else {}
            if incomplete.get("reason") == "length":
                msg = (
                    "[{}] 输出长度不够，模型还没写出正文就被截断了"
                    "（max_output_tokens={}）。通常是深度思考占满了额度，"
                    "把 enable_deep_thinking 关掉或调大输出上限即可。"
                ).format(mode_name, payload.get("max_output_tokens") or payload.get("max_tokens"))
            else:
                msg = "[{}] 模型没有返回正文内容：{}".format(
                    mode_name, json.dumps(result, ensure_ascii=False)[:300]
                )
            errors.append(msg); attempt_log.append(msg)
            continue

        # 把正文解析成人群包 JSON
        parsed = extract_json(content)

        attempt_log.append(
            "[{}] 成功：正文 {} 字，联网搜索 {} 次，JSON 解析{}".format(
                mode_name, len(content), search_count, "成功" if parsed else "失败"
            )
        )

        # 正文拿到了但解析不出 JSON —— 不算彻底失败，交给上层提示重试
        return {
            "ok": True,
            "content": content,
            "data": parsed,
            "mode": mode_name,
            "web_search_used": bool(use_search and search_count > 0),
            "search_count": search_count,
            "references": references,
            "error": None if parsed else "模型返回的内容无法解析成 JSON，请重试一次",
            "error_detail": None,
            "attempt_log": attempt_log,
            "raw_usage": result.get("usage"),
        }

    # 所有方式都失败
    detail = "\n".join("· " + e for e in errors)
    return {
        "ok": False,
        "content": None,
        "data": None,
        "mode": "全部调用方式均失败",
        "web_search_used": False,
        "search_count": 0,
        "references": [],
        "error": _friendly_error(detail),
        "error_detail": detail,
        "attempt_log": attempt_log,
        "raw_usage": None,
    }


# ============================================================
# 3. 两种接口的请求体
# ============================================================

def _build_responses_payload(system_prompt, user_prompt, model,
                             temperature, max_tokens, use_search,
                             deep_thinking=False):
    """
    Responses API 的请求体。

    和 Chat Completions 的区别：
      - 消息字段叫 input（不叫 messages）
      - 最大输出字段叫 max_output_tokens（不叫 max_tokens）
      - 内置联网搜索通过 tools: [{"type": "web_search"}] 开启

    关于 thinking 参数（这个坑很关键）：
        doubao-seed-2-1-pro 默认**开启深度思考**，而思考过程本身是要占
        输出 token 的。如果 max_output_tokens 给小了，模型会把额度全用在
        思考上，还没来得及写正文就被截断，返回里只有一条 reasoning，
        并且带 incomplete_details.reason = "length"，看起来像"模型什么都没返回"。
        我们这个场景（按固定 JSON 结构填人群包）提示词已经写得很细，
        不太需要深度思考，关掉之后又快又稳，所以默认关闭。
        想要更高质量可以在 secrets.toml 里把 enable_deep_thinking 打开。
    """
    payload = {
        "model": model,
        "input": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        "max_output_tokens": max_tokens,
        "temperature": temperature,
        "stream": False,
        "thinking": {"type": "enabled" if deep_thinking else "disabled"},
    }
    if use_search:
        # 内置联网搜索工具。模型会自己判断要不要搜、搜几次。
        payload["tools"] = [{"type": "web_search"}]
    return payload


def _build_chat_payload(system_prompt, user_prompt, model,
                        temperature, max_tokens, deep_thinking=False):
    """
    Chat Completions 的请求体（兜底用，不带任何工具）。

    这里特意**不加** tools 参数：
    chat/completions 本身不支持联网搜索，硬塞一个 function 形式的
    web_search 只会让模型返回一个空的 tool_call，反而把结果搞坏。

    同样默认关掉深度思考，否则响应会明显变慢（官方说关掉能快 30% 左右）。
    """
    return {
        "model": model,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        "temperature": temperature,
        "max_tokens": max_tokens,
        "thinking": {"type": "enabled" if deep_thinking else "disabled"},
    }


# ============================================================
# 4. 返回内容解析
# ============================================================

def _extract_content_chat(result):
    """从 Chat Completions 的返回里取出模型生成的正文。"""
    try:
        choices = result.get("choices") or []
        if not choices:
            return None
        message = choices[0].get("message") or {}
        content = message.get("content")

        if isinstance(content, str) and content.strip():
            return content

        # content 可能是分段列表
        if isinstance(content, list):
            parts = []
            for item in content:
                if isinstance(item, dict) and item.get("text"):
                    parts.append(item["text"])
                elif isinstance(item, str):
                    parts.append(item)
            if parts:
                return "\n".join(parts)

        # 推理模型可能把内容放在 reasoning_content
        reasoning = message.get("reasoning_content")
        if isinstance(reasoning, str) and reasoning.strip():
            return reasoning
    except Exception:
        pass
    return None


def _extract_content_responses(result):
    """
    从 Responses API 的返回里取出模型生成的正文。

    Responses API 的返回结构是一个 output 列表，里面混着好几种条目：
        {"type": "reasoning", ...}        深度思考过程
        {"type": "web_search_call", ...}  联网搜索动作
        {"type": "message", "content": [{"type": "output_text", "text": "..."}]}
    我们要的是 message 里的 output_text，其余跳过。
    """
    try:
        # 有些实现会直接给一个拼好的 output_text 字段，有就先用
        if isinstance(result.get("output_text"), str) and result["output_text"].strip():
            return result["output_text"]

        texts = []
        for item in result.get("output") or []:
            if not isinstance(item, dict):
                continue
            # 跳过思考过程和搜索动作，只要正式回复
            if item.get("type") in ("reasoning", "web_search_call", "function_call"):
                continue

            content = item.get("content")
            if isinstance(content, str) and content.strip():
                texts.append(content)
            elif isinstance(content, list):
                for part in content:
                    if isinstance(part, dict) and isinstance(part.get("text"), str):
                        texts.append(part["text"])
                    elif isinstance(part, str):
                        texts.append(part)

        if texts:
            return "\n".join(texts)
    except Exception:
        pass
    return None


def _count_search_calls(result):
    """数一下模型实际发起了几次联网搜索（用来确认联网真的生效了）。"""
    count = 0
    try:
        for item in result.get("output") or []:
            if isinstance(item, dict) and item.get("type") == "web_search_call":
                count += 1
    except Exception:
        pass
    return count


def _extract_references_responses(result):
    """
    取出联网搜索引用的网页链接。

    Responses API 把引用放在正文片段的 annotations 里，
    形如 {"type": "url_citation", "url": "...", "title": "..."}。
    取不到就返回空列表，不影响主流程。
    """
    refs = []
    seen = set()
    try:
        for item in result.get("output") or []:
            if not isinstance(item, dict):
                continue
            for part in item.get("content") or []:
                if not isinstance(part, dict):
                    continue
                for ann in part.get("annotations") or []:
                    if not isinstance(ann, dict):
                        continue
                    url = ann.get("url") or ann.get("link") or ""
                    title = ann.get("title") or ann.get("name") or ""
                    if url and url not in seen:
                        seen.add(url)
                        refs.append({"title": str(title), "url": str(url)})

        # 有的版本把搜索结果挂在 web_search_call 条目上
        for item in result.get("output") or []:
            if isinstance(item, dict) and item.get("type") == "web_search_call":
                for r in (item.get("results") or item.get("search_results") or []):
                    if isinstance(r, dict):
                        url = r.get("url") or r.get("link") or ""
                        title = r.get("title") or r.get("name") or ""
                        if url and url not in seen:
                            seen.add(url)
                            refs.append({"title": str(title), "url": str(url)})
    except Exception:
        pass
    return refs


def extract_json(text):
    """
    从模型返回的文本里把 JSON 抠出来并解析。
    模型有时会加 ```json 代码块或前后寒暄，所以要做容错。
    解析失败返回 None。
    """
    if not text:
        return None

    cleaned = text.strip()

    # 情况一：直接就是合法 JSON
    try:
        parsed = json.loads(cleaned)
        if isinstance(parsed, dict):
            return parsed
    except ValueError:
        pass

    # 情况二：被 ```json ... ``` 代码块包裹
    fence_match = re.search(r"```(?:json)?\s*(.+?)```", cleaned, re.DOTALL)
    if fence_match:
        try:
            parsed = json.loads(fence_match.group(1).strip())
            if isinstance(parsed, dict):
                return parsed
        except ValueError:
            pass

    # 情况三：文本里夹着一段 JSON —— 从第一个 { 到最后一个 } 截取
    start = cleaned.find("{")
    end = cleaned.rfind("}")
    if start != -1 and end > start:
        snippet = cleaned[start:end + 1]
        try:
            parsed = json.loads(snippet)
            if isinstance(parsed, dict):
                return parsed
        except ValueError:
            # 最后一招：去掉 JSON 里常见的尾随逗号再试
            repaired = re.sub(r",\s*([}\]])", r"\1", snippet)
            try:
                parsed = json.loads(repaired)
                if isinstance(parsed, dict):
                    return parsed
            except ValueError:
                pass

    return None


# ============================================================
# 5. 错误提示与连通性自检
# ============================================================

def _is_auth_error(detail):
    """判断失败原因是不是"密钥/端点不对"这类认证问题。"""
    text = (detail or "").lower()
    keywords = [
        "authenticationerror", "unauthorized", "api key", "apikey",
        "http 401", "http 403", "invalid key", "accessdenied",
    ]
    return any(k in text for k in keywords)


def _friendly_error(detail):
    """把接口的英文报错翻译成运营同事看得懂的提示 + 解决办法。"""
    if _is_auth_error(detail):
        return (
            "❌ 火山方舟接口拒绝了本次请求：**API Key 不正确**。\n\n"
            "最常见的原因：**把密钥的「名称」当成密钥的「值」填了**。\n"
            "在方舟控制台的 API Key 列表里，每一条有两行：\n\n"
            "- 上面一行是**名称**，系统默认按创建时间自动命名，形如 `api-key-20260930141255`\n"
            "- 下面一行才是**密钥值**，默认是打码 / 截断显示的（右边有个眼睛图标）\n\n"
            "要填的是**下面那行**。请点一下那个眼睛图标让它完整显示出来，"
            "再用复制按钮把完整值复制下来（注意别只复制屏幕上看得见的那一截，"
            "打码状态下末尾的 `.` 表示后面还有字符没显示）。\n\n"
            "然后填进项目里的 `.streamlit/secrets.toml`：\n\n"
            "```toml\n[ark]\napi_key = \"这里填完整的密钥值\"\nmodel = \"doubao-seed-2-1-pro-260628\"\n```\n\n"
            "保存后回到网页，按 `R` 键或点右上角「⋮ → Rerun」刷新即可。\n\n"
            "补充说明：`model` 填**模型 ID** 就行，不需要 `ep-` 开头的接入点 ID。\n\n"
            "----\n接口原始返回：\n" + detail
        )
    low = (detail or "").lower()
    if "nameresolutionerror" in low or "failed to resolve" in low or "nodename nor servname" in low:
        return (
            "❌ 连不上火山方舟：**电脑解析不了域名** `ark.cn-beijing.volces.com`。\n\n"
            "密钥和模型 ID 这次没有被用到，请求在找服务器地址这一步就失败了。"
            "常见原因是本机 DNS 异常、公司网络限制，或者代理软件把域名解析搞乱了。\n\n"
            "可以先试这几步：\n"
            "1. 浏览器打开 https://www.volcengine.com 看能不能进\n"
            "2. 如果开着 Clash / Surge 之类的代理，先关掉，或者让它不要接管这个域名\n"
            "3. 换一下网络（比如手机热点）后再点一次「测试接口连通性」\n\n"
            "----\n详细信息：\n" + detail
        )
    if "代理" in (detail or "") or "proxyerror" in low:
        return (
            "❌ 网络请求被本机代理拦截了。\n\n"
            "本工具已经设置成不走代理直连，如果还报这个错，说明有代理软件"
            "（Clash / Surge / V2Ray 之类）在强制接管全局网络。"
            "把它关掉或者设置成绕过 `ark.cn-beijing.volces.com` 再试。\n\n"
            "----\n详细信息：\n" + detail
        )
    if "超时" in (detail or ""):
        return (
            "❌ 调用接口超时。联网搜索 + 生成 10 个人群包比较耗时，"
            "请稍等一会儿再点一次「生成人群包」。\n\n----\n详细信息：\n" + detail
        )
    return "❌ 调用火山方舟接口失败。\n\n----\n详细原因：\n" + detail


def test_connection(api_key, model, base_url):
    """
    接口连通性自检：发一个极短的请求，只为确认密钥和端点是否可用。
    走不带联网的 Chat Completions，关掉深度思考，最省 token 也最快。

    返回：(是否连通, 提示信息)
    """
    result = call_doubao(
        system_prompt="你是一个连通性测试助手。",
        user_prompt="请只回复两个字：正常",
        api_key=api_key,
        model=model,
        base_url=base_url,
        enable_web_search=False,
        temperature=0,
        max_tokens=64,
        deep_thinking=False,
    )
    if result["ok"]:
        return True, "接口连通正常，模型返回：{}".format((result.get("content") or "").strip()[:50])
    return False, result["error"]


def test_web_search(api_key, model, base_url):
    """
    联网搜索自检：问一个必须联网才能答的问题，确认 Responses API
    的内置联网搜索是否真的生效。

    返回：(是否联网成功, 提示信息)
    """
    result = call_doubao(
        system_prompt="你是一个联网能力测试助手，回答要简短。",
        user_prompt="请联网搜索并用一句话告诉我：今天的日期是几月几号？",
        api_key=api_key,
        model=model,
        base_url=base_url,
        enable_web_search=True,
        temperature=0,
        max_tokens=4000,      # 留够额度，避免还没写正文就被截断
        deep_thinking=False,  # 关掉深度思考，自检要快
    )
    if not result["ok"]:
        return False, result["error"]

    if result["web_search_used"]:
        return True, "联网搜索正常，本次模型搜索了 {} 次，引用 {} 个网页。模型回答：{}".format(
            result["search_count"], len(result["references"]),
            (result.get("content") or "").strip()[:80],
        )
    return False, (
        "接口调用成功，但模型这次没有真的发起联网搜索（实际调用方式：{}）。"
        "生成人群包时可能只用模型自身知识。".format(result["mode"])
    )
