# 美图内部 · 广告人群包生成工具

根据品牌广告需求，自动生成「美图定制人群包圈选规则」Excel 文件。
业务规则完全沿用 `old_skill/` 里的那套 skill（六大圈选维度、10 个人群包标准组合、Excel 标准格式）。

---

## 一、怎么启动（每次用之前做这一步）

打开「终端」，把下面两行**逐行**复制进去回车：

```bash
cd /Users/meitu/Desktop/resource_builder
.venv/bin/streamlit run app.py
```

浏览器会自动打开 `http://localhost:8501`。

关闭工具：回到终端窗口，按 `Control + C`。

---

## 二、怎么用

1. **登录**：输入密码 `company2026`
2. **填表单**：品牌名称、广告产品/活动、广告目标、目标人群描述是必填；
   合作内容、竞品、预算、地区、补充说明可以不填（不填的 AI 会自己判断补上）
3. **点「生成人群包」**：AI 会先联网搜索行业人群洞察和竞品投放策略，
   再按 old_skill 的规则生成 10 个人群包（约 45-55 个子标签），需要 1-3 分钟
4. **下载 Excel**：页面上会出现下载按钮，文件同时自动存到 `output/` 文件夹
5. **历史记录**：页面最下方按时间倒序列出最近 20 条，每条旁边都能直接重新下载

---

## 三、⚠️ 重要：还差一个完整的 API Key

### 模型标识已经配好了

`model` 现在填的是 **`doubao-seed-2-1-pro-260628`**，这是 Doubao-Seed-2.1-pro 在方舟的模型 ID。

方舟的 `model` 字段支持两种填法，二选一：

| 填法 | 样子 | 要不要建接入点 |
|---|---|---|
| 模型 ID（现在用的） | `doubao-seed-2-1-pro-260628` | **不需要** |
| 推理接入点 ID | `ep-20260930150000-abcde` | 需要先在控制台手动创建 |

所以**你找不到 `ep-` 开头的 ID 是正常的**，用模型 ID 直接调就行。

### 还需要做的：把 API Key 换成「密钥值」

现在 `api_key` 填的是 `api-key-20260930141255`，接口一直返回：

```
HTTP 401  AuthenticationError: The API key format is incorrect
```

原因是：**这串是密钥的「名称」，不是密钥的「值」。**

方舟控制台的 API Key 列表里，每条记录是上下两行：

- **上面一行**是名称，系统按创建时间自动命名 → `api-key-20260930141255`
  （`20260930141255` 就是 2026-09-30 14:12:55 这个创建时间）
- **下面一行**才是真正的密钥值，默认打码显示，右边有个 👁 眼睛图标 →
  `ark-888c4de7-de4d-4dbd-9d16-2e05.`（末尾那个 `.` 表示后面还有字符被省略了）

### 操作步骤

1. 打开火山方舟控制台 → 左侧 **「API Key 管理」**
2. 找到那条 `api-key-20260930141255`，点它**下面那行**右边的 **👁 眼睛图标**，让密钥完整显示
3. 用旁边的**复制按钮**复制完整值（别手动框选屏幕上那一截，打码状态下是不全的）
4. 用「文本编辑」打开 `/Users/meitu/Desktop/resource_builder/.streamlit/secrets.toml`，
   把 `api_key` 那一行的引号里换成刚复制的完整值，保存：
   ```toml
   [ark]
   api_key = "这里粘贴完整的密钥值"
   model = "doubao-seed-2-1-pro-260628"
   ```
5. 回到网页，按 `R` 键刷新（或点右上角「⋮ → Rerun」）
6. 展开 **「🔧 接口配置与连通性自检」** → 点「测试接口连通性」，
   显示绿色「接口连通正常」就配好了

### 在此之前可以用「演示模式」

表单底部有个「演示模式」勾选框。勾上之后不调用接口、不消耗次数，
用模板数据把「填表 → 生成 Excel → 下载 → 历史记录」整条流程完整跑一遍，
方便先确认工具能用、或者演示给同事看。

演示模式产出的 Excel 在标题、数据说明、备注里都会标注「演示样例」，不要直接给客户。

---

## 四、文件都是干什么的

| 文件 / 文件夹 | 作用 |
|---|---|
| `app.py` | 网页主程序（界面、登录、表单、结果展示、历史记录） |
| `skill_rules.py` | 实时读取 `old_skill/` 的业务规则，拼成 AI 的 system prompt |
| `ark_client.py` | 调用火山方舟豆包接口（开联网搜索）、解析返回、错误提示 |
| `excel_builder.py` | 按 old_skill 标准格式写 `.xlsx`（openpyxl） |
| `history_store.py` | 历史记录读写、每月调用次数计数 |
| `demo_sample.py` | 演示模式用的模板示例数据 |
| `test_excel.py` | 本地自测脚本，不调接口不花钱：`.venv/bin/python test_excel.py` |
| `.streamlit/secrets.toml` | **密钥和密码都在这里**，不要外发、不要提交到 Git |
| `.streamlit/config.toml` | 网页配色主题（美图红 #F11D48） |
| `old_skill/` | 原来的业务 skill，是 AI 的规则来源，**不要删** |
| `output/` | 生成的 Excel 文件都存在这里 |
| `history.json` | 历史记录 |
| `usage.json` | 当月调用次数（次月自动清零） |

---

## 五、几个设置在哪改

打开 `.streamlit/secrets.toml`：

- **改登录密码**：改 `app_password` 那一行
- **改每月调用上限**（现在是 500）：改 `monthly_call_limit`
- **关掉联网搜索**：把 `enable_web_search` 改成 `false`

改完保存，刷新网页生效。

---

## 六、业务规则要调整怎么办

**不用改代码**。直接编辑 `old_skill/` 里的三个文件：

- `old_skill/SKILL.md` —— 工作流程、六大维度定义、人群包类型对照表
- `old_skill/references/meitu_audience_format.md` —— Excel 格式、列定义、样式规范
- `old_skill/references/case_studies.md` —— 历史交付案例（AI 会参考这些案例的风格）

网页每次生成时都会**重新读一遍**这三个文件，保存后下次生成就生效了。

---

## 七、常见问题

**页面打不开 / 终端报错 `command not found`**
确认第一行 `cd` 路径没打错，并且用的是 `.venv/bin/streamlit` 而不是 `streamlit`。

**点了生成，等了很久最后报错**
先展开「🔧 接口配置与连通性自检」点一下测试。如果是红色的认证错误，就是第三节那个 Key 的问题。

**报错说 AI 返回内容没法解析**
偶发情况（模型输出格式跑偏了），再点一次「生成人群包」就行。

**历史记录里显示「文件已丢失」**
说明 `output/` 里对应的 Excel 被删掉或移走了。记录还在，文件需要重新生成。

**依赖坏了想重装**
```bash
cd /Users/meitu/Desktop/resource_builder
rm -rf .venv
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
```
