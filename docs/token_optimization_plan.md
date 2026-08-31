# 📊 Token 消耗诊断与代码化替代方案

> 日期：2026-08-31 | 目标：省钱 + 不降质量

---

## 一、当前流程拆解（13 个 agent，每次 100-200 万 token）

```
daily-policy-full 工作流
├── Phase 1: 政策简报（7 个 agent，每个 50-100K）
│   ├── agent("求是杂志")     → WebSearch 搜 1 个网站，返回结构化 JSON
│   ├── agent("人民日报")     → WebSearch 搜 1 个网站，返回结构化 JSON
│   ├── agent("新华社")       → WebSearch 搜 1 个网站，返回结构化 JSON
│   ├── agent("国务院")       → WebSearch 搜 1 个网站，返回结构化 JSON
│   ├── agent("农业农村部")   → WebSearch 搜 1 个网站，返回结构化 JSON
│   ├── agent("科技部")       → WebSearch 搜 1 个网站，返回结构化 JSON
│   └── agent("发改委")       → WebSearch 搜 1 个网站，返回结构化 JSON
│
├── Phase 2: 深度分析（3 个 agent，每个 100-200K）
│   ├── agent("政策1 三轮分析") → 12 问递进，强制 JSON schema
│   ├── agent("政策2 三轮分析") → 12 问递进，强制 JSON schema
│   └── agent("政策3 三轮分析") → 12 问递进，强制 JSON schema
│
└── Phase 3: 生成输出（3 个 agent，每个 100-200K）
    ├── agent("政策解读报告+PPT+思维导图") → 纯格式转换
    ├── agent("政策解读报告+PPT+思维导图") → 纯格式转换
    └── agent("政策解读报告+PPT+思维导图") → 纯格式转换
```

---

## 二、逐项分析：哪些该用代码，哪些该用 AI

### ❌ 应该用代码解决的（完全不消耗 token）

| # | 环节 | 当前做法 | 问题 | 代码替代方案 |
|---|------|----------|------|-------------|
| 1 | **创建日期文件夹** | AI 生成路径 | AI 算日期浪费上下文 | `mkdir 生成\日期` 一行代码 |
| 2 | **7 大信源搜索** | 7 个 agent 并行 | 每个都是搜 1 个网站，重复启动开销 | **Python 脚本**：7 个 URL 用 requests 抓取标题列表，50 行代码搞定 |
| 3 | **政策汇总去重** | AI 按 JSON 排序 | 简单的按字段排序+去重 | `sorted(items, key=...)` 一行 Python |
| 4 | **生成 Markdown 简报** | AI 格式化输出 | 纯模板拼接，毫无推理价值 | Python f-string 模板，10 行代码 |
| 5 | **生成思维导图结构** | AI 重新理解政策后输出 | 分析已经做完了，只是提取层级 | **代码从分析结果中提取**：遍历 JSON 生成 `.md` 大纲文件 |
| 6 | **生成 PPT 分镜大纲** | AI 再读一遍政策重新写 | 分析结果已有，只是套表格模板 | **代码套模板**：从分析 JSON → 填充 PPT 分镜表格 |
| 7 | **生成 .pptx 文件** | AI 口述格式，再写 python-pptx | 纯机械操作 | **直接调用 python-pptx 脚本**（已有 generate_docx.py 类似模式） |
| 8 | **生成 Word 报告** | AI 口述内容，再写 python-docx | 同上 | **直接调用脚本**：分析结果 → python-docx 模板渲染 |
| 9 | **文件命名与保存** | AI 决定文件名 | 无脑操作 | 代码自动命名：`政策简报_YYYYMMDD.md` |
| 10 | **生成 NotebookLM 上传** | AI 口述怎么操作 | 已有 qiaomu skill，不需要额外 agent | 直接调用 skill，不单独开 agent |

### ✅ 必须用 AI 的（代码替代不了）

| # | 环节 | 原因 | 当前消耗 | 优化建议 |
|---|------|------|----------|----------|
| 1 | **政策理解与摘要** | 需要阅读理解能力 | 融入采集阶段 | 7 合 1，1 个 agent 完成 |
| 2 | **深度分析（三轮 12 问）** | 需要推理和洞察 | 3 个 agent 各开一次 | 改为 1 个 agent 批量分析前 3 条 |
| 3 | **判断政策重要性/评级** | 需要政治敏感性和判断力 | 融入采集阶段 | 同上，合并到 1 个 agent |
| 4 | **口播词创作** | 需要创意和网感 | 在生成输出阶段 | 保留 1 个 agent |
| 5 | **PPT 视觉设计建议** | 需要审美和传播学知识 | 在生成输出阶段 | 同上，合并到口播词 agent |

---

## 三、优化后的架构对比

### 优化前（现状）
```
13 个 agent → ~150 万 token
├── 7 个采集 agent（每个只做搜索）
├── 3 个分析 agent（每个只做分析）
└── 3 个生成 agent（每个只做格式转换）
```

### 优化后（代码 + AI 混合）
```
2-3 个 agent → ~30-50 万 token（省 70-80%）
├── 1 个 Python 脚本（0 token）
│   ├── 7 个信源网站抓取（requests）
│   ├── 汇总去重排序
│   ├── 生成 Markdown 简报
│   ├── 创建文件夹
│   └── 调用 python-pptx / python-docx 生成文件
│
├── 1 个 agent（~15-25 万 token）
│   └── 简报 → 理解 + 摘要 + 评级 + 深度分析（一次性完成）
│
└── 1 个 agent（~10-15 万 token）
    └── 分析结果 → 口播词 + PPT 视觉建议 + NotebookLM
```

---

## 四、可以完全代码化的模块清单

### 模块 1：信源采集脚本 `policy_scraper.py`（约 100 行）

```python
# 功能：从 7 大信源网站抓取最新标题和摘要
# 输入：无
# 输出：policy_items.json（供 AI 分析用）
# 替代：7 个 agent → 0 个 agent
```

### 模块 2：简报生成脚本 `brief_generator.py`（约 50 行）

```python
# 功能：将政策条目汇总成 Markdown 简报
# 输入：policy_items.json
# 输出：政策简报_YYYYMMDD.md
# 替代：AI 模板拼接 → 0 token
```

### 模块 3：思维导图生成脚本 `mindmap_generator.py`（约 80 行）

```python
# 功能：从分析结果生成思维导图大纲
# 输入：analysis_results.json
# 输出：思维导图_政策名.md（可导入 XMind/幕布）
# 替代：1 个 agent → 0 个 agent
```

### 模块 4：PPT 自动生成脚本 `ppt_generator.py`（约 200 行）

```python
# 功能：从分析结果自动生成 PPT
# 输入：analysis_results.json + ppt_outline.json
# 输出：政策解读_政策名.pptx
# 替代：1 个 agent 口述 → 0 个 agent
```

### 模块 5：Word 报告生成脚本 `report_generator.py`（约 100 行）

```python
# 功能：从分析结果自动生成 Word 报告
# 输入：analysis_results.json
# 输出：政策深度分析报告_政策名.docx
# 替代：1 个 agent 口述 → 0 个 agent
```

---

## 五、Skill 精简建议

### 必须保留（2 个）

| Skill | 作用 | 为什么保留 |
|-------|------|-----------|
| `policy` | 政策解读方法论、口播词、PPT 分镜指南 | 核心的分析框架和传播方法 |
| `official_DD` | 公文写作 | 需要产出正式报告时用到 |

### 可以保留但不每次用的

| Skill | 作用 | 建议 |
|-------|------|------|
| `qiaomu-anything-to-notebooklm` | 上传到 NotebookLM | 只在需要 NotebookLM 深度分析时手动触发，不用写进日常流程 |

### 不需要单独列为 skill 的

| 内容 | 建议 |
|------|------|
| 信源清单（sources.md） | 写入 Python 脚本配置，不用 AI 读 |
| 分析框架（frameworks.md） | 写入 AI 的 prompt 模板，不用单独 skill |
| 各种参考文档 | 合并精简到 1 个文件，按需读取 |

---

## 六、立即可做的优化（不写代码，只改工作流）

如果暂时不想写 Python 脚本，只改 `daily-policy-full.js` 就能省 50%：

1. **7 个 agent 合并为 1 个**：一次搜索覆盖所有信源
2. **3 个分析 agent 合并为 1 个**：批量处理前 3 条
3. **3 个生成 agent 合并为 1 个**：一次性输出所有材料
4. **去掉 JSON schema**：让 agent 直接输出文本，避免重试浪费

优化后：**3 个 agent → 约 50 万 token（省 67%）**

---

## 七、终极方案（写代码，省 80%）

把整个流程拆成：

```
信源采集（Python） → AI 分析（1个agent） → 文件生成（Python）
   0 token            ~20-30 万 token         0 token
```

**每天只消耗 20-30 万 token，省 80%+**

---

## 八、建议行动路线

1. **第一步（今天）**：只改工作流，合并 agent，立省 50%
2. **第二步（本周）**：写 2 个核心 Python 脚本（采集 + 文件生成），省 80%
3. **第三步（下周）**：全套自动化，配置 cron 每天 16:00 自动跑

哥哥觉得先做哪一步？嘻嘻 🚀
