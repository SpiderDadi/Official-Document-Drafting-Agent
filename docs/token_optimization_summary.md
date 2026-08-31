#  每日政策全流程 — Token 优化成果汇总

> 日期：2026-08-31 | 目标：省钱 + 不降质量

---

## 一、优化前后对比

| 指标 | 优化前 | 优化后 | 节省 |
|------|--------|--------|------|
| Agent 数量 | 13 个 | 2 个 | 85% |
| Token 消耗 | ~150 万 | ~30-50 万 | 70-80% |
| 信源采集 | 7 个 agent 各搜 1 个站 | Python 脚本 requests 抓取 | 0 token |
| 文件生成 | AI 口述格式 | Python 脚本直接生成 | 0 token |
| 思维导图 | AI 重新理解生成 | 从分析结果提取结构 | 0 token |
| PPT 生成 | AI 口述格式 | python-pptx 直接生成 | 0 token |
| Word 报告 | AI 口述格式 | python-docx 直接生成 | 0 token |

---

## 二、新增脚本清单

| 脚本 | 行数 | 作用 | Token 消耗 |
|------|------|------|------------|
| `scripts/policy_scraper.py` | ~200 | 7 大信源采集 + Markdown 简报 | 0 |
| `scripts/report_generator.py` | ~200 | Word 深度分析报告生成 | 0 |
| `scripts/mindmap_generator.py` | ~130 | 思维导图大纲生成（可导入XMind） | 0 |
| `scripts/ppt_generator.py` | ~300 | PPT 自动生成（统一配色+排版） | 0 |
| `scripts/run_daily_policy.py` | ~120 | 全流程一键运行总控脚本 | 0 |
| `workflows/daily-policy-lite.js` | ~120 | 精简版工作流（2 个 agent） | ~30-50 万 |

---

## 三、如何使用

### 方式 1：Claude 工作流（推荐）

在 Claude 对话中说：
```
运行 daily-policy-lite 工作流
```

工作流会自动：
1. 搜索 7 大信源 → 生成简报
2. 对前 3 条深度分析 → 输出完整分析结果
3. 提示你运行 Python 脚本生成文件

### 方式 2：一键脚本（最快）

```bash
cd Z:\工作\CC\Official-Document-Drafting-Agent
python scripts\run_daily_policy.py
```

### 方式 3：分步执行

```bash
# Step 1: 信源采集（0 token）
python scripts\policy_scraper.py --markdown --limit 5

# Step 2: 在 Claude 中分析（让 Claude 读取 policy_items.json，进行深度分析）

# Step 3: 生成文件（0 token）
python scripts\report_generator.py -c 分析结果.json -o 政策深度分析报告.docx
python scripts\mindmap_generator.py -c 分析结果.json -o 思维导图.md
python scripts\ppt_generator.py -c 分析结果.json -o 政策解读.pptx
```

---

## 四、生成文件清单

每次全流程运行后，在 `生成\日期文件夹\` 中会生成：

| 文件 | 格式 | 内容 |
|------|------|------|
| `policy_items.json` | JSON | 信源采集原始数据 |
| `政策简报_日期.md` | Markdown | 每日政策简报 |
| `政策深度分析报告.docx` | Word | 深度分析报告（GB/T 9704-2012 格式） |
| `思维导图.md` | Markdown | 可导入 XMind/幕布的思维导图 |
| `政策解读.pptx` | PPT | 政策读演示文稿 |
| `analysis_result.json` | JSON | AI 分析结果（可复用） |

---

## 五、Skill 精简建议

### 必须保留

| Skill | 作用 | 频率 |
|-------|------|------|
| `policy` | 政策解读方法论 | 每次 |
| `official_DD` | 公文写作 | 需要时 |

### 可选保留

| Skill | 建议 |
|-------|------|
| `qiaomu-anything-to-notebooklm` | 只在需要 NotebookLM 时手动触发 |

### 不需要单独列为 skill 的内容

- 信源清单 → 写入 Python 脚本配置
- 分析框架 → 写入 AI prompt 模板
- 各种参考文档 → 合并精简到 1 个文件

---

## 六、下一步优化建议

1. **配置定时任务**：Windows Task Scheduler 每天 16:00 自动运行 `run_daily_policy.py`
2. **完善信源抓取**：部分网站可能需要 Selenium/Playwright 才能抓取动态内容
3. **增加 PDF 导出**：用 `reportlab` 或 `weasyprint` 生成 PDF 版本
4. **增加邮件通知**：生成完成后自动发送邮件摘要

---

## 七、已优化的核心原则

1. **能用代码不用 AI**：所有模板拼接、文件操作、格式转换全部用 Python
2. **合并 agent**：7 个采集合并为 1 个脚本，3 个分析合并为 1 个 agent
3. **去 JSON schema**：不在 Workflow 中强制 JSON schema，避免重试浪费
4. **一次分析多处用**：AI 分析结果保存为 JSON，所有生成脚本共享同一份数据
