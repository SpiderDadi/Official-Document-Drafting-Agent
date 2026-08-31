export const meta = {
  name: 'daily-policy-lite',
  description: '精简版每日政策全流程：2个agent完成采集+分析（省70-80% token），Python脚本生成文件',
  phases: [
    { title: '采集简报', detail: '1个agent搜索7信源，生成简报JSON' },
    { title: '深度分析', detail: '1个agent对前3条做三轮递进分析，输出完整JSON' },
    { title: '文件生成', detail: '用户运行Python脚本生成Word/思维导图（0 token）' },
  ],
}

/**
 * 精简版每日政策全流程工作流
 *
 * 核心优化：
 * - 原来 13 个 agent → 现在 2 个 agent
 * - Token 消耗从 ~150 万 → ~30-50 万（省 70-80%）
 * - 文件生成（Word/思维导图/PPT）用 Python 脚本
 *
 * 流程：
 * 1. agent 1: 搜索7大信源 → 输出简报 JSON
 * 2. agent 2: 简报 + 三轮深度分析 → 输出完整分析 JSON
 * 3. 用户运行 Python 脚本 → 生成 Word / 思维导图 / PPT
 *
 * 注意：Workflow 脚本不支持直接调用 Bash，所以文件生成
 * 由用户在工作流结束后手动运行脚本（或我后续帮你执行）。
 */

// ============================================================
// Phase 1: 采集与简报
// ============================================================

phase('采集简报')

log('📡 开始搜索7大信源...')

const briefResult = await agent(`今天是政策信源采集任务。请搜索以下7大官方信源的最新政策动态，汇总成结构化简报。

信源清单：
1. 求是杂志（qstheory.cn）— 最新一期头条
2. 人民日报（paper.people.com.cn）— 头版、评论版
3. 新华社（news.cn）— 权威发布
4. 国务院（gov.cn/zhengce）— 最新政策文件
5. 农业农村部（moa.gov.cn）— 最新政策
6. 科技部（most.gov.cn）— 最新科技政策
7. 发改委（ndrc.gov.cn）— 最新经济政策

要求：
- 每条包含：title（标题）、source（信源名）、summary（100字以内摘要）、keyPoints（关键要点一句话）、rating（评级：★★★★★到★☆☆☆☆）
- 汇总5-8条，覆盖至少3个不同领域（经济/科技/农业/民生等）
- 按重要性排序（★★★★★的排前面）

请直接输出JSON数组格式，不要其他内容：
[{"title":"...","source":"...","summary":"...","keyPoints":"...","rating":"★★★★★"}]`, {
    label: '信源采集',
    phase: '采集简报',
  })

const briefItems = Array.isArray(briefResult) ? briefResult : (briefResult?.brief || [])
log(`✅ 采集完成，获得 ${briefItems.length} 条政策简报`)

// ============================================================
// Phase 2: 深度分析
// ============================================================

phase('深度分析')

const top3 = briefItems.slice(0, 3)
if (top3.length === 0) {
  log('⚠️ 未获取到有效政策简报，跳过深度分析')
  return { date: new Date().toISOString().split('T')[0], briefItems: [], policies: [] }
}

log(`🔍 对前 ${top3.length} 条重点政策进行深度分析...`)

// 构建精简的输入（只传必要字段，减少 context 膨胀）
const policyInput = top3.map((p, i) =>
  `${i+1}. 【${p.source}】${p.title}\n   摘要：${p.summary}\n   要点：${p.keyPoints}\n   评级：${p.rating}`
).join('\n\n')

const analysisResult = await agent(`你是政策深度分析师。请对以下${top3.length}条今日重点政策进行逐条分析：

${policyInput}

每条政策按三轮递进方式分析：

第一轮（概览与框架，4问必答）：
1. 这个政策的核心主题和目的是什么？
2. 政策整体结构和逻辑框架是什么？
3. 提出了哪些核心论点和主张？
4. 最具颠覆性或创新性的内容是什么？

第二轮（深度挖掘，5问必答）：
5. 政策的论证逻辑和前提假设是什么？
6. 引用了哪些关键数据或案例？
7. 是否存在内部矛盾或争议点？
8. 最独特的贡献或核心洞察是什么？
9. 如果要提出最尖锐的批评，会是什么？

第三轮（综合与反刍，3问必答）：
10. 读者最应该带走的一个认知改变是什么？
11. 可以提取出哪些可操作的行动指南？
12. 用三个最有力的理由说服别人关注这个政策

请输出JSON格式（不要其他内容）：
{"policies":[{"title":"...","source":"...","summary":"...","keyPoints":"...","rating":"...","analysis":{"summary":"一段话总结","qa":[{"round":"第一轮","question":"...","answer":"..."}],"actionItems":["建议1","建议2"]}}]}`, {
    label: '深度分析',
    phase: '深度分析',
  })

const policies = analysisResult?.policies || []
log(`✅ 深度分析完成，分析 ${policies.length} 条政策`)

// ============================================================
// Phase 3: 文件生成指引
// ============================================================

phase('文件生成')

const dateFolder = getDateFolder()
const outputDir = `Z:\\工作\\CC\\Official-Document-Drafting-Agent\\生成\\${dateFolder}`

log(`📁 输出目录：${outputDir}`)
log('')
log('🔧 接下来请运行以下 Python 脚本生成文件（0 token 消耗）：')
log('')
log(`  1. 生成 Markdown 简报`);
log(`     → 简报已在工作流中生成，直接保存即可`)
log('')
log(`  2. 生成 Word 深度分析报告`)
log(`     python scripts\\report_generator.py --json '<完整分析JSON>' --output "${outputDir}\\政策深度分析报告.docx"`)
log('')
log(`  3. 生成思维导图大纲`)
log(`     python scripts\\mindmap_generator.py --json '<完整分析JSON>' --output "${outputDir}\\思维导图.md"`)
log('')

// 生成 Markdown 简报内容
const briefMD = briefItems.map((item, i) =>
  `### ${i + 1}. ${item.title}\n- **信源**：${item.source}\n- **摘要**：${item.summary}\n- **要点**：${item.keyPoints}\n- **评级**：${item.rating}\n`
).join('\n---\n\n')

const briefContent = `# ${dateFolder} 政策与宏观动态简报\n\n${briefMD}`

log('\n' + '='.repeat(60))
log('📋 今日政策简报')
log('='.repeat(60))
log(briefContent)
log('\n' + '='.repeat(60))
log('✅ 精简版全流程完成！（2 个 agent，省 70-80% token）')
log(`📁 文件生成目录：${outputDir}`)
log('='.repeat(60))

return {
  date: dateFolder,
  outputDir,
  totalItems: briefItems.length,
  analyzedItems: policies.length,
  brief: briefItems,
  policies,
}
