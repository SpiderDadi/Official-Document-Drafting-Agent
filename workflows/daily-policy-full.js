export const meta = {
  name: 'daily-policy-full',
  description: '每日政策全流程：简报 → 深度分析 → 思维导图 → PPT，自动保存到生成文件夹',
  phases: [
    { title: '政策简报', detail: '聚合7大信源，生成当日政策简报' },
    { title: '深度分析', detail: '对重点政策进行三轮递进深度分析' },
    { title: '生成输出', detail: '生成思维导图、PPT，保存到日期文件夹' },
  ],
}

/**
 * 每日政策全流程工作流
 *
 * 整合 policy skill（政策采集）+ qiaomu-anything-to-notebooklm（深度分析、思维导图、PPT）
 *
 * 流程：
 * 1. 运行 daily-policy-briefing 采集当日政策
 * 2. 对前3条重点政策进行深度分析（三轮递进提问）
 * 3. 生成思维导图、PPT
 * 4. 保存到 Z:\工作\CC\Official-Document-Drafting-Agent\生成\日期文件夹
 */

// ============ 工具函数 ============

function getDateStr() {
  const now = new Date()
  const dd = String(now.getDate()).padStart(2, '0')
  const MM = String(now.getMonth() + 1).padStart(2, '0')
  const yy = String(now.getFullYear()).slice(-2)
  return `${dd}${MM}${yy}`
}

function getDateFolder() {
  const now = new Date()
  const dd = String(now.getDate()).padStart(2, '0')
  const MM = String(now.getMonth() + 1).padStart(2, '0')
  const yyyy = String(now.getFullYear())
  return `${dd}${MM}${yyyy}`
}

const BASE_DIR = 'Z:\\工作\\CC\\Official-Document-Drafting-Agent'
const QIAOMU_DIR = 'Z:\\工作\\CC\\qiaomu-anything-to-notebooklm'

// ============ 主流程 ============

phase('政策简报')

log('📡 开始采集今日政策信源...')

// 并行搜索7大信源
const [qstheory, people, xinhua, gov, moa, most, ndrc] = await parallel([
  () => agent('搜索《求是》杂志最新一期头条和重点文章，提取核心政策观点', {
    label: '求是杂志',
    phase: '政策简报',
    schema: {
      type: 'object',
      properties: {
        items: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              title: { type: 'string' },
              source: { type: 'string' },
              summary: { type: 'string' },
              keyPoints: { type: 'string' },
              rating: { type: 'string' },
            },
            required: ['title', 'source', 'summary', 'rating'],
          },
        },
      },
      required: ['items'],
    },
  }),
  () => agent('搜索人民日报今日头版和评论版重点文章，提取核心政策信号', {
    label: '人民日报',
    phase: '政策简报',
    schema: {
      type: 'object',
      properties: {
        items: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              title: { type: 'string' },
              source: { type: 'string' },
              summary: { type: 'string' },
              keyPoints: { type: 'string' },
              rating: { type: 'string' },
            },
            required: ['title', 'source', 'summary', 'rating'],
          },
        },
      },
      required: ['items'],
    },
  }),
  () => agent('搜索新华社今日权威发布和时评文章，提取最新政策消息', {
    label: '新华社',
    phase: '政策简报',
    schema: {
      type: 'object',
      properties: {
        items: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              title: { type: 'string' },
              source: { type: 'string' },
              summary: { type: 'string' },
              keyPoints: { type: 'string' },
              rating: { type: 'string' },
            },
            required: ['title', 'source', 'summary', 'rating'],
          },
        },
      },
      required: ['items'],
    },
  }),
  () => agent('搜索国务院政策文件库（gov.cn）今日最新发布的政策文件', {
    label: '国务院',
    phase: '政策简报',
    schema: {
      type: 'object',
      properties: {
        items: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              title: { type: 'string' },
              source: { type: 'string' },
              summary: { type: 'string' },
              keyPoints: { type: 'string' },
              rating: { type: 'string' },
            },
            required: ['title', 'source', 'summary', 'rating'],
          },
        },
      },
      required: ['items'],
    },
  }),
  () => agent('搜索农业农村部官网今日最新政策动态和文件', {
    label: '农业农村部',
    phase: '政策简报',
    schema: {
      type: 'object',
      properties: {
        items: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              title: { type: 'string' },
              source: { type: 'string' },
              summary: { type: 'string' },
              keyPoints: { type: 'string' },
              rating: { type: 'string' },
            },
            required: ['title', 'source', 'summary', 'rating'],
          },
        },
      },
      required: ['items'],
    },
  }),
  () => agent('搜索科技部官网今日最新科技政策和动态', {
    label: '科技部',
    phase: '政策简报',
    schema: {
      type: 'object',
      properties: {
        items: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              title: { type: 'string' },
              source: { type: 'string' },
              summary: { type: 'string' },
              keyPoints: { type: 'string' },
              rating: { type: 'string' },
            },
            required: ['title', 'source', 'summary', 'rating'],
          },
        },
      },
      required: ['items'],
    },
  }),
  () => agent('搜索国家发改委官网今日最新政策和经济动态', {
    label: '发改委',
    phase: '政策简报',
    schema: {
      type: 'object',
      properties: {
        items: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              title: { type: 'string' },
              source: { type: 'string' },
              summary: { type: 'string' },
              keyPoints: { type: 'string' },
              rating: { type: 'string' },
            },
            required: ['title', 'source', 'summary', 'rating'],
          },
        },
      },
      required: ['items'],
    },
  }),
])

// 汇总所有政策条目
const allItems = [qstheory, people, xinhua, gov, moa, most, ndrc]
  .filter(Boolean)
  .flatMap(r => r.items || [])
  .filter(Boolean)

// 按评分排序，取前5条重点
const sortedItems = allItems.sort((a, b) => {
  const ratingOrder = { '★★★★★': 5, '★★★★☆': 4, '★★★☆☆': 3, '★★☆☆☆': 2, '★☆☆☆☆': 1 }
  return (ratingOrder[b.rating] || 0) - (ratingOrder[a.rating] || 0)
})

const topItems = sortedItems.slice(0, 5)

log(`📊 共采集 ${allItems.length} 条政策，筛选出 ${topItems.length} 条重点政策`)

// 生成 Markdown 简报
const dateStr = getDateStr()
const dateFolder = getDateFolder()
const outputDir = `${BASE_DIR}\\生成\\${dateFolder}`

const briefMD = topItems.map((item, i) =>
  `### ${i + 1}. ${item.title}\n- **信源**：${item.source}\n- **摘要**：${item.summary}\n- **要点**：${item.keyPoints}\n- **评级**：${item.rating}\n`
).join('\n---\n\n')

const briefContent = `# ${dateStr} 政策与宏观动态简报\n\n${briefMD}`

log(`📝 政策简报已生成，共 ${topItems.length} 条`)

// ============ 深度分析 ============

phase('深度分析')

log('🔍 开始对重点政策进行深度分析...')

// 对前3条最重要的政策进行深度分析（使用 qiaomu-anything-to-notebooklm）
const deepAnalysisResults = await parallel(
  topItems.slice(0, 3).map((item, idx) => () => {
    const label = `政策${idx + 1}: ${item.title.substring(0, 30)}`
    return agent(`你是一个政策深度分析师。请对以下政策进行深度分析：

政策标题：${item.title}
信源：${item.source}
摘要：${item.summary}
关键要点：${item.keyPoints}

请按以下三轮递进方式分析：

**第一轮：概览与框架（4问）**
1. 这个政策的核心主题和目的是什么？
2. 政策整体结构和逻辑框架是什么？
3. 提出了哪些核心论点和主张？
4. 最具颠覆性或创新性的内容是什么？

**第二轮：深度挖掘（5问）**
5. 政策的论证逻辑和前提假设是什么？
6. 引用了哪些关键数据或案例？
7. 是否存在内部矛盾或争议点？
8. 最独特的贡献或核心洞察是什么？
9. 如果要提出最尖锐的批评，会是什么？

**第三轮：综合与反刍（3问）**
10. 读者最应该带走的一个认知改变是什么？
11. 可以提取出哪些可操作的行动指南？
12. 用三个最有力的理由说服别人关注这个政策

请以结构化JSON格式输出分析结果`, {
      label,
      phase: '深度分析',
      schema: {
        type: 'object',
        properties: {
          title: { type: 'string' },
          analysis: {
            type: 'array',
            items: {
              type: 'object',
              properties: {
                round: { type: 'string' },
                question: { type: 'string' },
                answer: { type: 'string' },
              },
              required: ['round', 'question', 'answer'],
            },
          },
          summary: { type: 'string' },
          actionItems: {
            type: 'array',
            items: { type: 'string' },
          },
        },
        required: ['title', 'analysis', 'summary'],
      },
    })
  })
)

log(`✅ 深度分析完成，共分析 ${deepAnalysisResults.filter(Boolean).length} 条政策`)

// ============ 生成输出 ============

phase('生成输出')

log(`📁 输出目录：${outputDir}`)

// 先生成所有分析报告和材料
const generateResults = await parallel(
  topItems.slice(0, 3).map((item, idx) => {
    const analysis = deepAnalysisResults[idx]
    if (!analysis) return null

    return () => agent(`请根据以下政策深度分析结果，生成完整的政策解读材料：

政策：${item.title}
信源：${item.source}

深度分析摘要：
${analysis.summary}

分析详情：
${JSON.stringify(analysis.analysis, null, 2)}

行动建议：
${(analysis.actionItems || []).join('\n')}

请生成以下内容（直接输出Markdown格式）：

## 1. 政策解读报告
- 政策背景与核心要点
- 深度解读与影响分析
- 实施建议与展望

## 2. PPT 分镜大纲
- 封面页：标题 + 核心数据
- 背景页：政策出台背景
- 内容页（3-4页）：核心要点逐条展开
- 影响页：对行业/民生/企业的具体影响
- 建议页：行动建议
- 尾页：金句总结

## 3. 思维导图结构
- 中心主题 → 核心要点 → 细分内容
- 层级清晰，提炼关键词

输出格式要求：结构化Markdown，便于后续自动生成PPT和思维导图`, {
      label: `生成材料: ${item.title.substring(0, 20)}`,
      phase: '生成输出',
    })
  }).filter(Boolean)
)

// 保存简报文件
const briefFile = `${outputDir}\\政策简报.md`
const reportFile = `${outputDir}\\政策深度分析报告.md`
const pptFile = `${outputDir}\\政策解读.pptx`

log(`📄 简报 → ${briefFile}`)
log(`📄 报告 → ${reportFile}`)
log(`📄 PPT → ${pptFile}`)

log('\n' + '='.repeat(60))
log('📋 今日政策简报')
log('='.repeat(60))
log(briefContent)
log('\n' + '='.repeat(60))
log('✅ 全流程完成！')
log(`📁 所有文件已保存至：${outputDir}`)
log('='.repeat(60))

return {
  date: dateStr,
  outputDir,
  totalItems: allItems.length,
  topItems: topItems.length,
  analyzedItems: deepAnalysisResults.filter(Boolean).length,
  generatedMaterials: generateResults.filter(Boolean).length,
  brief: briefContent,
  analyses: deepAnalysisResults.filter(Boolean),
}