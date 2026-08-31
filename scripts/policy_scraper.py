#!/usr/bin/env python3
"""
政策信源采集脚本
从 7 大官方信源网站抓取最新政策标题和摘要。

用法：
  python policy_scraper.py
  python policy_scraper.py --output policy_items.json
  python policy_scraper.py --limit 3      # 每站最多取 3 条

输出：
  生成 JSON 文件，供后续 AI 分析使用
  同时生成 Markdown 简报（零 token 消耗）

设计原则：
  - 纯代码实现，0 token 消耗
  - 优先 requests 直接抓取，降级到 WebSearch
  - 所有数据本地缓存，避免重复请求
"""

import json
import argparse
import sys
import re
from datetime import datetime
from pathlib import Path

try:
    import requests
    HAS_REQUESTS = True
except ImportError:
    HAS_REQUESTS = False

try:
    from bs4 import BeautifulSoup
    HAS_BS4 = True
except ImportError:
    HAS_BS4 = False

# ============================================================
# 信源配置
# ============================================================

SOURCES = [
    {
        "id": "qstheory",
        "name": "求是杂志",
        "url": "https://www.qstheory.cn/",
        "fallback_search": "求是杂志 最新 2026 头条 政策",
    },
    {
        "id": "people",
        "name": "人民日报",
        # 使用人民网首页动态页面，不依赖硬编码日期
        "url": "http://www.people.com.cn/",
        "alt_urls": [
            "http://politics.people.com.cn/",
            "http://cpc.people.com.cn/",
        ],
        "fallback_search": "人民日报 今日 头版 政策 2026",
    },
    {
        "id": "xinhua",
        "name": "新华社",
        "url": "https://www.news.cn/",
        "fallback_search": "新华社 权威发布 政策 2026",
    },
    {
        "id": "gov",
        "name": "国务院",
        # 修正 URL：优先使用 zhengce（政策列表页），其次 zuixin（最新政策）
        "url": "https://www.gov.cn/zhengce/",
        "alt_urls": [
            "https://www.gov.cn/zhengce/list.htm",
            "https://www.gov.cn/zhengce/zuixin/",
        ],
        "fallback_search": "国务院 最新政策文件 gov.cn 2026",
    },
    {
        "id": "moa",
        "name": "农业农村部",
        # 政务动态页列表完整，标题在 a[title] 属性中
        "url": "https://www.moa.gov.cn/xw/zwdt/",
        "alt_urls": [
            "https://www.moa.gov.cn/",
            "https://www.moa.gov.cn/xw/zxfb/",
        ],
        "fallback_search": "农业农村部 最新政策 动态 2026",
    },
    {
        "id": "most",
        "name": "科技部",
        # kjbgz 为科技部最新工作动态页，时效性好
        "url": "https://www.most.gov.cn/kjbgz/",
        "alt_urls": [
            "https://www.most.gov.cn/xxgk/xinxifenlei/fdzdgknr/fgzc/",
            "https://www.most.gov.cn/xxgk/xinxifenlei/fdzdgknr/fgzc/gfxwj/",
        ],
        "fallback_search": "科技部 最新政策 动态 2026",
    },
    {
        "id": "ndrc",
        "name": "发改委",
        "url": "https://www.ndrc.gov.cn/",
        "fallback_search": "发改委 最新政策 经济动态 2026",
    },
]

# 通用请求头（模拟浏览器）
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
}

TIMEOUT = 15  # 请求超时秒数


def fetch_url(url, timeout=TIMEOUT):
    """
    抓取网页内容，返回 (html_text, error) 元组。
    失败时返回 (None, error_message)。
    """
    try:
        resp = requests.get(url, headers=HEADERS, timeout=timeout)
        resp.encoding = resp.apparent_encoding or "utf-8"
        if resp.status_code == 200:
            return resp.text, None
        return None, f"HTTP {resp.status_code}"
    except requests.RequestException as e:
        return None, str(e)


def extract_full_title(a_tag):
    """
    从 <a> 标签中提取完整标题。
    优先使用 title 属性（完整标题），降级到可见文本。
    同时清理截断符号（... 或 …）和日期后缀。
    """
    title_attr = (a_tag.get("title") or "").strip()
    visible_text = (a_tag.get_text(strip=True) or "").strip()

    # 优先用 title 属性（通常包含完整标题）
    if title_attr and len(title_attr) > len(visible_text):
        best = title_attr
    else:
        best = visible_text

    if not best:
        return ""

    # 清理截断符号
    best = best.rstrip(".")
    best = best.rstrip("。")
    best = best.rstrip("…")
    best = re.sub(r'\.{2,}$', '', best)

    # 清理日期后缀 (如 "...08-28" -> "..." + 日期)
    best = re.sub(r'[.\d]{5,10}$', '', best)

    # 去掉空格和多余空白
    best = ' '.join(best.split())

    return best


def parse_html_items(html, source_name):
    """
    从 HTML 中提取标题列表。
    通用策略：提取所有 <a> 标签的文本，过滤掉导航和无意义内容。
    标题优先取 title 属性（完整标题），避免 CSS 截断导致的省略号。
    """
    if not html:
        return []

    soup = BeautifulSoup(html, "html.parser")

    # 通用策略：收集所有有文本的链接
    items = []
    seen_titles = set()

    # 按信源配置不同的选择器优先级
    # 农业农村部 - 政务动态页用 .ztlb 列表项
    if source_name == "农业农村部":
        selectors = [
            ".ztlb a",
            ".zwdt-item a",
            ".list li a",
            ".news-list a",
            "ul.list li a",
        ]
    # 国务院 - 政策列表通常在 zhengce-content，但主页面是.list li a
    elif source_name == "国务院":
        selectors = [
            ".list li a",
            ".zhengce-list li a",
            "#content ul li a",
            ".policy-item a",
        ]
    # 科技部 - 政策法规页面
    elif source_name == "科技部":
        selectors = [
            ".main-content table tr td a",
            ".flfg-list li a",
            ".most-table tbody tr td a",
            ".flfg-item a",
            ".news-item a",
            ".list li a",
        ]
    else:
        # 其他信源使用通用选择器
        selectors = [
            ".list a", ".news-list a", ".newslist a", ".article-list a",
            ".content-list a", ".main-news a", ".news-item a", ".txtList a",
            "ul.list li a", "ul li a",
            ".news a", ".focus a", ".headline a",
            "#content a", ".text a",
        ]

    for selector in selectors:
        links = soup.select(selector)
        for a in links:
            text = extract_full_title(a)
            if text and len(text) > 6 and text not in seen_titles:
                seen_titles.add(text)
                items.append({
                    "title": text,
                    "source": source_name,
                    "url": a.get("href", ""),
                    "summary": "",
                })
        if items:
            break

    # 降级：所有 <a> 标签（排除导航链接）
    if not items:
        nav_keywords = ['首页', '新闻', '专题', '图集', '视频', '播客', '图片', '直播', '论坛', '访谈']
        for a in soup.find_all("a"):
            text = extract_full_title(a)
            if text and len(text) > 6 and len(text) < 200 and text not in seen_titles:
                # 排除导航链接
                is_nav = any(kw in text for kw in nav_keywords)
                if not is_nav:
                    seen_titles.add(text)
                    items.append({
                        "title": text,
                        "source": source_name,
                        "url": a.get("href", ""),
                        "summary": "",
                    })

    return items[:50]  # 最多返回 50 条，避免太多


def scrape_source(source_cfg, max_items=5):
    """
    抓取单个信源，返回政策条目列表。
    支持主 URL + alt_urls 备选地址，全部失败时返回占位。
    """
    source_name = source_cfg["name"]

    # 构建候选 URL 列表：主 URL + 备选 URL
    urls_to_try = [source_cfg["url"]] + source_cfg.get("alt_urls", [])

    for i, url in enumerate(urls_to_try):
        tag = "主地址" if i == 0 else f"备选{i}"
        print(f"  📡 正在抓取 {source_name}（{tag}：{url[:60]}...）")

        html, error = fetch_url(url)

        if error:
            print(f"    ⚠️ {tag} 失败: {error}")
            continue

        items = parse_html_items(html, source_name)

        if items:
            print(f"    ✅ {source_name} 抓取到 {len(items)} 条")
            return items[:max_items]

        print(f"    ⚠️ {tag} 未提取到有效内容，尝试下一个地址...")

    # 所有地址都失败
    print(f"    ❌ {source_name} 所有地址均失败")
    return [{
        "title": f"[待搜索] {source_name} 今日重点政策",
        "source": source_name,
        "url": "",
        "summary": f"自动抓取失败，请在 AI 分析阶段通过 WebSearch 搜索：{source_cfg['fallback_search']}",
        "needs_search": True,
        "search_query": source_cfg["fallback_search"],
    }]


def generate_brief_markdown(items, date_str=None):
    """
    将政策条目列表生成为 Markdown 简报格式。
    纯代码模板拼接，零 token 消耗。
    """
    if not date_str:
        date_str = datetime.now().strftime("%Y%m%d")

    lines = [
        f"# {date_str} 政策与宏观动态简报",
        f"\n> 生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"> 信源数量：{len(set(item['source'] for item in items))}",
        f"> 政策条目：{len(items)} 条",
        "",
    ]

    # 按来源分组
    from collections import defaultdict
    by_source = defaultdict(list)
    for item in items:
        by_source[item["source"]].append(item)

    idx = 1
    for source_name, source_items in by_source.items():
        lines.append(f"## {source_name}")
        lines.append("")

        for item in source_items:
            if item.get("needs_search"):
                lines.append(f"### {idx}. {item['title']}")
                lines.append(f"- **状态**：⚠️ 需要 AI 补充搜索")
                lines.append(f"- **建议搜索**：{item.get('search_query', '')}")
                lines.append("")
            else:
                lines.append(f"### {idx}. {item['title']}")
                if item.get("url"):
                    lines.append(f"- **链接**：{item['url']}")
                lines.append(f"- **来源**：{item['source']}")
                if item.get("summary"):
                    lines.append(f"- **摘要**：{item['summary']}")
                lines.append("")

            idx += 1

    lines.append("---")
    lines.append("")
    lines.append("*本简报由 policy_scraper.py 自动生成，0 token 消耗*")
    lines.append("")

    return "\n".join(lines)


def main():
    parser = argparse.ArgumentParser(description="政策信源采集脚本")
    parser.add_argument("--output", "-o", type=str, default="",
                        help="输出 JSON 文件路径（默认自动生成）")
    parser.add_argument("--limit", "-l", type=int, default=5,
                        help="每个信源最多抓取条数（默认 5）")
    parser.add_argument("--markdown", "-m", action="store_true",
                        help="同时生成 Markdown 简报")
    parser.add_argument("--sources", "-s", type=str, default="",
                        help="指定信源，逗号分隔（默认全部 7 个）")

    args = parser.parse_args()

    # 确定输出目录
    today = datetime.now().strftime("%y%m%d")
    base_dir = Path(__file__).parent.parent / "生成" / today
    base_dir.mkdir(parents=True, exist_ok=True)

    # 输出路径
    if args.output:
        json_path = Path(args.output)
    else:
        json_path = base_dir / "policy_items.json"

    # 确定要抓取的信源
    if args.sources:
        source_ids = [s.strip() for s in args.sources.split(",")]
        sources = [s for s in SOURCES if s["id"] in source_ids]
    else:
        sources = SOURCES

    print(f"\n{'='*60}")
    print(f"📡 政策信源采集开始")
    print(f"📅 日期：{datetime.now().strftime('%Y-%m-%d')}")
    print(f"📁 输出：{json_path}")
    print(f"📊 信源：{len(sources)} 个")
    print(f"{'='*60}\n")

    all_items = []

    for source_cfg in sources:
        items = scrape_source(source_cfg, max_items=args.limit)
        all_items.extend(items)

    # 去重（按标题相似度）
    seen = set()
    unique_items = []
    for item in all_items:
        title_key = item["title"][:20]
        if title_key not in seen:
            seen.add(title_key)
            unique_items.append(item)

    # 保存 JSON
    output_data = {
        "date": datetime.now().strftime("%Y-%m-%d"),
        "source_count": len(sources),
        "item_count": len(unique_items),
        "items": unique_items,
    }

    json_path.parent.mkdir(parents=True, exist_ok=True)
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(output_data, f, ensure_ascii=False, indent=2)

    print(f"\n✅ JSON 已保存：{json_path}")

    # 生成 Markdown 简报
    if args.markdown:
        md_path = base_dir / f"政策简报_{today}.md"
        md_content = generate_brief_markdown(unique_items)
        with open(md_path, "w", encoding="utf-8") as f:
            f.write(md_content)
        print(f"📝 Markdown 简报已保存：{md_path}")

    # 统计摘要
    needs_search = sum(1 for item in unique_items if item.get("needs_search"))
    direct_captured = len(unique_items) - needs_search

    print(f"\n{'='*60}")
    print(f"📊 采集结果摘要")
    print(f"   总条目：{len(unique_items)} 条")
    print(f"   直接抓取：{direct_captured} 条")
    print(f"   需 AI 补充：{needs_search} 条")
    print(f"{'='*60}\n")

    return output_data


if __name__ == "__main__":
    main()
