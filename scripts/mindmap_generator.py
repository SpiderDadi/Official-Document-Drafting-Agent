#!/usr/bin/env python3
"""
思维导图大纲生成器
从政策分析结果生成思维导图结构（Markdown 格式）。

用法：
  python mindmap_generator.py --config analysis.json --output 思维导图.md

输出格式：
  - 标准 Markdown 层级大纲（可直接导入 XMind/幕布/FreeMind）
  - 中心主题 → 核心要点 → 细分内容

设计原则：
  - 纯代码实现，0 token 消耗
  - 从已有分析结果提取结构，无需 AI 重新理解
"""

import json
import argparse
import sys
from datetime import datetime
from pathlib import Path


def generate_mindmap(policy, center_theme=None):
    """
    从单个政策分析结果生成思维导图大纲。

    输入 policy 格式：
    {
        "title": "政策标题",
        "source": "信源",
        "summary": "摘要",
        "analysis": {
            "summary": "分析总结",
            "analysis": [
                {"round": "第一轮...", "question": "...", "answer": "..."}
            ],
            "actionItems": ["...", "..."]
        }
    }

    输出：Markdown 层级大纲文本
    """
    if not center_theme:
        center_theme = policy.get("title", "政策解析")

    lines = [
        f"# {center_theme}",
        "",
    ]

    # 第一层：基本信息
    lines.append("## 基本信息")
    lines.append(f"- 信源：{policy.get('source', '未知')}")
    lines.append(f"- 评级：{policy.get('rating', '未评级')}")
    if policy.get("summary"):
        summary = policy["summary"][:80] + ("..." if len(policy["summary"]) > 80 else "")
        lines.append(f"- 摘要：{summary}")
    lines.append("")

    # 第二层：核心要点
    analysis = policy.get("analysis", {})
    if isinstance(analysis, dict):
        if analysis.get("summary"):
            lines.append("## 分析总结")
            summary_text = analysis["summary"]
            # 按句号/分号拆分要点
            points = [p.strip() for p in summary_text.replace("。", "。\n").replace("；", "；\n").split("\n") if p.strip()]
            for p in points[:5]:  # 最多5个要点
                lines.append(f"- {p}")
            lines.append("")

        # 第三层：三轮分析
        rounds_data = analysis.get("analysis", [])
        if rounds_data:
            # 按轮次分组
            current_round = None
            round_items = []
            for item in rounds_data:
                r = item.get("round", "")
                if r != current_round:
                    if current_round and round_items:
                        _render_round_mindmap(lines, current_round, round_items)
                    current_round = r
                    round_items = [item]
                else:
                    round_items.append(item)
            if current_round and round_items:
                _render_round_mindmap(lines, current_round, round_items)

        # 第四层：行动建议
        action_items = analysis.get("actionItems", [])
        if action_items:
            lines.append("## 行动建议")
            for item in action_items:
                lines.append(f"- {item}")
            lines.append("")

    return "\n".join(lines)


def _render_round_mindmap(lines, round_name, qa_items):
    """将一轮分析渲染为思维导图节点"""
    # 简化轮次名称
    short_name = round_name.split("：")[0] if "：" in round_name else round_name

    lines.append(f"## {short_name}")
    for item in qa_items:
        q = item.get("question", "")
        a = item.get("answer", "")
        if q and a:
            # 答案取前60字
            a_short = a[:60] + ("..." if len(a) > 60 else "")
            lines.append(f"- **{q[:30]}**：{a_short}")
        elif q:
            lines.append(f"- {q}")
        elif a:
            lines.append(f"- {a[:80]}")
    lines.append("")


def main():
    parser = argparse.ArgumentParser(description="思维导图大纲生成器")
    parser.add_argument('--config', '-c', type=str,
                        help='JSON 配置文件路径')
    parser.add_argument('--json', '-j', type=str,
                        help='JSON 配置字符串')
    parser.add_argument('--output', '-o', type=str, required=True,
                        help='输出思维导图文件路径')

    args = parser.parse_args()

    if args.config:
        config = json.loads(Path(args.config).read_text(encoding="utf-8"))
    elif args.json:
        config = json.loads(args.json)
    else:
        print("错误：请提供 --config 或 --json 参数", file=sys.stderr)
        sys.exit(1)

    policies = config.get("policies", config.get("analyses", []))
    if not policies and "title" in config:
        policies = [config]

    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    all_content = []
    for policy in policies:
        title = policy.get("title", "政策解析")
        short_title = title[:20].replace("/", "_").replace("\\", "_")
        content = generate_mindmap(policy)
        all_content.append(content)

    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n\n---\n\n".join(all_content))

    print(f"✅ 思维导图大纲已保存：{output_path}")
    print(f"   共 {len(policies)} 个政策，可直接导入 XMind/幕布")


if __name__ == "__main__":
    main()
