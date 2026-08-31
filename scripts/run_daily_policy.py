#!/usr/bin/env python3
"""
政策全流程一键运行脚本
串联：信源采集 → 简报生成 → 分析报告 → 思维导图 → PPT

用法：
  python run_daily_policy.py              # 全流程
  python run_daily_policy.py --scrape-only  # 只采集
  python run_daily_policy.py --skip-scrape  # 跳过采集（用已有数据）
  python run_daily_policy.py --sources ndrc,moa  # 只采集指定信源

所有输出保存到：Z:\工作\CC\Official-Document-Drafting-Agent\生成\日期文件夹\
"""

import json
import sys
import os
import argparse
from datetime import datetime
from pathlib import Path

# 项目根目录
BASE_DIR = Path(__file__).resolve().parent.parent
SCRIPTS_DIR = BASE_DIR / "scripts"


def get_output_dir():
    """获取当日输出目录"""
    now = datetime.now()
    folder = f"{now.day:02d}{now.month:02d}{now.year}"
    out_dir = BASE_DIR / "生成" / folder
    out_dir.mkdir(parents=True, exist_ok=True)
    return out_dir


def run_step(step_name, script_path, args_str=""):
    """运行一个 Python 脚本步骤"""
    cmd = f'python "{script_path}" {args_str}'
    print(f"\n{'─'*60}")
    print(f"  ▶ {step_name}")
    print(f"  命令：{cmd}")
    print(f"{'─'*60}")

    ret = os.system(cmd)
    if ret != 0:
        print(f"  ⚠️ {step_name} 执行失败（返回码：{ret}）")
        return False
    print(f"  ✅ {step_name} 完成")
    return True


def main():
    parser = argparse.ArgumentParser(description="政策全流程一键运行")
    parser.add_argument("--scrape-only", action="store_true",
                        help="只执行信源采集，不做后续分析")
    parser.add_argument("--skip-scrape", action="store_true",
                        help="跳过信源采集（使用已有 policy_items.json）")
    parser.add_argument("--sources", type=str, default="",
                        help="指定信源，逗号分隔")
    parser.add_argument("--limit", type=int, default=5,
                        help="每站最多抓取条数")
    parser.add_argument("--json-input", type=str, default="",
                        help="直接输入分析JSON（跳过AI，直接生成文件）")

    args = parser.parse_args()
    output_dir = get_output_dir()

    print(f"\n{'='*60}")
    print(f"📡 政策全流程一键运行")
    print(f"📅 日期：{datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print(f"📁 输出目录：{output_dir}")
    print(f"{'='*60}")

    # ============================================================
    # Step 1: 信源采集
    # ============================================================
    items_file = output_dir / "policy_items.json"

    if not args.skip_scrape and not args.json_input:
        scrape_args = f'--limit {args.limit} --markdown'
        if args.sources:
            scrape_args += f' --sources {args.sources}'

        success = run_step("Step 1/5: 信源采集",
                          SCRIPTS_DIR / "policy_scraper.py",
                          scrape_args)
        if not success:
            print("\n⚠️ 信源采集失败，部分信源将由 AI 补充搜索")

    elif args.json_input:
        # 直接使用输入的 JSON
        with open(args.json_input, "r", encoding="utf-8") as f:
            data = json.load(f)
        print(f"\n✅ 使用已有分析数据：{args.json_input}")
        items_file = Path(args.json_input)

    # ============================================================
    # Step 2: AI 深度分析（需要 Claude 参与，不在本脚本中执行）
    # ============================================================
    if args.scrape_only:
        print("\n✅ 信源采集完成，跳过后续步骤（--scrape-only）")
        print(f"\n📁 采集结果已保存至：{output_dir}")
        return

    # 检查是否有 AI 分析结果
    analysis_file = output_dir / "analysis_result.json"

    if not analysis_file.exists() and not args.json_input:
        print(f"\n{'='*60}")
        print("🔍 Step 2/5: 需要 AI 进行深度分析")
        print(f"{'='*60}")
        print(f"\n⚠️ 未找到分析结果文件：{analysis_file}")
        print("\n请通过以下方式之一完成分析：")
        print("  1. 在 Claude 中运行 daily-policy-lite 工作流")
        print("  2. 手动对 policy_items.json 中的政策进行深度分析")
        print(f"  3. 将分析结果保存为 JSON，然后重新运行本脚本：")
        print(f"     python run_daily_policy.py --json-input analysis_result.json")
        print(f"\n💡 或者直接使用 --json-input 参数输入分析结果")
        return

    if args.json_input:
        analysis_file = Path(args.json_input)

    # ============================================================
    # Step 3: 生成 Word 报告
    # ============================================================
    report_file = output_dir / "政策深度分析报告.docx"
    run_step("Step 3/5: 生成 Word 深度分析报告",
            SCRIPTS_DIR / "report_generator.py",
            f'-c "{analysis_file}" -o "{report_file}"')

    # ============================================================
    # Step 4: 生成思维导图
    # ============================================================
    mindmap_file = output_dir / "思维导图.md"
    run_step("Step 4/5: 生成思维导图大纲",
            SCRIPTS_DIR / "mindmap_generator.py",
            f'-c "{analysis_file}" -o "{mindmap_file}"')

    # ============================================================
    # Step 5: 生成 PPT
    # ============================================================
    ppt_file = output_dir / "政策解读.pptx"
    run_step("Step 5/5: 生成 PPT",
            SCRIPTS_DIR / "ppt_generator.py",
            f'-c "{analysis_file}" -o "{ppt_file}"')

    # ============================================================
    # 总结
    # ============================================================
    print(f"\n{'='*60}")
    print("✅ 全流程完成！")
    print(f"{'='*60}")
    print(f"\n📁 输出目录：{output_dir}")
    print(f"\n📄 已生成文件：")
    for f in output_dir.iterdir():
        if f.is_file():
            size = f.stat().st_size / 1024
            print(f"   {f.name} ({size:.1f} KB)")
    print(f"\n{'='*60}")


if __name__ == "__main__":
    main()
