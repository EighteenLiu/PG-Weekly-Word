from __future__ import annotations

import argparse
from pathlib import Path

from .core import generate_weekly_report


def main() -> None:
    parser = argparse.ArgumentParser(description="生成平谷区垃圾分类周报")
    parser.add_argument("--district", required=True, help="区级基础台账 .xls")
    parser.add_argument("--city", required=True, help="周市级检查台账 .xlsx")
    parser.add_argument("--template", required=True, help="Jinja 周报模板 .docx")
    parser.add_argument("--group", default="", help="可选：街乡镇分组清单 .xlsx；留空则使用 docs 中的固定 JSON")
    parser.add_argument("--output", required=True, help="输出目录")
    parser.add_argument("--year", default="2026", help="报告年份")
    parser.add_argument("--district-period", default="", help="区级周期，如 07.04-07.10")
    parser.add_argument("--city-period", default="", help="市级周期，如 07.02-07.05")
    parser.add_argument("--sheet", default="", help="市级台账工作表名称")
    parser.add_argument("--previous-district-issues", type=int, default=None, help="上周区级垃圾分类问题总数")
    args = parser.parse_args()
    result = generate_weekly_report(
        district_base_path=Path(args.district),
        city_ledger_path=Path(args.city),
        template_path=Path(args.template),
        group_path=Path(args.group) if args.group else None,
        output_dir=Path(args.output),
        report_year=args.year,
        district_period=args.district_period,
        city_period=args.city_period,
        city_sheet=args.sheet or None,
        previous_district_issues=args.previous_district_issues,
    )
    print(result.report_path)
    print("问题台账和统计摘要已作为临时文件生成并清理。")
    for warning in result.context.get("warnings", []):
        print(f"配置提示：{warning}")
