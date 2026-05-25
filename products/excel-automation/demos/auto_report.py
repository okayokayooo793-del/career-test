"""
自动报表生成脚本
功能：从原始 Excel 数据自动生成统计汇总 + 图表 + 分析结论
场景：销售月报、运营周报、财务报表、KPI 看板
用法：python auto_report.py <数据文件路径>
"""

import pandas as pd
import sys
import os
from pathlib import Path
from datetime import datetime, timedelta
import warnings
warnings.filterwarnings('ignore')


def generate_report(file_path):
    """
    从 Excel 数据生成自动分析报告
    支持自动识别日期列、金额列、分类列
    """
    file_path = Path(file_path)
    if not file_path.exists():
        print(f"❌ 文件不存在：{file_path}")
        return

    # 1. 读取数据
    df = pd.read_excel(file_path)
    print(f"📊 读取 {len(df)} 行数据，{len(df.columns)} 列")
    print(f"   列名：{', '.join(df.columns.tolist())}")

    # 2. 自动识别列类型
    date_cols = []
    num_cols = []
    cat_cols = []

    for col in df.columns:
        if df[col].dtype == 'datetime64[ns]' or any(kw in str(col).lower() for kw in ['日期', '时间', 'date', 'time']):
            date_cols.append(col)
        elif df[col].dtype in ['int64', 'float64']:
            num_cols.append(col)
        else:
            # 尝试转日期
            try:
                pd.to_datetime(df[col])
                date_cols.append(col)
            except:
                if df[col].nunique() < len(df) * 0.3:
                    cat_cols.append(col)

    print(f"\n🔍 自动识别：")
    print(f"   日期列：{date_cols if date_cols else '未识别'}")
    print(f"   数值列：{num_cols if num_cols else '未识别'}")
    print(f"   分类列：{cat_cols if cat_cols else '未识别'}")

    # 3. 生成统计报告
    report_lines = []
    report_lines.append("=" * 60)
    report_lines.append("  📊 数据分析报告")
    report_lines.append(f"  生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M')}")
    report_lines.append(f"  数据来源：{file_path.name}")
    report_lines.append(f"  数据行数：{len(df)}")
    report_lines.append("=" * 60)

    # 4. 数值列统计
    if num_cols:
        report_lines.append("\n📈 数值统计：")
        stats = df[num_cols].describe()
        for col in num_cols:
            report_lines.append(f"  【{col}】")
            report_lines.append(f"    总和：{df[col].sum():,.2f}")
            report_lines.append(f"    均值：{df[col].mean():,.2f}")
            report_lines.append(f"    最大：{df[col].max():,.2f}")
            report_lines.append(f"    最小：{df[col].min():,.2f}")

    # 5. 分类统计
    if cat_cols:
        report_lines.append("\n📋 分类汇总：")
        for col in cat_cols[:3]:  # 最多分析前3个分类列
            report_lines.append(f"\n  【{col}】分布：")
            value_counts = df[col].value_counts().head(10)
            for cat, count in value_counts.items():
                pct = count / len(df) * 100
                report_lines.append(f"    {cat}：{count} ({pct:.1f}%)")

    # 6. 如果有日期列和数值列，做趋势分析
    if date_cols and num_cols:
        date_col = date_cols[0]
        try:
            df[date_col] = pd.to_datetime(df[date_col])
            df_sorted = df.sort_values(date_col)
            report_lines.append(f"\n📅 时间趋势（按 {date_col}）：")
            for nc in num_cols[:2]:
                first = df_sorted[nc].iloc[:3].mean() if len(df_sorted) >= 3 else df_sorted[nc].iloc[0]
                last = df_sorted[nc].iloc[-3:].mean() if len(df_sorted) >= 3 else df_sorted[nc].iloc[-1]
                change = ((last - first) / first * 100) if first != 0 else 0
                direction = "↑" if change > 0 else "↓" if change < 0 else "→"
                report_lines.append(f"  {nc}：初期均值 {first:,.2f} → 末期均值 {last:,.2f} {direction} {abs(change):.1f}%")
        except:
            pass

    # 7. 保存报告
    report_text = "\n".join(report_lines)
    output_dir = file_path.parent
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    # 保存文本报告
    txt_path = output_dir / f"分析报告_{timestamp}.txt"
    with open(txt_path, 'w', encoding='utf-8') as f:
        f.write(report_text)

    # 保存统计 Excel
    xlsx_path = output_dir / f"统计数据_{timestamp}.xlsx"
    with pd.ExcelWriter(xlsx_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='原始数据', index=False)
        if num_cols:
            df[num_cols].describe().to_excel(writer, sheet_name='数值统计')
        if cat_cols:
            for col in cat_cols[:3]:
                df[col].value_counts().head(20).to_excel(writer, sheet_name=f'{col[:28]}_分布')

    print(f"\n" + report_text)
    print(f"\n✅ 报告已保存：")
    print(f"   📄 {txt_path}")
    print(f"   📊 {xlsx_path}")

    return txt_path, xlsx_path


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("=" * 60)
        print("  📊 自动报表生成脚本")
        print("=" * 60)
        print("\n用法：")
        print("  python auto_report.py <Excel文件路径>")
        print("\n示例：")
        print("  python auto_report.py sales_2024.xlsx")
        print("\n自动识别：")
        print("  • 日期列 → 趋势分析（初期→末期变化）")
        print("  • 数值列 → 求和/均值/最大/最小")
        print("  • 分类列 → 分布占比 TOP10")
        print("\n输出：")
        print("  • 文本分析报告（.txt）")
        print("  • 统计汇总表格（.xlsx，含多个Sheet）")
        print("\n📩 定制需求：添加图表、Word报告、PPT输出、定时生成")
    else:
        generate_report(sys.argv[1])
