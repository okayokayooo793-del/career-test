"""
智能数据分类脚本
功能：根据自定义规则自动分类数据，标记异常，输出分类结果
场景：客户分级（ABCD）、订单状态分类、产品打标签、异常交易检测
用法：python classify_data.py <数据文件路径>
"""

import pandas as pd
import sys
import os
from pathlib import Path
from datetime import datetime


def auto_classify(file_path, rules=None):
    """
    智能分类 + 异常检测
    如果没有自定义规则，自动基于数据特征生成分类
    """
    file_path = Path(file_path)
    if not file_path.exists():
        print(f"❌ 文件不存在：{file_path}")
        return

    df = pd.read_excel(file_path)
    print(f"📊 读取 {len(df)} 行数据")

    # 自动识别数值列和分类依据
    num_cols = df.select_dtypes(include=['int64', 'float64']).columns.tolist()
    print(f"   数值列：{num_cols}")

    results = {}

    # ----- 自动分类 -----
    for col in num_cols:
        col_data = df[col].dropna()
        if len(col_data) < 4:
            continue

        q1 = col_data.quantile(0.25)
        q2 = col_data.quantile(0.50)
        q3 = col_data.quantile(0.75)
        iqr = q3 - q1

        def classify(val):
            if pd.isna(val):
                return '无数据'
            if val >= q3 + 1.5 * iqr:
                return '🔥 异常高'
            if val >= q3:
                return 'A级（高）'
            if val >= q2:
                return 'B级（中）'
            if val >= q1:
                return 'C级（一般）'
            if val >= q1 - 1.5 * iqr:
                return 'D级（低）'
            return '⚠️ 异常低'

        classify_col = f'{col}_等级'
        df[classify_col] = df[col].apply(classify)

        # 统计各等级数量
        dist = df[classify_col].value_counts()
        results[col] = {
            'q1': q1, 'q2': q2, 'q3': q3,
            'distribution': dist.to_dict()
        }
        print(f"\n🏷️ 【{col}】自动分级：")
        print(f"   阈值：Q1={q1:.1f}  Q2={q2:.1f}  Q3={q3:.1f}")
        for level, count in dist.items():
            print(f"   {level}：{count} 条 ({count/len(df)*100:.1f}%)")

    # ----- 异常检测 -----
    print(f"\n🚨 异常数据检测：")
    anomaly_found = False
    for col in num_cols:
        col_data = df[col].dropna()
        q1 = col_data.quantile(0.25)
        q3 = col_data.quantile(0.75)
        iqr = q3 - q1
        lower = q1 - 1.5 * iqr
        upper = q3 + 1.5 * iqr
        anomalies = df[(df[col] < lower) | (df[col] > upper)]
        if len(anomalies) > 0:
            anomaly_found = True
            print(f"   ⚠️ {col}：发现 {len(anomalies)} 条异常值（正常范围：{lower:.1f} ~ {upper:.1f}）")
            if len(anomalies) <= 5:
                for idx, row in anomalies.iterrows():
                    print(f"      第{idx+1}行：{row[col]}")

    if not anomaly_found:
        print("   ✅ 未发现明显异常数据")

    # ----- 保存结果 -----
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = file_path.parent / f"分类结果_{timestamp}.xlsx"

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='分类结果', index=False)
        # 统计汇总
        summary_rows = []
        for col, info in results.items():
            for level, count in info['distribution'].items():
                summary_rows.append({'字段': col, '等级': level, '数量': count})
        if summary_rows:
            summary_df = pd.DataFrame(summary_rows)
            summary_df.to_excel(writer, sheet_name='等级分布汇总', index=False)
        # 异常数据单独sheet
        anomaly_data = df[df.select_dtypes(include=['object']).columns.str.contains('异常', na=False).any(axis=1)]
        if len(anomaly_data) > 0:
            anomaly_data.to_excel(writer, sheet_name='异常数据', index=False)

    print(f"\n✅ 分类结果已保存：{output_path}")
    print(f"   包含 3 个 Sheet：分类结果 / 等级分布汇总 / 异常数据")
    return output_path


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("=" * 60)
        print("  🏷️ 智能数据分类脚本")
        print("=" * 60)
        print("\n用法：")
        print("  python classify_data.py <Excel文件路径>")
        print("\n示例：")
        print("  python classify_data.py customers.xlsx")
        print("\n自动完成：")
        print("  1. 四分位数自动分级（ABCD四级）")
        print("  2. IQR 异常检测（标记异常高/低）")
        print("  3. 等级分布统计汇总")
        print("  4. 输出分类结果 Excel（3个Sheet）")
        print("\n📩 定制需求：自定义分类规则、多条件组合分类、定期重新分类")
    else:
        auto_classify(sys.argv[1])
