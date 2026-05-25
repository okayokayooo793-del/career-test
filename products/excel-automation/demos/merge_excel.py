"""
多 Excel 文件合并脚本
功能：读取指定文件夹内所有 .xlsx/.xls 文件，自动合并到一个文件
场景：各部门日报/周报汇总、多门店销售数据合并、多来源数据整合
用法：python merge_excel.py <文件夹路径>
"""

import pandas as pd
import sys
import os
from pathlib import Path
from datetime import datetime

def merge_excel_files(folder_path, output_name=None):
    """
    合并文件夹内所有 Excel 文件
    参数：
        folder_path: 要合并的文件夹路径
        output_name: 输出文件名（可选，默认带时间戳）
    """
    folder = Path(folder_path)
    if not folder.exists():
        print(f"❌ 文件夹不存在：{folder_path}")
        return

    # 1. 扫描所有 Excel 文件
    excel_files = []
    for ext in ['*.xlsx', '*.xls', '*.xlsm']:
        excel_files.extend(folder.glob(ext))

    # 过滤掉临时文件（以 ~$ 开头）
    excel_files = [f for f in excel_files if not f.name.startswith('~$')]

    if not excel_files:
        print(f"❌ 在 {folder_path} 中未找到 Excel 文件")
        return

    print(f"📁 找到 {len(excel_files)} 个文件：")
    for f in excel_files:
        print(f"   • {f.name}")

    # 2. 逐个读取并合并
    all_data = []
    failed = []

    for f in excel_files:
        try:
            df = pd.read_excel(f)
            # 添加来源文件列，方便追溯
            df['来源文件'] = f.name
            all_data.append(df)
            print(f"   ✅ {f.name} → {len(df)} 行")
        except Exception as e:
            failed.append(f.name)
            print(f"   ⚠️ {f.name} 读取失败：{e}")

    if not all_data:
        print("❌ 没有成功读取任何文件")
        return

    # 3. 合并所有数据
    merged = pd.concat(all_data, ignore_index=True)

    # 4. 去除完全重复行
    before = len(merged)
    merged = merged.drop_duplicates()
    after = len(merged)
    print(f"\n📊 合并后共 {after} 行（去重 {before - after} 行）")

    # 5. 保存结果
    if output_name is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_name = f"合并结果_{timestamp}.xlsx"

    output_path = folder / output_name
    merged.to_excel(output_path, index=False, engine='openpyxl')

    print(f"\n✅ 合并完成！结果保存在：{output_path}")
    print(f"📏 总行数：{after} ｜ 总列数：{len(merged.columns)}")

    # 6. 数据概览
    print(f"\n📋 数据预览（前 5 行）：")
    print(merged.head().to_string())

    return output_path


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("=" * 60)
        print("  📑 多 Excel 文件合并脚本")
        print("=" * 60)
        print("\n用法：")
        print("  python merge_excel.py <文件夹路径>")
        print("\n示例：")
        print("  python merge_excel.py C:\\Users\\张三\\Desktop\\日报汇总")
        print("\n说明：")
        print("  1. 自动读取文件夹内所有 .xlsx/.xls 文件")
        print("  2. 自动对齐列名，列名相同的合并，不同的新增列")
        print("  3. 自动去重，添加「来源文件」列方便追溯")
        print("  4. 合并结果保存在同一文件夹内")
        print("\n📩 需要定制功能？联系我添加：")
        print("  - 指定合并哪些 Sheet")
        print("  - 按特定列去重")
        print("  - 自定义合并规则（按行/按列）")
        print("  - 输出格式指定（CSV/Excel/JSON）")
    else:
        merge_excel_files(sys.argv[1])
