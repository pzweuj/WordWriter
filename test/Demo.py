# coding=utf-8
# pzw
# 20251022
# WordWriter v4.0.0 Demo - 使用新的面向对象 API

import sys
sys.path.insert(0, '..')  # 添加父目录到路径

from WordWriter import WordWriter

# 准备替换数据
replace_dict = {
    "#[testheader1]#": "测试页眉1",
    "#[testheader2]#": "页眉测试2",
    "#[testString]#": "，文本替换成功",
    "#[testfooter]#": "测试页脚",
    "#[TX-testString2]#": "，文本框文本替换成功",
    "#[testTableString1]#": "单元格文本替换成功",
    "#[testTableString2]#": "单元格文本替换成功",
    "#[IMAGE-test1-(30,30)]#": "testPicture.png",
    "#[IMAGE-test2]#": "testPicture2.png",
    "#[IMAGE-test3-(10,10)]#": "testPicture.png",
    "#[TABLE-test1]#": "testTable.txt"
}

print("=" * 60)
print("WordWriter v4.0.0 Demo - 面向对象 API")
print("=" * 60)

# ============================================================================
# 方式1: 链式调用（推荐，最简洁）
# ============================================================================
print("\n方式1: 链式调用")
print("-" * 60)

WordWriter("test.docx") \
    .replace(replace_dict) \
    .save("output_v4_chain.docx")

print("✓ 完成！文件已保存为: output_v4_chain.docx")

# ============================================================================
# 方式2: 分步调用（更多控制）
# ============================================================================
print("\n方式2: 分步调用")
print("-" * 60)

# 创建 WordWriter 实例
writer = WordWriter("test.docx")
print(f"1. 创建实例: {writer}")

# 加载模板
writer.load()
print(f"2. 加载模板: {writer}")

# 查看找到的标签（v4.0 新功能）
tags = writer.get_tags()
print(f"3. 找到 {len(tags)} 个标签:")
for tag in sorted(tags):
    print(f"   - {tag}")

# 替换标签
print("4. 替换标签...")
writer.replace(replace_dict, logs=False)

# 保存文档
writer.save("output_v4_step.docx")
print("✓ 完成！文件已保存为: output_v4_step.docx")

# ============================================================================
# 方式3: 类方法（类似旧 API）
# ============================================================================
print("\n方式3: 类方法（一步完成）")
print("-" * 60)

WordWriter.process("test.docx", "output_v4_process.docx", replace_dict, logs=False)
print("✓ 完成！文件已保存为: output_v4_process.docx")

# ============================================================================
# 方式4: 上下文管理器（自动资源管理）
# ============================================================================
print("\n方式4: 上下文管理器")
print("-" * 60)

with WordWriter("test.docx") as writer:
    writer.replace(replace_dict, logs=False)
    writer.save("output_v4_context.docx")

print("✓ 完成！文件已保存为: output_v4_context.docx")

# ============================================================================
# 总结
# ============================================================================
print("\n" + "=" * 60)
print("所有方式都已完成！")
print("=" * 60)
print("\n推荐使用方式1（链式调用）获得最简洁的代码。")
print("如需更多控制，使用方式2（分步调用）。")
print("\n注意：旧的函数式 API 仍然可用，保持向后兼容。")
print("=" * 60)
