#!/usr/bin/env python3
"""
单文件PDF转换测试脚本

这个脚本允许你测试单个PDF文件的转换，而不会处理目录中的其他文件。
"""

import os
import sys
from pdf_converter import convert_pdf_to_docx, generate_docx_path, check_dependencies

def convert_single_file(pdf_filename):
    """
    转换单个PDF文件
    
    Args:
        pdf_filename: 要转换的PDF文件名
    """
    print(f"单文件PDF转换器")
    print("=" * 50)
    
    # 检查依赖
    if not check_dependencies():
        print("❌ 依赖检查失败，无法继续")
        return False
    
    # 检查文件是否存在
    if not os.path.exists(pdf_filename):
        print(f"❌ 错误：找不到文件 '{pdf_filename}'")
        return False
    
    # 生成输出文件路径
    docx_path = generate_docx_path(pdf_filename)
    docx_filename = os.path.basename(docx_path)
    
    print(f"📄 输入文件: {pdf_filename}")
    print(f"📄 输出文件: {docx_filename}")
    print()
    
    # 检查输出文件是否已存在
    if os.path.exists(docx_path):
        response = input(f"⚠️  文件 '{docx_filename}' 已存在。是否覆盖？(y/n): ").strip().lower()
        if response not in ['y', 'yes']:
            print("❌ 用户取消转换")
            return False
    
    try:
        print(f"🔄 开始转换: {pdf_filename}")
        
        # 执行转换
        success = convert_pdf_to_docx(pdf_filename, docx_path)
        
        if success:
            print(f"✅ 转换成功: {pdf_filename} → {docx_filename}")
            
            # 检查输出文件大小
            if os.path.exists(docx_path):
                file_size = os.path.getsize(docx_path)
                print(f"📊 输出文件大小: {file_size:,} 字节")
            
            return True
        else:
            print(f"❌ 转换失败: {pdf_filename}")
            return False
            
    except Exception as e:
        print(f"❌ 转换过程中出错: {e}")
        return False

def main():
    """主函数"""
    if len(sys.argv) != 2:
        print("使用方法:")
        print(f"  python {sys.argv[0]} <PDF文件名>")
        print()
        print("示例:")
        print(f"  python {sys.argv[0]} mytest.pdf")
        print(f"  python {sys.argv[0]} dikongjingji.pdf")
        return
    
    pdf_filename = sys.argv[1]
    success = convert_single_file(pdf_filename)
    
    print()
    print("=" * 50)
    if success:
        print("🎉 转换完成！")
    else:
        print("❌ 转换失败")

if __name__ == "__main__":
    main()