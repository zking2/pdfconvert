#!/usr/bin/env python3
"""
PDF to DOCX Single File Converter

A command-line utility to convert a specific PDF file to DOCX format.
Usage: python pdf_to_docx_single.py <pdf_file_path>
"""

import os
import sys
from pdf2docx import Converter


def check_dependencies():
    """检查pdf2docx依赖是否安装"""
    try:
        import pdf2docx
        return True
    except ImportError:
        print("❌ 错误: 未安装pdf2docx库")
        print("   请使用以下命令安装: pip install pdf2docx")
        return False


def convert_single_pdf(pdf_path, output_path=None):
    """
    转换单个PDF文件到DOCX格式
    
    Args:
        pdf_path: PDF文件路径
        output_path: 输出DOCX文件路径（可选，默认为同名.docx文件）
    """
    # 检查PDF文件是否存在
    if not os.path.exists(pdf_path):
        print(f"❌ 错误: PDF文件不存在: {pdf_path}")
        return False
    
    # 如果没有指定输出路径，生成默认路径
    if output_path is None:
        base_name = os.path.splitext(pdf_path)[0]
        output_path = f"{base_name}.docx"
    
    # 检查输出文件是否已存在
    if os.path.exists(output_path):
        response = input(f"⚠️  文件 '{output_path}' 已存在，是否覆盖? (y/n): ").strip().lower()
        if response not in ['y', 'yes', '是']:
            print("❌ 转换已取消")
            return False
    
    print(f"🔄 开始转换: {os.path.basename(pdf_path)}")
    
    try:
        # 执行转换
        cv = Converter(pdf_path)
        cv.convert(output_path)
        cv.close()
        
        print(f"✅ 转换成功: {os.path.basename(output_path)}")
        print(f"   输出文件: {output_path}")
        return True
        
    except Exception as e:
        print(f"❌ 转换失败: {e}")
        # 清理可能的部分输出文件
        if os.path.exists(output_path):
            try:
                os.remove(output_path)
            except:
                pass
        return False


def main():
    """主函数"""
    print("PDF to DOCX 单文件转换器")
    print("=" * 40)
    
    # 检查依赖
    if not check_dependencies():
        return
    
    # 检查命令行参数
    if len(sys.argv) < 2:
        print("使用方法:")
        print(f"  python {os.path.basename(__file__)} <PDF文件路径> [输出文件路径]")
        print()
        print("示例:")
        print(f"  python {os.path.basename(__file__)} mytest.pdf")
        print(f"  python {os.path.basename(__file__)} mytest.pdf output.docx")
        return
    
    pdf_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None
    
    # 执行转换
    success = convert_single_pdf(pdf_file, output_file)
    
    if success:
        print("\n🎉 转换完成!")
    else:
        print("\n❌ 转换失败!")
    
    print("=" * 40)


if __name__ == "__main__":
    main()