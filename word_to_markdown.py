#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word转Markdown转换脚本

功能：
1. 自动扫描word目录中的所有.docx文件
2. 将Word文档转换为Markdown格式
3. 保持原始文件名（仅修改扩展名）
4. 输出到docs目录

依赖：
- pandoc：系统需安装pandoc命令行工具
- python3：Python 3.6+

安装pandoc：
- Ubuntu/Debian: apt-get install pandoc
- CentOS/RHEL: yum install pandoc
- macOS: brew install pandoc

使用方法：
    python word_to_markdown.py
    或
    python word_to_markdown.py --input docs/word --output docs
"""

import os
import sys
import subprocess
import argparse
from pathlib import Path
from typing import List, Tuple


class Word2Markdown:
    """Word到Markdown转换器"""

    def __init__(self, input_dir: str, output_dir: str):
        """
        初始化转换器

        Args:
            input_dir: Word文档所在目录
            output_dir: Markdown输出目录
        """
        self.input_dir = Path(input_dir)
        self.output_dir = Path(output_dir)

        # 确保输出目录存在
        self.output_dir.mkdir(parents=True, exist_ok=True)

    def check_pandoc(self) -> bool:
        """
        检查系统是否安装了pandoc

        Returns:
            bool: True表示已安装，False表示未安装
        """
        try:
            result = subprocess.run(
                ['pandoc', '--version'],
                capture_output=True,
                text=True,
                timeout=5
            )
            if result.returncode == 0:
                print(f"✓ 找到pandoc: {result.stdout.split()[1]}")
                return True
        except (FileNotFoundError, subprocess.TimeoutExpired):
            pass

        return False

    def find_word_files(self) -> List[Path]:
        """
        查找输入目录中的所有Word文档

        Returns:
            List[Path]: Word文档路径列表
        """
        if not self.input_dir.exists():
            print(f"✗ 输入目录不存在: {self.input_dir}")
            return []

        word_files = list(self.input_dir.glob("*.docx"))

        if not word_files:
            print(f"✗ 在 {self.input_dir} 中未找到.docx文件")
        else:
            print(f"✓ 找到 {len(word_files)} 个Word文档")

        return word_files

    def convert_file(self, word_file: Path) -> Tuple[bool, str]:
        """
        转换单个Word文件到Markdown

        Args:
            word_file: Word文件路径

        Returns:
            Tuple[bool, str]: (是否成功, 消息)
        """
        # 生成输出文件名
        markdown_file = self.output_dir / f"{word_file.stem}.md"

        # 如果Markdown文件已存在，询问是否覆盖
        if markdown_file.exists():
            print(f"  ⚠ 警告: {markdown_file.name} 已存在，将被覆盖")

        # 构建pandoc命令
        # pandoc参数说明:
        # -f docx: 输入格式为Word
        # -t markdown: 输出格式为Markdown
        # -o: 输出文件
        # --extract-media=./images: 提取图片到images目录
        # --wrap=none: 不自动换行
        # --toc: 生成目录
        # --toc-depth=3: 目录深度为3级
        # 使用绝对路径，避免pandoc执行时的路径问题
        cmd = [
            'pandoc',
            '-f', 'docx',
            '-t', 'markdown',
            '-o', str(markdown_file.absolute()),
            '--extract-media=./images',
            '--wrap=none',
            '--toc',
            '--toc-depth=3',
            str(word_file.absolute())
        ]

        try:
            print(f"  🔄 转换中: {word_file.name} -> {markdown_file.name}")
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=60,
                cwd=str(self.output_dir)
            )

            if result.returncode == 0:
                # 检查输出文件大小
                if markdown_file.exists() and markdown_file.stat().st_size > 0:
                    size = markdown_file.stat().st_size
                    print(f"  ✓ 成功: {markdown_file.name} ({size:,} bytes)")
                    return True, f"转换成功: {markdown_file.name}"
                else:
                    return False, f"输出文件为空: {markdown_file.name}"
            else:
                error_msg = result.stderr.strip() if result.stderr else "未知错误"
                print(f"  ✗ 失败: {error_msg}")
                return False, f"转换失败: {error_msg}"

        except subprocess.TimeoutExpired:
            return False, "转换超时（60秒）"
        except Exception as e:
            return False, f"转换异常: {str(e)}"

    def convert_all(self) -> Tuple[int, int, List[str]]:
        """
        转换所有Word文件

        Returns:
            Tuple[int, int, List[str]]: (成功数, 失败数, 错误消息列表)
        """
        word_files = self.find_word_files()

        if not word_files:
            return 0, 0, []

        success_count = 0
        fail_count = 0
        errors = []

        print(f"\n{'='*60}")
        print(f"开始转换 {len(word_files)} 个Word文档")
        print(f"输入目录: {self.input_dir}")
        print(f"输出目录: {self.output_dir}")
        print(f"{'='*60}\n")

        for i, word_file in enumerate(word_files, 1):
            print(f"[{i}/{len(word_files)}] {word_file.name}")

            success, msg = self.convert_file(word_file)

            if success:
                success_count += 1
            else:
                fail_count += 1
                errors.append(f"{word_file.name}: {msg}")

            print()

        return success_count, fail_count, errors

    def print_summary(self, success_count: int, fail_count: int, errors: List[str]):
        """
        打印转换总结

        Args:
            success_count: 成功数量
            fail_count: 失败数量
            errors: 错误消息列表
        """
        print(f"{'='*60}")
        print(f"转换完成！")
        print(f"{'='*60}")
        print(f"成功: {success_count}")
        print(f"失败: {fail_count}")
        print(f"总计: {success_count + fail_count}")

        if errors:
            print(f"\n错误详情:")
            for error in errors:
                print(f"  ✗ {error}")

        print(f"\n输出目录: {self.output_dir.absolute()}")
        print(f"{'='*60}")


def main():
    """主函数"""
    parser = argparse.ArgumentParser(
        description='Word转Markdown转换脚本',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例用法:
  # 使用默认目录（docs/word -> docs）
  python word_to_markdown.py

  # 指定输入输出目录
  python word_to_markdown.py --input /path/to/word --output /path/to/markdown

  # 仅检查不转换
  python word_to_markdown.py --check
        """
    )

    parser.add_argument(
        '--input',
        default='docs/word',
        help='Word文档所在目录（默认: docs/word）'
    )

    parser.add_argument(
        '--output',
        default='docs',
        help='Markdown输出目录（默认: docs）'
    )

    parser.add_argument(
        '--check',
        action='store_true',
        help='仅检查环境和文件，不执行转换'
    )

    args = parser.parse_args()

    # 创建转换器
    converter = Word2Markdown(args.input, args.output)

    # 检查pandoc
    print(f"{'='*60}")
    print(f"Word转Markdown转换器")
    print(f"{'='*60}")

    if not converter.check_pandoc():
        print("✗ 错误: 未找到pandoc命令")
        print("\n请先安装pandoc:")
        print("  Ubuntu/Debian: sudo apt-get install pandoc")
        print("  CentOS/RHEL:   sudo yum install pandoc")
        print("  macOS:         brew install pandoc")
        print("  或访问: https://pandoc.org/installing.html")
        sys.exit(1)

    # 如果是仅检查模式
    if args.check:
        word_files = converter.find_word_files()
        if word_files:
            print(f"\n将转换以下文件:")
            for f in word_files:
                print(f"  - {f.name}")
        sys.exit(0)

    # 执行转换
    success_count, fail_count, errors = converter.convert_all()

    # 打印总结
    converter.print_summary(success_count, fail_count, errors)

    # 根据结果设置退出码
    sys.exit(0 if fail_count == 0 else 1)


if __name__ == '__main__':
    main()
