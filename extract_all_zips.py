#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
extract_all_zips.py
--------------------
递归扫描指定目录及其所有子文件夹，找到所有 .zip 文件并逐一解压。

特别处理:
1. 中文路径 / 中文文件名（Windows 上打包的 zip 内部文件名通常是 GBK 编码，
   如果直接用 Python 默认的 zipfile 读取会出现乱码，本脚本会自动修正）。
2. 每个 zip 会被解压到与其同目录、且同名（去掉 .zip 后缀）的新文件夹里，
   避免不同压缩包互相覆盖文件。
3. 可选：如果解压出来的内容里还包含 zip 文件（压缩包套压缩包），
   加 --nested 参数可以继续递归解压。
4. 加密 / 损坏的 zip 会被跳过并打印警告，不会中断整个流程。
5. 默认会跳过"已经解压过"的 zip（即目标文件夹已存在且非空），
   避免重复解压；可用 --force 强制重新解压。

用法示例（Windows PowerShell / CMD）:

    python extract_all_zips.py "Z:\\Xin\\meridian\\财联社\\机构报名表-0820\\报名表"

不加路径参数时，默认使用当前脚本所在目录下的 "当前工作目录" (os.getcwd())。

可选参数:
    --nested        解压后如果发现里面还有 zip，继续递归解压
    --force         即使目标文件夹已存在也强制重新解压（会覆盖同名文件）
    --dry-run       只打印将要执行的操作，不实际解压（用于预检查）
"""

import argparse
import os
import sys
import zipfile
from pathlib import Path


def fix_filename_encoding(raw_name: str) -> str:
    """
    修正 zip 内部文件名的乱码问题。

    背景: zip 格式的文件名编码没有强制标准。Windows 上很多压缩软件
    （老版本 WinRAR / 资源管理器自带压缩）会用本地编码（简体中文一般是 GBK/CP936）
    写入文件名，但没有设置 UTF-8 标志位。Python 的 zipfile 模块在这种情况下
    会默认按 CP437（西欧字符集）解码，导致中文文件名变成乱码。

    这里的做法: 把 Python 已经"错误地"按 CP437 解码出来的字符串，
    重新编码回原始字节，再尝试用 GBK 解码还原成正确的中文。
    """
    try:
        # zipfile 内部默认按 cp437 解码，先转换回原始字节
        raw_bytes = raw_name.encode('cp437')
    except UnicodeEncodeError:
        # 已经是正常的 unicode（比如设置了 UTF-8 标志位），直接返回
        return raw_name

    # 依次尝试常见的中文编码
    for encoding in ('gbk', 'gb18030', 'utf-8'):
        try:
            fixed = raw_bytes.decode(encoding)
            return fixed
        except UnicodeDecodeError:
            continue

    # 都失败的话，保底返回原始名字，避免脚本崩溃
    return raw_name


def extract_one_zip(zip_path: Path, dest_dir: Path, dry_run: bool = False) -> bool:
    """
    解压单个 zip 文件到 dest_dir，正确处理中文文件名。
    返回 True 表示成功，False 表示失败/跳过。
    """
    try:
        with zipfile.ZipFile(zip_path, 'r') as zf:
            infolist = zf.infolist()

            if dry_run:
                print(f"  [DRY-RUN] 将解压 {len(infolist)} 个条目 -> {dest_dir}")
                return True

            dest_dir.mkdir(parents=True, exist_ok=True)

            for info in infolist:
                # 判断该条目是否已经声明为 UTF-8（bit 11 / 0x800）
                is_utf8_flagged = bool(info.flag_bits & 0x800)
                name = info.filename if is_utf8_flagged else fix_filename_encoding(info.filename)

                # 统一路径分隔符，防止 Windows/Unix 混用问题
                name = name.replace('\\', '/')
                target_path = dest_dir / name

                if info.is_dir():
                    target_path.mkdir(parents=True, exist_ok=True)
                    continue

                target_path.parent.mkdir(parents=True, exist_ok=True)

                with zf.open(info, 'r') as source, open(target_path, 'wb') as target:
                    target.write(source.read())

        return True

    except zipfile.BadZipFile:
        print(f"  [错误] 不是有效的 zip 文件，已跳过: {zip_path}")
        return False
    except RuntimeError as e:
        # 常见于加密压缩包（需要密码）
        print(f"  [错误] 解压失败（可能需要密码）: {zip_path}  原因: {e}")
        return False
    except Exception as e:
        print(f"  [错误] 解压时出现未知问题: {zip_path}  原因: {e}")
        return False


def find_all_zips(root: Path):
    """递归查找 root 目录下所有 .zip 文件（不区分大小写）。"""
    for dirpath, dirnames, filenames in os.walk(root):
        for fname in filenames:
            if fname.lower().endswith('.zip'):
                yield Path(dirpath) / fname


def process_zip(zip_path: Path, force: bool, nested: bool, dry_run: bool, depth: int = 0):
    indent = '  ' * depth
    dest_dir = zip_path.parent / zip_path.stem  # 同目录下，与 zip 同名的文件夹

    if dest_dir.exists() and any(dest_dir.iterdir()) and not force:
        print(f"{indent}[跳过] 已存在解压结果，跳过: {zip_path.name}")
        return

    print(f"{indent}[解压] {zip_path}")
    ok = extract_one_zip(zip_path, dest_dir, dry_run=dry_run)

    if ok and not dry_run:
        print(f"{indent}  -> 完成: {dest_dir}")

    # 如果开启了 --nested，检查解压出来的内容里是否还有 zip
    if ok and nested and not dry_run:
        for inner_zip in find_all_zips(dest_dir):
            process_zip(inner_zip, force=force, nested=nested, dry_run=dry_run, depth=depth + 1)


def main():
    parser = argparse.ArgumentParser(description="递归解压目录下所有 zip 文件（含中文文件名修复）")
    parser.add_argument(
        'root',
        nargs='?',
        default=os.getcwd(),
        help='要扫描的根目录（默认当前目录）'
    )
    parser.add_argument('--nested', action='store_true', help='解压出的内容如果还有 zip，继续递归解压')
    parser.add_argument('--force', action='store_true', help='即使已有解压结果也强制重新解压')
    parser.add_argument('--dry-run', action='store_true', help='只打印将要执行的操作，不实际解压')
    args = parser.parse_args()

    #root = Path(args.root)
    root = Path("Z:\Xin\meridian\财联社\机构报名表-0820\报名表")
    if not root.exists():
        print(f"错误: 目录不存在 -> {root}")
        sys.exit(1)

    print(f"扫描目录: {root}")
    zips = list(find_all_zips(root))
    print(f"共找到 {len(zips)} 个 zip 文件\n")

    if not zips:
        print("没有找到任何 zip 文件。")
        return

    for i, zp in enumerate(zips, 1):
        print(f"[{i}/{len(zips)}]")
        process_zip(zp, force=args.force, nested=args.nested, dry_run=args.dry_run)
        print()

    print("全部处理完成。")


if __name__ == '__main__':
    main()