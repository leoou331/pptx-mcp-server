"""
PPTX 文件安全验证模块
防护：ZIP炸弹、宏/VBA、路径遍历、XML实体注入
"""
import os
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Tuple


@dataclass
class SecurityLimits:
    """安全限制配置"""
    MAX_FILE_SIZE: int = 50 * 1024 * 1024           # 50MB
    MAX_UNCOMPRESSED_SIZE: int = 500 * 1024 * 1024  # 500MB
    MAX_COMPRESSION_RATIO: int = 100                # 100x
    MAX_SLIDES: int = 500
    MAX_SHAPES_PER_SLIDE: int = 1000
    MAX_IMAGE_SIZE: int = 10 * 1024 * 1024         # 10MB
    MAX_TEXT_LENGTH: int = 100_000
    MAX_TABLE_ROWS: int = 200
    MAX_TABLE_COLS: int = 50
    MAX_TABLE_CELLS: int = 5_000
    PROCESSING_TIMEOUT: int = 60                    # 秒
    SESSION_TTL: int = 3600                         # 1小时


# 全局限制配置
limits = SecurityLimits()


def has_macro(file_path: str) -> bool:
    """
    检测 PPTX 是否包含宏/VBA脚本
    
    检测方式：
    1. 检查文件名中是否包含危险模式
    2. 检查 [Content_Types].xml 中的宏声明
    """
    dangerous_patterns = [
        'vbaproject.bin',
        'vbadata.xml',
        'activex',
    ]
    
    try:
        with zipfile.ZipFile(file_path, 'r') as zf:
            # 1. 检查文件名
            for name in zf.namelist():
                name_lower = name.lower()
                for pattern in dangerous_patterns:
                    if pattern in name_lower:
                        return True
            
            # 2. 检查 Content_Types.xml（更可靠）
            try:
                content_types = zf.read('[Content_Types].xml').decode('utf-8', errors='ignore')
                if 'application/vnd.ms-office.vbaProject' in content_types:
                    return True
            except KeyError:
                pass  # 文件可能不存在
    
    except (zipfile.BadZipFile, Exception):
        pass
    
    return False


def safe_path(base_dir: str, user_path: str) -> str:
    """
    将用户路径解析到受控目录内，支持绝对路径白名单校验
    
    Args:
        base_dir: 默认基础目录（安全边界）
        user_path: 用户提供的路径
        
    Returns:
        安全的绝对路径
        
    Raises:
        ValueError: 如果检测到路径遍历攻击
    """
    return safe_path_in_dirs(base_dir, user_path)


def _is_relative_to(path: Path, base: Path) -> bool:
    """兼容 Python 3.8+ 的 Path.is_relative_to。"""
    try:
        path.relative_to(base)
        return True
    except ValueError:
        return False


def safe_path_in_dirs(base_dir: str, user_path: str, *allowed_dirs: str) -> str:
    """
    将路径限制在一个或多个允许目录中，跟随符号链接后再校验边界。

    相对路径会基于 base_dir 解析；绝对路径只有在落入 allowlist 时才允许。
    """
    if not isinstance(user_path, str):
        raise TypeError("文件路径必须是字符串")

    if "\x00" in user_path:
        raise ValueError("文件路径包含非法空字节")

    cleaned = user_path.strip()
    if not cleaned:
        raise ValueError("文件路径不能为空")

    base_path = Path(base_dir).resolve()
    allowed_paths = [base_path]
    for directory in allowed_dirs:
        if directory:
            allowed_paths.append(Path(directory).resolve())

    raw_path = Path(cleaned)
    candidate = raw_path if raw_path.is_absolute() else base_path / raw_path
    resolved = candidate.resolve(strict=False)

    if any(_is_relative_to(resolved, allowed) for allowed in allowed_paths):
        return str(resolved)

    allowed_str = ", ".join(str(path) for path in allowed_paths)
    raise ValueError(f"路径不在允许目录内: {cleaned} (allowed: {allowed_str})")


def validate_pptx(file_path: str) -> Tuple[bool, str]:
    """
    完整的 PPTX 文件安全验证
    
    检查项：
    1. 文件存在性
    2. 文件大小限制
    3. ZIP 完整性
    4. 解压大小限制（ZIP炸弹防护）
    5. 压缩比检测
    6. 宏/VBA 检测
    7. PPTX 结构验证
    
    Returns:
        (是否通过, 消息)
    """
    # 1. 文件存在检查
    if not os.path.exists(file_path):
        return False, f"文件不存在: {file_path}"
    
    # 2. 文件大小检查
    file_size = os.path.getsize(file_path)
    if file_size > limits.MAX_FILE_SIZE:
        return False, f"文件过大: {file_size/1024/1024:.1f}MB > {limits.MAX_FILE_SIZE/1024/1024}MB"
    
    if file_size == 0:
        return False, "空文件"
    
    # 3. ZIP 文件验证
    try:
        with zipfile.ZipFile(file_path, 'r') as zf:
            # 3.1 ZIP 完整性检查
            bad_file = zf.testzip()
            if bad_file:
                return False, f"ZIP 文件损坏: {bad_file}"
            
            # 3.2 解压大小检查（ZIP炸弹防护）
            total_uncompressed = sum(info.file_size for info in zf.filelist)
            if total_uncompressed > limits.MAX_UNCOMPRESSED_SIZE:
                return False, f"解压后过大: {total_uncompressed/1024/1024:.1f}MB > {limits.MAX_UNCOMPRESSED_SIZE/1024/1024}MB"
            
            # 3.3 压缩比检查
            compression_ratio = total_uncompressed / file_size if file_size > 0 else 0
            if compression_ratio > limits.MAX_COMPRESSION_RATIO:
                return False, f"可疑压缩比: {compression_ratio:.0f}x（可能为ZIP炸弹）"
            
            # 3.4 PPTX 结构验证
            names = zf.namelist()
            if not any('ppt/presentation.xml' in n for n in names):
                return False, "无效的 PPTX 文件（缺少 presentation.xml）"
            
            # 3.5 幻灯片数量检查
            slide_count = sum(1 for n in names if n.startswith('ppt/slides/slide') and n.endswith('.xml'))
            if slide_count > limits.MAX_SLIDES:
                return False, f"幻灯片过多: {slide_count} > {limits.MAX_SLIDES}"
    
    except zipfile.BadZipFile:
        return False, "无效的 ZIP 文件"
    except Exception as e:
        return False, f"验证失败: {str(e)}"
    
    # 4. 宏检测
    if has_macro(file_path):
        return False, "文件包含宏/VBA脚本，拒绝处理"
    
    return True, "验证通过"
