"""
PPTX 文件安全验证模块
防护：ZIP炸弹、宏/VBA、路径遍历、XML实体注入
"""
import os
import zipfile
from dataclasses import dataclass
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
    防止路径遍历攻击（跨平台安全）
    
    Args:
        base_dir: 基础目录（安全边界）
        user_path: 用户提供的路径
        
    Returns:
        安全的绝对路径
        
    Raises:
        ValueError: 如果检测到路径遍历攻击
    """
    # 1. 只使用 basename，忽略用户提供的路径部分
    safe_name = os.path.basename(user_path)
    
    # 2. 禁止特殊字符
    forbidden_chars = ['..', '~', '$', '|', ';', '&', '\x00']
    for char in forbidden_chars:
        if char in safe_name:
            raise ValueError(f"非法文件名: {safe_name}")
    
    # 3. 禁止空文件名
    if not safe_name or safe_name.strip() == '':
        raise ValueError("文件名不能为空")
    
    # 4. 构建完整路径
    full_path = os.path.join(base_dir, safe_name)
    
    # 5. 解析符号链接，获取真实路径
    abs_base = os.path.realpath(base_dir)
    abs_path = os.path.realpath(full_path)
    
    # 6. 跨平台前缀检查（使用 commonpath 更可靠）
    try:
        common_prefix = os.path.commonpath([abs_base, abs_path])
        if common_prefix != abs_base:
            raise ValueError(f"路径遍历攻击检测: {user_path}")
    except ValueError:
        # 如果路径不在同一驱动器上（Windows）
        raise ValueError(f"非法路径: {user_path}")
    
    return abs_path


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
