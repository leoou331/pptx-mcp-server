"""
PPTX 工具管理器
实现所有 MCP 工具（线程安全版本）

修复 P0-1: 使用 Session 级别的锁保护 Presentation 操作
"""
import os
import logging
from typing import Dict, Any, Optional

from pptx import Presentation
from pptx.util import Inches, Pt

from security.validator import validate_pptx, safe_path, limits
from security.session import SessionManager
from security.tempfile import temp_manager

log = logging.getLogger("pptx-server")


class PptxTools:
    """PPTX 工具集合（线程安全）"""
    
    def __init__(self, session_manager: SessionManager, work_dir: str):
        self.sessions = session_manager
        self.work_dir = work_dir
    
    def create(self, name: str = "Untitled") -> Dict[str, Any]:
        """
        创建空白演示文稿
        
        Args:
            name: 演示文稿名称
            
        Returns:
            会话信息
        """
        session_id = self.sessions.create(name)
        
        return {
            "session_id": session_id,
            "name": name,
            "slide_count": 0,
            "message": "演示文稿已创建"
        }
    
    def open(self, file_path: str) -> Dict[str, Any]:
        """
        打开现有文件
        
        Args:
            file_path: 文件路径
            
        Returns:
            会话信息
        """
        # 路径安全检查
        if not os.path.isabs(file_path):
            file_path = safe_path(self.work_dir, file_path)
        
        session_id = self.sessions.open(file_path)
        session = self.sessions.get(session_id)
        
        with session.lock:  # 使用 Session 级别的锁
            return {
                "session_id": session_id,
                "name": session.name,
                "source_file": file_path,
                "slide_count": len(session.presentation.slides),
                "message": "文件已打开"
            }
    
    def save(self, session_id: str, output_path: Optional[str] = None) -> Dict[str, Any]:
        """
        保存演示文稿
        
        Args:
            session_id: 会话 ID
            output_path: 输出路径（可选）
            
        Returns:
            保存结果
        """
        session = self.sessions.get(session_id)
        
        if output_path:
            output_path = safe_path(self.work_dir, output_path)
        else:
            output_path = temp_manager.create(suffix='.pptx')
        
        with session.lock:  # 使用 Session 级别的锁
            prs = session.presentation
            prs.save(output_path)
            session.dirty = False
            
            return {
                "file_path": output_path,
                "slide_count": len(prs.slides),
                "message": "文件已保存"
            }
    
    def close(self, session_id: str) -> Dict[str, Any]:
        """
        关闭会话
        
        Args:
            session_id: 会话 ID
            
        Returns:
            关闭结果
        """
        closed = self.sessions.close(session_id)
        return {
            "closed": closed,
            "message": "会话已关闭" if closed else "会话不存在"
        }
    
    def info(self, session_id: str) -> Dict[str, Any]:
        """
        获取演示文稿信息
        
        Args:
            session_id: 会话 ID
            
        Returns:
            演示文稿信息
        """
        session = self.sessions.get(session_id)
        
        with session.lock:  # 使用 Session 级别的锁
            prs = session.presentation
            
            # 收集幻灯片信息
            slides_info = []
            for i, slide in enumerate(prs.slides):
                text_preview = []
                for shape in slide.shapes:
                    if hasattr(shape, "text") and shape.text:
                        text_preview.append(shape.text[:100])
                
                slides_info.append({
                    "index": i,
                    "shape_count": len(slide.shapes),
                    "text_preview": " | ".join(text_preview)[:200]
                })
            
            return {
                "session_id": session_id,
                "name": session.name,
                "slide_count": len(prs.slides),
                "slide_width_inches": round(prs.slide_width.inches, 2),
                "slide_height_inches": round(prs.slide_height.inches, 2),
                "source_file": session.source_file,
                "dirty": session.dirty,
                "slides": slides_info
            }
    
    def add_slide(self, session_id: str, layout_index: int = 0) -> Dict[str, Any]:
        """
        添加幻灯片
        
        Args:
            session_id: 会话 ID
            layout_index: 布局索引
            
        Returns:
            添加结果
        """
        session = self.sessions.get(session_id)
        
        with session.lock:  # 使用 Session 级别的锁
            prs = session.presentation
            
            # 检查幻灯片数量限制
            if len(prs.slides) >= limits.MAX_SLIDES:
                raise ValueError(f"幻灯片数量已达上限: {limits.MAX_SLIDES}")
            
            # 验证布局索引
            if layout_index >= len(prs.slide_layouts):
                layout_index = 0
            
            slide = prs.slides.add_slide(prs.slide_layouts[layout_index])
            session.dirty = True
            
            return {
                "slide_index": len(prs.slides) - 1,
                "layout_index": layout_index,
                "total_slides": len(prs.slides),
                "message": "幻灯片已添加"
            }
    
    def add_text(
        self,
        session_id: str,
        slide_index: int,
        text: str,
        position: str = "body",
        left: float = 1.0,
        top: float = 1.0,
        width: float = 8.0,
        height: float = 1.0,
        font_size: int = 18
    ) -> Dict[str, Any]:
        """
        添加文本
        
        Args:
            session_id: 会话 ID
            slide_index: 幻灯片索引
            text: 文本内容
            position: 位置类型 (title/body/custom)
            left, top, width, height: 位置和尺寸（英寸）
            font_size: 字号
            
        Returns:
            添加结果
        """
        session = self.sessions.get(session_id)
        
        # 文本长度检查
        if len(text) > limits.MAX_TEXT_LENGTH:
            raise ValueError(f"文本过长: {len(text)} > {limits.MAX_TEXT_LENGTH}")
        
        with session.lock:  # 使用 Session 级别的锁
            prs = session.presentation
            
            # 验证幻灯片索引
            if slide_index < 0 or slide_index >= len(prs.slides):
                raise ValueError(f"幻灯片索引越界: {slide_index}")
            
            slide = prs.slides[slide_index]
            
            # 根据位置类型处理
            if position == "title" and slide.shapes.title:
                slide.shapes.title.text = text
            else:
                # 创建文本框
                textbox = slide.shapes.add_textbox(
                    Inches(left),
                    Inches(top),
                    Inches(width),
                    Inches(height)
                )
                text_frame = textbox.text_frame
                text_frame.text = text
                
                # 设置字号
                if font_size:
                    for paragraph in text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.font.size = Pt(font_size)
            
            session.dirty = True
            
            return {
                "slide_index": slide_index,
                "position": position,
                "text_length": len(text),
                "message": "文本已添加"
            }
    
    def add_image(
        self,
        session_id: str,
        slide_index: int,
        image_path: str,
        left: float,
        top: float,
        width: Optional[float] = None,
        height: Optional[float] = None
    ) -> Dict[str, Any]:
        """
        添加图片
        
        Args:
            session_id: 会话 ID
            slide_index: 幻灯片索引
            image_path: 图片路径
            left, top: 位置（英寸）
            width, height: 尺寸（英寸，可选）
            
        Returns:
            添加结果
        """
        session = self.sessions.get(session_id)
        
        # 路径安全检查
        if not os.path.isabs(image_path):
            image_path = safe_path(self.work_dir, image_path)
        
        # 检查文件存在
        if not os.path.exists(image_path):
            raise FileNotFoundError(f"图片不存在: {image_path}")
        
        # 检查图片大小
        image_size = os.path.getsize(image_path)
        if image_size > limits.MAX_IMAGE_SIZE:
            raise ValueError(f"图片过大: {image_size/1024/1024:.1f}MB > {limits.MAX_IMAGE_SIZE/1024/1024}MB")
        
        with session.lock:  # 使用 Session 级别的锁
            prs = session.presentation
            
            # 验证幻灯片索引
            if slide_index < 0 or slide_index >= len(prs.slides):
                raise ValueError(f"幻灯片索引越界: {slide_index}")
            
            slide = prs.slides[slide_index]
            
            # 添加图片
            if width and height:
                slide.shapes.add_picture(
                    image_path,
                    Inches(left),
                    Inches(top),
                    Inches(width),
                    Inches(height)
                )
            elif width:
                slide.shapes.add_picture(
                    image_path,
                    Inches(left),
                    Inches(top),
                    width=Inches(width)
                )
            else:
                slide.shapes.add_picture(
                    image_path,
                    Inches(left),
                    Inches(top)
                )
            
            session.dirty = True
            
            return {
                "slide_index": slide_index,
                "image_path": image_path,
                "image_size_kb": round(image_size / 1024, 1),
                "message": "图片已添加"
            }
    
    def add_table(
        self,
        session_id: str,
        slide_index: int,
        rows: int,
        cols: int,
        left: float = 1.0,
        top: float = 2.0,
        width: float = 8.0,
        height: float = 4.0,
        data: Optional[list] = None
    ) -> Dict[str, Any]:
        """
        添加表格
        
        Args:
            session_id: 会话 ID
            slide_index: 幻灯片索引
            rows, cols: 行数和列数
            left, top, width, height: 位置和尺寸（英寸）
            data: 表格数据（可选）
            
        Returns:
            添加结果
        """
        session = self.sessions.get(session_id)
        
        with session.lock:  # 使用 Session 级别的锁
            prs = session.presentation
            
            # 验证幻灯片索引
            if slide_index < 0 or slide_index >= len(prs.slides):
                raise ValueError(f"幻灯片索引越界: {slide_index}")
            
            slide = prs.slides[slide_index]
            
            # 添加表格
            table = slide.shapes.add_table(
                rows, cols,
                Inches(left),
                Inches(top),
                Inches(width),
                Inches(height)
            ).table
            
            # 填充数据
            if data:
                for i, row_data in enumerate(data[:rows]):
                    for j, cell_data in enumerate(row_data[:cols]):
                        if cell_data is not None:
                            table.cell(i, j).text = str(cell_data)
            
            session.dirty = True
            
            return {
                "slide_index": slide_index,
                "rows": rows,
                "cols": cols,
                "has_data": data is not None,
                "message": "表格已添加"
            }
    
    def read_content(self, session_id: str) -> Dict[str, Any]:
        """
        读取所有文本内容
        
        Args:
            session_id: 会话 ID
            
        Returns:
            所有幻灯片的文本内容
        """
        session = self.sessions.get(session_id)
        
        with session.lock:  # 使用 Session 级别的锁
            prs = session.presentation
            
            content = []
            for i, slide in enumerate(prs.slides):
                slide_text = []
                for shape in slide.shapes:
                    if hasattr(shape, "text") and shape.text:
                        slide_text.append(shape.text)
                
                content.append({
                    "slide_index": i,
                    "text": "\n".join(slide_text),
                    "char_count": sum(len(t) for t in slide_text)
                })
            
            return {
                "session_id": session_id,
                "total_slides": len(prs.slides),
                "total_chars": sum(c["char_count"] for c in content),
                "slides": content
            }
    
    def list_slides(self, session_id: str) -> Dict[str, Any]:
        """
        列出所有幻灯片
        
        Args:
            session_id: 会话 ID
            
        Returns:
            幻灯片列表
        """
        session = self.sessions.get(session_id)
        
        with session.lock:  # 使用 Session 级别的锁
            prs = session.presentation
            
            slides = []
            for i, slide in enumerate(prs.slides):
                shapes_info = []
                for shape in slide.shapes:
                    shape_type = str(shape.shape_type)
                    has_text = hasattr(shape, "text") and bool(shape.text)
                    
                    shape_info = {
                        "type": shape_type,
                        "has_text": has_text,
                        "text_preview": shape.text[:50] if has_text and shape.text else None
                    }
                    shapes_info.append(shape_info)
                
                slides.append({
                    "index": i,
                    "shape_count": len(slide.shapes),
                    "shapes": shapes_info[:10]  # 限制返回数量
                })
            
            return {
                "session_id": session_id,
                "total_slides": len(prs.slides),
                "slides": slides
            }
    
    def validate(self, file_path: str) -> Dict[str, Any]:
        """
        验证文件安全性
        
        Args:
            file_path: 文件路径
            
        Returns:
            验证结果
        """
        # 路径安全检查
        if not os.path.isabs(file_path):
            file_path = safe_path(self.work_dir, file_path)
        
        valid, message = validate_pptx(file_path)
        
        result = {
            "file_path": file_path,
            "valid": valid,
            "message": message
        }
        
        if valid:
            # 添加文件信息
            try:
                prs = Presentation(file_path)
                result["slide_count"] = len(prs.slides)
                result["slide_width_inches"] = round(prs.slide_width.inches, 2)
                result["slide_height_inches"] = round(prs.slide_height.inches, 2)
            except Exception as e:
                result["parse_error"] = str(e)
        
        return result
