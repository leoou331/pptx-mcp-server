"""
PPTX 工具管理器
实现所有 MCP 工具（线程安全版本）

修复 P0-1: 使用 Session 级别的锁保护 Presentation 操作
"""
import os
import re
import logging
from typing import Dict, Any, List, Optional, Tuple

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.oxml.ns import qn

from security.validator import validate_pptx, safe_path, limits
from security.session import SessionManager
from security.tempfile import temp_manager

log = logging.getLogger("pptx-server")

# 常量
EMU_PER_INCH = 914400


class PptxTools:
    """PPTX 工具集合（线程安全）"""
    
    def __init__(self, session_manager: SessionManager, work_dir: str):
        self.sessions = session_manager
        self.work_dir = work_dir

    @staticmethod
    def _validate_slide_index(prs: Presentation, slide_index: int) -> int:
        """验证并规范化 slide_index 参数。

        将 slide_index 转为 int 并检查是否在有效范围 [0, len(slides)) 内。

        Args:
            prs: python-pptx Presentation 对象
            slide_index: 幻灯片索引（0-based）

        Returns:
            规范化后的 int 类型 slide_index

        Raises:
            ValueError: slide_index 类型无法转为 int 或超出范围
        """
        if not isinstance(slide_index, int):
            try:
                slide_index = int(slide_index)
            except (ValueError, TypeError):
                raise ValueError(
                    f"slide_index 必须为整数，收到 {type(slide_index).__name__}: {slide_index!r}"
                )
        total = len(prs.slides)
        if slide_index < 0 or slide_index >= total:
            raise ValueError(
                f"slide_index {slide_index} 超出范围 (有效范围: 0-{total - 1})"
            )
        return slide_index
    
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


    def _validate_export_path(self, session_id: str, export_dir: str) -> None:
        """验证导出路径安全性，防止路径遍历攻击。

        Args:
            session_id: 会话 ID（只允许字母数字和 _-）
            export_dir: 导出目录路径

        Raises:
            ValueError: session_id 格式无效或路径遍历检测失败
        """
        if not re.match(r"^[a-zA-Z0-9_-]+$", session_id):
            raise ValueError(f"无效的 session_id: {session_id}")
        real_ed = os.path.realpath(export_dir)
        real_wd = os.path.realpath(self.work_dir)
        if not real_ed.startswith(real_wd + os.sep):
            raise ValueError("Path traversal detected")

    def _iter_picture_shapes(self, shapes) -> list:
        """递归遍历 shapes 集合，yield 所有包含图片的 shape（含 Group 内嵌套）。

        Args:
            shapes: python-pptx shapes 集合

        Yields:
            包含图片数据的 shape 对象
        """
        for shape in shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                yield shape
            elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                yield from self._iter_picture_shapes(shape.shapes)
            else:
                if hasattr(shape, "image"):
                    try:
                        _ = shape.image
                        yield shape
                    except Exception:
                        pass

    def _build_image_info(self, s_idx: int, sh_idx: int, shape) -> Dict[str, Any]:
        """构建单个图片 shape 的元数据信息。

        Args:
            s_idx: 幻灯片索引
            sh_idx: shape 在幻灯片中的索引
            shape: python-pptx shape 对象

        Returns:
            包含图片位置、尺寸、alt text 等信息的字典
        """
        try:
            ct = shape.image.content_type
        except Exception:
            ct = "unknown"
        alt = shape.name or ""
        try:
            # 搜索 PML 和 DrawingML 命名空间下的 cNvPr 以获取 alt text
            el = shape._element.find(f".//{qn('p:cNvPr')}")
            if el is None:
                el = shape._element.find(f".//{qn('a:cNvPr')}")
            if el is not None:
                alt = el.get("descr", shape.name or "")
        except Exception:
            pass
        return {
            "slide_index": s_idx, "shape_index": sh_idx, "name": shape.name,
            "content_type": ct,
            "left_inches": round(shape.left / EMU_PER_INCH, 4) if shape.left else 0,
            "top_inches": round(shape.top / EMU_PER_INCH, 4) if shape.top else 0,
            "width_inches": round(shape.width / EMU_PER_INCH, 4) if shape.width else 0,
            "height_inches": round(shape.height / EMU_PER_INCH, 4) if shape.height else 0,
            "z_order": sh_idx, "alt_text": alt,
        }

    def list_images(self, session_id: str, slide_index: Optional[int] = None) -> Dict[str, Any]:
        """列出演示文稿中的所有图片，返回位置、尺寸、内容类型等信息。

        Args:
            session_id: 会话 ID
            slide_index: 幻灯片索引（可选，不填则返回所有幻灯片的图片）

        Returns:
            包含图片列表和总数的字典
        """
        session = self.sessions.get(session_id)
        with session.lock:
            prs = session.presentation
            if slide_index is not None:
                slide_index = self._validate_slide_index(prs, slide_index)
                slides_to_check = [(slide_index, prs.slides[slide_index])]
            else:
                slides_to_check = list(enumerate(prs.slides))
            images = []
            for s_idx, slide in slides_to_check:
                for sh_idx, shape in enumerate(slide.shapes):
                    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                        images.append(self._build_image_info(s_idx, sh_idx, shape))
                    elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                        for nested in self._iter_picture_shapes(shape.shapes):
                            images.append(self._build_image_info(s_idx, sh_idx, nested))
                    elif hasattr(shape, "image"):
                        try:
                            _ = shape.image
                            images.append(self._build_image_info(s_idx, sh_idx, shape))
                        except Exception:
                            pass
            return {"session_id": session_id, "total_images": len(images), "images": images}

    def _export_inner(self, prs: Presentation, session_id: str,
                       slide_index: Optional[int] = None) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        """内部方法：将图片 blob 导出到工作目录。

        Args:
            prs: python-pptx Presentation 对象
            session_id: 会话 ID
            slide_index: 幻灯片索引（可选）

        Returns:
            (exported, errors) 元组：
            - exported: 成功导出的图片信息列表
            - errors: 导出失败的错误信息列表
        """
        ext_map = {
            "image/jpeg": "jpg", "image/jpg": "jpg", "image/png": "png",
            "image/gif": "gif", "image/bmp": "bmp", "image/tiff": "tiff",
            "image/x-emf": "emf", "image/x-wmf": "wmf",
        }
        if slide_index is not None:
            slide_index = self._validate_slide_index(prs, slide_index)
            its = [(slide_index, prs.slides[slide_index])]
        else:
            its = list(enumerate(prs.slides))
        export_dir = os.path.join(self.work_dir, "exported_images", session_id)
        self._validate_export_path(session_id, export_dir)
        os.makedirs(export_dir, exist_ok=True)
        out: List[Dict[str, Any]] = []
        errors: List[Dict[str, Any]] = []
        for s_idx, slide in its:
            for sh_idx, shape in enumerate(slide.shapes):
                pics: List[Tuple[int, Any]] = []
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    pics.append((sh_idx, shape))
                elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                    for n in self._iter_picture_shapes(shape.shapes):
                        pics.append((sh_idx, n))
                elif hasattr(shape, "image"):
                    try:
                        _ = shape.image
                        pics.append((sh_idx, shape))
                    except Exception:
                        pass
                for oi, ps in pics:
                    try:
                        blob = ps.image.blob
                        ct = ps.image.content_type
                        ext = ext_map.get(ct.lower(), "bin")
                        raw_name = f"slide{s_idx}_shape{oi}_{ps.name}.{ext}"
                        fn = "".join(
                            c if c.isalnum() or c in "._-" else "_"
                            for c in raw_name
                        )[:200]  # 限制文件名长度
                        fp = os.path.join(export_dir, fn)
                        with open(fp, "wb") as f:
                            f.write(blob)
                        # 注册到 temp_manager 以便自动清理
                        try:
                            temp_manager.register(fp)
                        except Exception:
                            pass
                        out.append({
                            "slide_index": s_idx, "shape_index": oi, "name": ps.name,
                            "content_type": ct, "file_path": fp,
                            "left_inches": round(ps.left / EMU_PER_INCH, 4) if ps.left else 0,
                            "top_inches": round(ps.top / EMU_PER_INCH, 4) if ps.top else 0,
                            "width_inches": round(ps.width / EMU_PER_INCH, 4) if ps.width else 0,
                            "height_inches": round(ps.height / EMU_PER_INCH, 4) if ps.height else 0,
                        })
                    except Exception as e:
                        log.warning(f"导出图片失败 slide={s_idx} shape={oi}: {e}")
                        errors.append({
                            "slide_index": s_idx,
                            "shape_index": oi,
                            "name": getattr(ps, "name", "unknown"),
                            "error": str(e),
                        })
        return out, errors

    def export_images(self, session_id: str, slide_index: Optional[int] = None) -> Dict[str, Any]:
        """导出图片到工作目录，通过 temp_manager 管理文件生命周期。

        Args:
            session_id: 会话 ID
            slide_index: 幻灯片索引（可选）

        Returns:
            包含导出结果、失败数量和错误详情的字典
        """
        if not re.match(r"^[a-zA-Z0-9_-]+$", session_id):
            raise ValueError(f"无效的 session_id: {session_id}")
        session = self.sessions.get(session_id)
        with session.lock:
            exported, errors = self._export_inner(session.presentation, session_id, slide_index)
        return {
            "session_id": session_id,
            "exported_count": len(exported),
            "failed_count": len(errors),
            "images": exported,
            "errors": errors,
        }

    def _estimate_shape_role(self, shape, bbox: Dict[str, float],
                             pw: float, ph: float) -> str:
        """根据 shape 类型、位置和尺寸估计其语义角色。

        Args:
            shape: python-pptx shape 对象
            bbox: 包含 left/top/width/height 的边界框字典（单位：英寸）
            pw: 页面宽度（英寸）
            ph: 页面高度（英寸）

        Returns:
            估计的角色字符串，如 'title', 'body', 'image' 等
        """
        st = shape.shape_type
        if st == MSO_SHAPE_TYPE.TABLE:
            return "table"
        if st == MSO_SHAPE_TYPE.CHART:
            return "chart"
        if st == MSO_SHAPE_TYPE.GROUP:
            return "group"
        if st == MSO_SHAPE_TYPE.PICTURE:
            ar = (bbox["width"] * bbox["height"]) / max(pw * ph, 0.001)
            return "hero_image" if ar > 0.3 else ("icon_or_logo" if ar < 0.05 else "image")
        if hasattr(shape, "text") and shape.text:
            text = shape.text.strip()
            try:
                if hasattr(shape, "placeholder_format") and shape.placeholder_format:
                    pt = shape.placeholder_format.type
                    if pt in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE):
                        return "title"
                    if pt == PP_PLACEHOLDER.SUBTITLE:
                        return "subtitle"
                    if pt == PP_PLACEHOLDER.BODY:
                        return "body"
            except Exception:
                pass
            if bbox["top"] < ph * 0.2 and bbox["height"] < ph * 0.15:
                return "title"
            if bbox["top"] < ph * 0.35 and len(text) < 100:
                return "subtitle_or_heading"
            ar = (bbox["width"] * bbox["height"]) / max(pw * ph, 0.001)
            return "body" if ar > 0.2 else "caption_or_label"
        return "decorative_shape"

    def _analyze_layout(self, elements: List[Dict[str, Any]],
                        pw: float, ph: float) -> Dict[str, Any]:
        """分析幻灯片元素的布局特征，包括阅读顺序、密度和重叠。

        注意：density_score 和 whitespace_ratio 是近似值，不考虑元素间的重叠。
        当多个元素重叠时，实际覆盖面积小于各元素面积之和，因此密度可能被高估。

        Args:
            elements: 元素信息列表（每个元素需包含 'bbox' 字段）
            pw: 页面宽度（英寸）
            ph: 页面高度（英寸）

        Returns:
            包含 reading_order, whitespace_ratio, density_score, overlaps 的字典
        """
        if not elements:
            return {"reading_order": [], "whitespace_ratio": 1.0, "density_score": 0.0, "overlaps": []}
        def rk(i: int) -> Tuple[int, float]:
            b = elements[i]["bbox"]
            return (int(b["top"] / max(ph / 10, 0.001)), b["left"] / max(pw, 0.001))
        ro = sorted(range(len(elements)), key=rk)
        # 近似计算：面积之和，未扣除重叠区域
        cov = sum(e["bbox"]["width"] * e["bbox"]["height"] for e in elements)
        d = min(1.0, cov / max(pw * ph, 0.001))
        def ov(b1: Dict, b2: Dict) -> bool:
            return not (
                b1["left"] + b1["width"] <= b2["left"] or
                b2["left"] + b2["width"] <= b1["left"] or
                b1["top"] + b1["height"] <= b2["top"] or
                b2["top"] + b2["height"] <= b1["top"]
            )
        # 限制重叠检测数量防止 O(n²) 性能问题
        MAX_OVERLAP_ELEMENTS = 50
        check_elems = elements[:MAX_OVERLAP_ELEMENTS]
        ovs = [
            [i, j]
            for i in range(len(check_elems))
            for j in range(i + 1, len(check_elems))
            if ov(check_elems[i]["bbox"], check_elems[j]["bbox"])
        ]
        ovs = ovs[:100]  # 最多返回 100 对
        return {
            "reading_order": ro,
            "whitespace_ratio": round(max(0, 1 - d), 3),
            "density_score": round(d, 3),
            "overlaps": ovs,
        }

    def _describe_inner(self, prs: Presentation, slide_index: int) -> Dict[str, Any]:
        """内部方法：构建幻灯片的结构化布局描述。

        Args:
            prs: python-pptx Presentation 对象
            slide_index: 幻灯片索引

        Returns:
            包含元素信息、布局分析等的结构化字典
        """
        slide_index = self._validate_slide_index(prs, slide_index)
        slide = prs.slides[slide_index]
        pw = (prs.slide_width or 9144000) / EMU_PER_INCH
        ph = (prs.slide_height or 6858000) / EMU_PER_INCH
        bg = {"type": "default", "color": None}
        try:
            fill = slide.background.fill
            if fill.type is not None:
                bg["type"] = str(fill.type)
                try:
                    bg["color"] = f"#{fill.fore_color.rgb}"
                except Exception:
                    pass
        except Exception:
            pass
        elements = []
        for sh_idx, shape in enumerate(slide.shapes):
            bbox = {
                "left": round(shape.left / EMU_PER_INCH, 4) if shape.left else 0,
                "top": round(shape.top / EMU_PER_INCH, 4) if shape.top else 0,
                "width": round(shape.width / EMU_PER_INCH, 4) if shape.width else 0,
                "height": round(shape.height / EMU_PER_INCH, 4) if shape.height else 0,
            }
            text_content = ""
            font_info = None
            if hasattr(shape, "text") and shape.text:
                text_content = shape.text[:500]
                try:
                    best_run, best_sz = None, -1
                    for para in shape.text_frame.paragraphs:
                        for run in para.runs:
                            sz = run.font.size.pt if run.font.size else 0
                            if sz > best_sz:
                                best_sz, best_run = sz, run
                    if best_run:
                        font_info = {
                            "size_pt": best_run.font.size.pt if best_run.font.size else None,
                            "bold": best_run.font.bold,
                            "italic": best_run.font.italic,
                        }
                except Exception:
                    pass
            img_ref = None
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                try:
                    img_ref = {
                        "content_type": shape.image.content_type,
                        "size_bytes": len(shape.image.blob),
                    }
                except Exception:
                    img_ref = {"content_type": "unknown", "size_bytes": 0}
            elements.append({
                "shape_index": sh_idx, "type": str(shape.shape_type), "name": shape.name,
                "bbox": bbox, "z_order": sh_idx, "text": text_content,
                "font_info": font_info, "image_ref": img_ref,
                "estimated_role": self._estimate_shape_role(shape, bbox, pw, ph),
            })
        return {
            "slide_index": slide_index,
            "page_size": {"width_inches": round(pw, 4), "height_inches": round(ph, 4)},
            "background": bg,
            "element_count": len(elements),
            "elements": elements,
            "layout_analysis": self._analyze_layout(elements, pw, ph),
        }

    def describe_slide(self, session_id: str, slide_index: int) -> Dict[str, Any]:
        """返回 slide 的结构化布局描述，含所有元素的位置、类型、文本、图片引用及布局分析。

        Args:
            session_id: 会话 ID
            slide_index: 幻灯片索引（0-based）

        Returns:
            结构化布局描述字典
        """
        session = self.sessions.get(session_id)
        with session.lock:
            slide_index = self._validate_slide_index(session.presentation, slide_index)
            r = self._describe_inner(session.presentation, slide_index)
            r["session_id"] = session_id
            return r

    def export_slide_snapshot(self, session_id: str, slide_index: int) -> Dict[str, Any]:
        """导出 slide 结构化快照（单次加锁，fallback 方案）。

        包含布局 JSON + 图片资源导出，不依赖 LibreOffice。

        Args:
            session_id: 会话 ID
            slide_index: 幻灯片索引（0-based）

        Returns:
            包含布局描述、导出图片和错误信息的字典
        """
        if not re.match(r"^[a-zA-Z0-9_-]+$", session_id):
            raise ValueError(f"无效的 session_id: {session_id}")
        session = self.sessions.get(session_id)
        with session.lock:
            prs = session.presentation
            slide_index = self._validate_slide_index(prs, slide_index)
            desc = self._describe_inner(prs, slide_index)
            imgs, errors = self._export_inner(prs, session_id, slide_index)
        return {
            "session_id": session_id, "slide_index": slide_index,
            "snapshot_type": "structural_layout",
            "note": "PNG rendering requires LibreOffice; returning structural layout JSON + image exports as fallback",
            "page_size": desc["page_size"], "background": desc["background"],
            "element_count": desc["element_count"], "elements": desc["elements"],
            "layout_analysis": desc["layout_analysis"],
            "exported_images": imgs,
            "export_failed_count": len(errors),
            "export_errors": errors,
        }

    def get_animation_info(self, session_id: str, slide_index: int) -> Dict[str, Any]:
        """获取 slide 动画和 transition 信息（通过解析 PML XML）。

        Args:
            session_id: 会话 ID
            slide_index: 幻灯片索引（0-based）

        Returns:
            包含动画列表、transition 信息等的字典
        """
        # nodeType 到 trigger 的映射表（基于 OOXML 规范）
        NODE_TYPE_TO_TRIGGER = {
            "clickEffect": "onClick",
            "withEffect": "withPrevious",
            "afterEffect": "afterPrevious",
            "mainSeq": "mainSequence",
            "interactiveSeq": "interactive",
        }
        session = self.sessions.get(session_id)
        with session.lock:
            prs = session.presentation
            slide_index = self._validate_slide_index(prs, slide_index)
            slide = prs.slides[slide_index]
            se = slide._element
            PML = "http://schemas.openxmlformats.org/presentationml/2006/main"
            te = se.find(f"{{{PML}}}transition")
            has_t = te is not None
            ti = None
            if has_t:
                ti = {
                    "type": "unknown",
                    "duration_ms": None,
                    "advance_on_click": te.get("advClick", "true").lower() != "false",
                    "advance_after_time_ms": None,
                }
                dur = te.get("dur")
                if dur:
                    try:
                        ti["duration_ms"] = int(dur)
                    except ValueError:
                        ti["duration_ms"] = dur
                adv = te.get("advTm")
                if adv:
                    try:
                        ti["advance_after_time_ms"] = int(adv)
                    except ValueError:
                        ti["advance_after_time_ms"] = adv
                for ch in te:
                    t = ch.tag.split("}")[-1] if "}" in ch.tag else ch.tag
                    if t != "extLst":
                        ti["type"] = t
                        break
            anims = []
            anis = set()
            tm = se.find(f"{{{PML}}}timing")
            if tm is not None:
                order = 0
                for par in tm.iter(f"{{{PML}}}par"):
                    par_cTn = par.find(f"{{{PML}}}cTn")
                    trigger = "onClick"
                    delay_ms = 0
                    seq_dur = None
                    if par_cTn is not None:
                        nt = par_cTn.get("nodeType", "")
                        trigger = NODE_TYPE_TO_TRIGGER.get(nt, "unknown" if nt else "onClick")
                        d = par_cTn.get("delay", "0")
                        if d and d != "indefinite":
                            try:
                                delay_ms = int(d)
                            except Exception:
                                pass
                        dv = par_cTn.get("dur")
                        if dv and dv != "indefinite":
                            try:
                                seq_dur = int(dv)
                            except Exception:
                                pass
                    for tgt in par.findall(f".//{{{PML}}}spTgt"):
                        sp_id = tgt.get("spid")
                        sn = None
                        si = None
                        if sp_id:
                            for idx, sh in enumerate(slide.shapes):
                                try:
                                    if str(sh.shape_id) == str(sp_id):
                                        sn = sh.name
                                        si = idx
                                        anis.add(idx)
                                        break
                                except Exception:
                                    pass
                        ef_dur = seq_dur
                        for cb in par.findall(f".//{{{PML}}}cBhvr"):
                            ic = cb.find(f"{{{PML}}}cTn")
                            if ic is not None:
                                dv2 = ic.get("dur")
                                if dv2 and dv2 != "indefinite":
                                    try:
                                        ef_dur = int(dv2)
                                    except Exception:
                                        pass
                                break
                        et = "unknown"
                        for el in par.iter():
                            t = el.tag.split("}")[-1] if "}" in el.tag else el.tag
                            if t in ("animEffect", "anim", "animMotion", "animScale", "animRot", "set", "animClr"):
                                et = t
                                break
                        anims.append({
                            "order": order, "shape_name": sn, "shape_index": si,
                            "effect_type": et, "trigger": trigger,
                            "duration_ms": ef_dur, "delay_ms": delay_ms,
                        })
                        order += 1
            return {
                "session_id": session_id, "slide_index": slide_index,
                "has_animations": len(anims) > 0, "has_transition": has_t,
                "transition_info": ti, "animation_count": len(anims),
                "animations": anims, "animated_shape_indices": sorted(anis),
            }

