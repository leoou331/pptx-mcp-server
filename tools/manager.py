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

    def list_images(self, session_id: str, slide_index=None):
        """列出演示文稿中的所有图片"""
        from pptx.enum.shapes import MSO_SHAPE_TYPE
        session = self.sessions.get(session_id)
        with session.lock:
            prs = session.presentation
            if slide_index is not None:
                if slide_index < 0 or slide_index >= len(prs.slides):
                    raise ValueError(f"幻灯片索引越界: {slide_index}")
                slides_to_check = [(slide_index, prs.slides[slide_index])]
            else:
                slides_to_check = list(enumerate(prs.slides))
            images = []
            for s_idx, slide in slides_to_check:
                for sh_idx, shape in enumerate(slide.shapes):
                    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                        try:
                            content_type = shape.image.content_type
                        except Exception:
                            content_type = "unknown"
                        alt_text = shape.name or ""
                        try:
                            cNvPr = shape._element.find(
                                ".//{http://schemas.openxmlformats.org/drawingml/2006/main}cNvPr"
                            )
                            if cNvPr is not None:
                                alt_text = cNvPr.get("descr", shape.name or "")
                        except Exception:
                            pass
                        images.append({
                            "slide_index": s_idx,
                            "shape_index": sh_idx,
                            "name": shape.name,
                            "content_type": content_type,
                            "left_inches": round(shape.left / 914400, 4) if shape.left else 0,
                            "top_inches": round(shape.top / 914400, 4) if shape.top else 0,
                            "width_inches": round(shape.width / 914400, 4) if shape.width else 0,
                            "height_inches": round(shape.height / 914400, 4) if shape.height else 0,
                            "z_order": sh_idx,
                            "alt_text": alt_text,
                        })
            return {"session_id": session_id, "total_images": len(images), "images": images}

    def export_images(self, session_id: str, slide_index=None):
        """导出图片到临时目录"""
        import os
        from pptx.enum.shapes import MSO_SHAPE_TYPE
        session = self.sessions.get(session_id)
        with session.lock:
            prs = session.presentation
            if slide_index is not None:
                if slide_index < 0 or slide_index >= len(prs.slides):
                    raise ValueError(f"幻灯片索引越界: {slide_index}")
                slides_to_check = [(slide_index, prs.slides[slide_index])]
            else:
                slides_to_check = list(enumerate(prs.slides))
            exported = []
            ext_map = {
                "image/jpeg": "jpg", "image/jpg": "jpg", "image/png": "png",
                "image/gif": "gif", "image/bmp": "bmp", "image/tiff": "tiff",
                "image/x-emf": "emf", "image/x-wmf": "wmf",
            }
            for s_idx, slide in slides_to_check:
                for sh_idx, shape in enumerate(slide.shapes):
                    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                        try:
                            img_blob = shape.image.blob
                            content_type = shape.image.content_type
                            ext = ext_map.get(content_type.lower(), "bin")
                            export_dir = os.path.join(self.work_dir, "exported_images", session_id)
                            os.makedirs(export_dir, exist_ok=True)
                            raw_name = f"slide{s_idx}_shape{sh_idx}_{shape.name}.{ext}"
                            filename = "".join(c if c.isalnum() or c in "._-" else "_" for c in raw_name)
                            file_path = os.path.join(export_dir, filename)
                            with open(file_path, "wb") as f:
                                f.write(img_blob)
                            exported.append({
                                "slide_index": s_idx,
                                "shape_index": sh_idx,
                                "name": shape.name,
                                "content_type": content_type,
                                "file_path": file_path,
                                "left_inches": round(shape.left / 914400, 4) if shape.left else 0,
                                "top_inches": round(shape.top / 914400, 4) if shape.top else 0,
                                "width_inches": round(shape.width / 914400, 4) if shape.width else 0,
                                "height_inches": round(shape.height / 914400, 4) if shape.height else 0,
                            })
                        except Exception as e:
                            log.warning(f"导出图片失败 slide={s_idx} shape={sh_idx}: {e}")
            return {"session_id": session_id, "exported_count": len(exported), "images": exported}

    def _estimate_shape_role(self, shape, bbox, pw, ph):
        """启发式估计 shape 语义角色"""
        from pptx.enum.shapes import MSO_SHAPE_TYPE
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            area_ratio = (bbox["width"] * bbox["height"]) / max(pw * ph, 0.001)
            if area_ratio > 0.3:
                return "hero_image"
            elif area_ratio < 0.05:
                return "icon_or_logo"
            return "image"
        if hasattr(shape, "text") and shape.text:
            text = shape.text.strip()
            try:
                from pptx.enum.shapes import PP_PLACEHOLDER
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
            area_ratio = (bbox["width"] * bbox["height"]) / max(pw * ph, 0.001)
            if area_ratio > 0.2:
                return "body"
            return "caption_or_label"
        return "decorative_shape"

    def _analyze_layout(self, elements, pw, ph):
        """简单布局分析"""
        if not elements:
            return {"reading_order": [], "whitespace_ratio": 1.0, "density_score": 0.0}
        def rk(i):
            b = elements[i]["bbox"]
            return (int(b["top"] / max(ph / 10, 0.001)), b["left"] / max(pw, 0.001))
        reading_order = sorted(range(len(elements)), key=rk)
        covered = sum(e["bbox"]["width"] * e["bbox"]["height"] for e in elements)
        total = max(pw * ph, 0.001)
        density = min(1.0, covered / total)
        return {
            "reading_order": reading_order,
            "whitespace_ratio": round(max(0, 1 - density), 3),
            "density_score": round(density, 3),
        }

    def describe_slide(self, session_id: str, slide_index: int):
        """返回 slide 的结构化布局描述"""
        from pptx.enum.shapes import MSO_SHAPE_TYPE
        session = self.sessions.get(session_id)
        with session.lock:
            prs = session.presentation
            if slide_index < 0 or slide_index >= len(prs.slides):
                raise ValueError(f"幻灯片索引越界: {slide_index}")
            slide = prs.slides[slide_index]
            pw = prs.slide_width / 914400
            ph = prs.slide_height / 914400
            bg_info = {"type": "default", "color": None}
            try:
                fill = slide.background.fill
                if fill.type is not None:
                    bg_info["type"] = str(fill.type)
                    try:
                        bg_info["color"] = f"#{fill.fore_color.rgb}"
                    except Exception:
                        pass
            except Exception:
                pass
            elements = []
            for sh_idx, shape in enumerate(slide.shapes):
                bbox = {
                    "left": round(shape.left / 914400, 4) if shape.left else 0,
                    "top": round(shape.top / 914400, 4) if shape.top else 0,
                    "width": round(shape.width / 914400, 4) if shape.width else 0,
                    "height": round(shape.height / 914400, 4) if shape.height else 0,
                }
                text_content = ""
                font_info = None
                if hasattr(shape, "text") and shape.text:
                    text_content = shape.text[:500]
                    try:
                        tf = shape.text_frame
                        if tf.paragraphs and tf.paragraphs[0].runs:
                            run = tf.paragraphs[0].runs[0]
                            font_info = {
                                "size_pt": run.font.size.pt if run.font.size else None,
                                "bold": run.font.bold,
                                "italic": run.font.italic,
                            }
                    except Exception:
                        pass
                image_ref = None
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    try:
                        image_ref = {
                            "content_type": shape.image.content_type,
                            "size_bytes": len(shape.image.blob),
                        }
                    except Exception:
                        image_ref = {"content_type": "unknown", "size_bytes": 0}
                elements.append({
                    "shape_index": sh_idx,
                    "type": str(shape.shape_type),
                    "name": shape.name,
                    "bbox": bbox,
                    "z_order": sh_idx,
                    "text": text_content,
                    "font_info": font_info,
                    "image_ref": image_ref,
                    "estimated_role": self._estimate_shape_role(shape, bbox, pw, ph),
                })
            return {
                "session_id": session_id,
                "slide_index": slide_index,
                "page_size": {"width_inches": round(pw, 4), "height_inches": round(ph, 4)},
                "background": bg_info,
                "element_count": len(elements),
                "elements": elements,
                "layout_analysis": self._analyze_layout(elements, pw, ph),
            }

    def export_slide_snapshot(self, session_id: str, slide_index: int):
        """导出 slide 结构化快照（fallback 方案）"""
        desc = self.describe_slide(session_id, slide_index)
        exp = self.export_images(session_id, slide_index=slide_index)
        return {
            "session_id": session_id,
            "slide_index": slide_index,
            "snapshot_type": "structural_layout",
            "note": "直接 PNG 渲染需要 LibreOffice 等额外依赖，返回结构化布局 JSON + 图片资源作为 fallback",
            "page_size": desc["page_size"],
            "background": desc["background"],
            "element_count": desc["element_count"],
            "elements": desc["elements"],
            "layout_analysis": desc["layout_analysis"],
            "exported_images": exp["images"],
        }

    def get_animation_info(self, session_id: str, slide_index: int):
        """获取 slide 动画和 transition 信息（通过 XML 解析）"""
        session = self.sessions.get(session_id)
        with session.lock:
            prs = session.presentation
            if slide_index < 0 or slide_index >= len(prs.slides):
                raise ValueError(f"幻灯片索引越界: {slide_index}")
            slide = prs.slides[slide_index]
            slide_elem = slide._element
            PML = "http://schemas.openxmlformats.org/presentationml/2006/main"

            # Transition
            trans_elem = slide_elem.find(f"{{{PML}}}transition")
            has_transition = trans_elem is not None
            transition_info = None
            if has_transition:
                transition_info = {
                    "type": "unknown",
                    "duration_ms": None,
                    "advance_on_click": trans_elem.get("advClick", "true").lower() != "false",
                    "advance_after_time_ms": None,
                }
                dur = trans_elem.get("dur")
                if dur:
                    try:
                        transition_info["duration_ms"] = int(dur)
                    except ValueError:
                        transition_info["duration_ms"] = dur
                adv_tm = trans_elem.get("advTm")
                if adv_tm:
                    try:
                        transition_info["advance_after_time_ms"] = int(adv_tm)
                    except ValueError:
                        transition_info["advance_after_time_ms"] = adv_tm
                for child in trans_elem:
                    tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
                    if tag != "extLst":
                        transition_info["type"] = tag
                        break

            # Animations
            animations = []
            animated_shape_indices = set()
            timing_elem = slide_elem.find(f"{{{PML}}}timing")
            if timing_elem is not None:
                order = 0
                for par in timing_elem.iter(f"{{{PML}}}par"):
                    for tgt in par.findall(f".//{{{PML}}}spTgt"):
                        sp_id = tgt.get("spid")
                        shape_name = None
                        shape_idx = None
                        if sp_id:
                            for idx, sh in enumerate(slide.shapes):
                                try:
                                    if str(sh.shape_id) == str(sp_id):
                                        shape_name = sh.name
                                        shape_idx = idx
                                        animated_shape_indices.add(idx)
                                        break
                                except Exception:
                                    pass
                        trigger = "onClick"
                        delay_ms = 0
                        duration_ms = None
                        cTn = par.find(f".//{{{PML}}}cTn")
                        if cTn is not None:
                            try:
                                d = cTn.get("delay", "0")
                                if d and d != "indefinite":
                                    delay_ms = int(d)
                                dv = cTn.get("dur")
                                if dv and dv != "indefinite":
                                    duration_ms = int(dv)
                            except Exception:
                                pass
                        effect_type = "unknown"
                        for el in par.iter():
                            tag = el.tag.split("}")[-1] if "}" in el.tag else el.tag
                            if tag in ("animEffect", "anim", "animMotion", "animScale", "animRot", "set"):
                                effect_type = tag
                                break
                        animations.append({
                            "order": order,
                            "shape_name": shape_name,
                            "shape_index": shape_idx,
                            "effect_type": effect_type,
                            "trigger": trigger,
                            "duration_ms": duration_ms,
                            "delay_ms": delay_ms,
                        })
                        order += 1
            return {
                "session_id": session_id,
                "slide_index": slide_index,
                "has_animations": len(animations) > 0,
                "has_transition": has_transition,
                "transition_info": transition_info,
                "animation_count": len(animations),
                "animations": animations,
                "animated_shape_indices": sorted(animated_shape_indices),
            }

