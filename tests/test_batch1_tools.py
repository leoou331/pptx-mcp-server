"""
Batch1 新增工具测试：
- pptx_add_shape
- pptx_add_chart
- pptx_manage_text
- pptx_format_table_cell
"""
import pytest

from security.session import SessionManager
from tools.manager import PptxTools


@pytest.fixture
def tools_env(tmp_path):
    """创建工具实例和带幻灯片的会话。"""
    manager = SessionManager()
    tools = PptxTools(manager, str(tmp_path))
    session_id = tools.create("TestDeck")["session_id"]
    tools.add_slide(session_id, layout_index=0)
    return tools, session_id


# ===== pptx_add_shape =====

class TestAddShape:
    def test_add_rectangle(self, tools_env):
        tools, sid = tools_env
        result = tools.add_shape(
            session_id=sid,
            slide_index=0,
            shape_type="RECTANGLE",
            left=914400,
            top=914400,
            width=1828800,
            height=914400,
        )
        assert result["message"] == "形状已添加"
        assert result["shape_type"] == "RECTANGLE"
        assert result["slide_index"] == 0

    def test_add_oval_with_text_and_colors(self, tools_env):
        tools, sid = tools_env
        result = tools.add_shape(
            session_id=sid,
            slide_index=0,
            shape_type="OVAL",
            left=0,
            top=0,
            width=914400,
            height=914400,
            text="Hello",
            fill_color="FF0000",
            line_color="0000FF",
        )
        assert result["message"] == "形状已添加"
        assert result["shape_type"] == "OVAL"

    def test_add_shape_invalid_type(self, tools_env):
        tools, sid = tools_env
        with pytest.raises(ValueError, match="不支持的 shape_type"):
            tools.add_shape(
                session_id=sid,
                slide_index=0,
                shape_type="NONEXISTENT",
                left=0, top=0, width=914400, height=914400,
            )

    def test_add_shape_invalid_color(self, tools_env):
        tools, sid = tools_env
        with pytest.raises(ValueError, match="无效的颜色格式"):
            tools.add_shape(
                session_id=sid,
                slide_index=0,
                shape_type="RECTANGLE",
                left=0, top=0, width=914400, height=914400,
                fill_color="ZZZZZZ",
            )

    def test_add_shape_zero_width(self, tools_env):
        tools, sid = tools_env
        with pytest.raises(ValueError, match="width 必须大于 0"):
            tools.add_shape(
                session_id=sid,
                slide_index=0,
                shape_type="RECTANGLE",
                left=0, top=0, width=0, height=914400,
            )

    def test_add_shape_case_insensitive(self, tools_env):
        tools, sid = tools_env
        result = tools.add_shape(
            session_id=sid,
            slide_index=0,
            shape_type="triangle",
            left=0, top=0, width=914400, height=914400,
        )
        assert result["shape_type"] == "TRIANGLE"


# ===== pptx_add_chart =====

class TestAddChart:
    def test_add_column_chart(self, tools_env):
        tools, sid = tools_env
        result = tools.add_chart(
            session_id=sid,
            slide_index=0,
            chart_type="CLUSTERED_COLUMN",
            categories=["Q1", "Q2", "Q3"],
            series_data={"Sales": [100, 200, 150]},
            left=914400,
            top=914400,
            width=5486400,
            height=3657600,
        )
        assert result["message"] == "图表已添加"
        assert result["chart_type"] == "CLUSTERED_COLUMN"
        assert result["categories_count"] == 3
        assert result["series_count"] == 1

    def test_add_pie_chart_with_title(self, tools_env):
        tools, sid = tools_env
        result = tools.add_chart(
            session_id=sid,
            slide_index=0,
            chart_type="PIE",
            categories=["A", "B", "C"],
            series_data={"Share": [30, 50, 20]},
            left=0, top=0, width=5486400, height=3657600,
            title="Market Share",
        )
        assert result["has_title"] is True

    def test_add_chart_multiple_series(self, tools_env):
        tools, sid = tools_env
        result = tools.add_chart(
            session_id=sid,
            slide_index=0,
            chart_type="LINE",
            categories=["Jan", "Feb"],
            series_data={
                "Series A": [10, 20],
                "Series B": [30, 40],
            },
            left=0, top=0, width=5486400, height=3657600,
        )
        assert result["series_count"] == 2

    def test_add_chart_invalid_type(self, tools_env):
        tools, sid = tools_env
        with pytest.raises(ValueError, match="不支持的 chart_type"):
            tools.add_chart(
                session_id=sid,
                slide_index=0,
                chart_type="RADAR",
                categories=["A"],
                series_data={"S": [1]},
                left=0, top=0, width=5486400, height=3657600,
            )

    def test_add_chart_mismatched_lengths(self, tools_env):
        tools, sid = tools_env
        with pytest.raises(ValueError, match="长度.*不匹配"):
            tools.add_chart(
                session_id=sid,
                slide_index=0,
                chart_type="CLUSTERED_COLUMN",
                categories=["A", "B"],
                series_data={"S": [1, 2, 3]},
                left=0, top=0, width=5486400, height=3657600,
            )

    def test_add_chart_empty_categories(self, tools_env):
        tools, sid = tools_env
        with pytest.raises(ValueError, match="categories 必须是非空列表"):
            tools.add_chart(
                session_id=sid,
                slide_index=0,
                chart_type="LINE",
                categories=[],
                series_data={"S": []},
                left=0, top=0, width=5486400, height=3657600,
            )


# ===== pptx_manage_text =====

class TestManageText:
    def test_add_text(self, tools_env):
        tools, sid = tools_env
        result = tools.manage_text(
            session_id=sid,
            operation="add",
            slide_index=0,
            text="Hello World",
            left=914400,
            top=914400,
            width=5486400,
            height=914400,
            font_size=24,
            bold=True,
            alignment="center",
        )
        assert result["operation"] == "add"
        assert result["message"] == "文本框已添加"
        assert result["text_length"] == 11

    def test_add_text_with_color(self, tools_env):
        tools, sid = tools_env
        result = tools.manage_text(
            session_id=sid,
            operation="add",
            slide_index=0,
            text="Red Text",
            left=0, top=0, width=914400, height=914400,
            color="FF0000",
        )
        assert result["operation"] == "add"

    def test_add_text_missing_required(self, tools_env):
        tools, sid = tools_env
        with pytest.raises(ValueError, match="add 操作需要 text"):
            tools.manage_text(
                session_id=sid,
                operation="add",
                slide_index=0,
                left=0, top=0, width=914400, height=914400,
            )

    def test_format_text(self, tools_env):
        tools, sid = tools_env
        # 先添加一个文本框
        tools.add_text(sid, 0, "Format me", position="custom")
        # 获取 shape_index（第一个用户添加的形状）
        slides = tools.list_slides(sid)
        shape_count = slides["slides"][0]["shape_count"]

        result = tools.manage_text(
            session_id=sid,
            operation="format",
            slide_index=0,
            shape_index=shape_count - 1,
            bold=True,
            font_size=32,
            alignment="right",
        )
        assert result["operation"] == "format"
        assert result["message"] == "文本格式已更新"

    def test_format_missing_shape_index(self, tools_env):
        tools, sid = tools_env
        with pytest.raises(ValueError, match="format 操作需要 shape_index"):
            tools.manage_text(
                session_id=sid,
                operation="format",
                slide_index=0,
            )

    def test_extract_all(self, tools_env):
        tools, sid = tools_env
        tools.add_text(sid, 0, "Slide 0 text", position="custom")
        tools.add_slide(sid, layout_index=0)
        tools.add_text(sid, 1, "Slide 1 text", position="custom")

        result = tools.manage_text(
            session_id=sid,
            operation="extract",
        )
        assert result["operation"] == "extract"
        assert result["total_slides"] == 2
        # 每张幻灯片至少有我们添加的文本
        all_text = " ".join(s["text"] for s in result["slides"])
        assert "Slide 0 text" in all_text
        assert "Slide 1 text" in all_text

    def test_extract_single_slide(self, tools_env):
        tools, sid = tools_env
        tools.add_text(sid, 0, "Only this", position="custom")

        result = tools.manage_text(
            session_id=sid,
            operation="extract",
            slide_index=0,
        )
        assert result["total_slides"] == 1
        assert "Only this" in result["slides"][0]["text"]

    def test_invalid_operation(self, tools_env):
        tools, sid = tools_env
        with pytest.raises(ValueError, match="无效的 operation"):
            tools.manage_text(
                session_id=sid,
                operation="delete",
            )


# ===== pptx_format_table_cell =====

class TestFormatTableCell:
    @pytest.fixture
    def table_env(self, tools_env):
        """创建带表格的幻灯片。"""
        tools, sid = tools_env
        tools.add_table(
            session_id=sid,
            slide_index=0,
            rows=3,
            cols=3,
            data=[["A1", "B1", "C1"], ["A2", "B2", "C2"], ["A3", "B3", "C3"]],
        )
        # 获取表格 shape_index
        slides = tools.list_slides(sid)
        shape_index = slides["slides"][0]["shape_count"] - 1
        return tools, sid, shape_index

    def test_set_text(self, table_env):
        tools, sid, si = table_env
        result = tools.format_table_cell(
            session_id=sid,
            slide_index=0,
            shape_index=si,
            row=0,
            col=0,
            text="New Text",
        )
        assert result["message"] == "表格单元格已格式化"
        assert result["row"] == 0
        assert result["col"] == 0

    def test_set_bold_and_font_size(self, table_env):
        tools, sid, si = table_env
        result = tools.format_table_cell(
            session_id=sid,
            slide_index=0,
            shape_index=si,
            row=1,
            col=1,
            bold=True,
            font_size=24,
        )
        assert result["message"] == "表格单元格已格式化"

    def test_set_fill_color(self, table_env):
        tools, sid, si = table_env
        result = tools.format_table_cell(
            session_id=sid,
            slide_index=0,
            shape_index=si,
            row=0,
            col=0,
            fill_color="00FF00",
        )
        assert result["message"] == "表格单元格已格式化"

    def test_set_alignment(self, table_env):
        tools, sid, si = table_env
        result = tools.format_table_cell(
            session_id=sid,
            slide_index=0,
            shape_index=si,
            row=0,
            col=0,
            alignment="center",
        )
        assert result["message"] == "表格单元格已格式化"

    def test_row_out_of_range(self, table_env):
        tools, sid, si = table_env
        with pytest.raises(ValueError, match="row.*超出范围"):
            tools.format_table_cell(
                session_id=sid,
                slide_index=0,
                shape_index=si,
                row=99,
                col=0,
            )

    def test_col_out_of_range(self, table_env):
        tools, sid, si = table_env
        with pytest.raises(ValueError, match="col.*超出范围"):
            tools.format_table_cell(
                session_id=sid,
                slide_index=0,
                shape_index=si,
                row=0,
                col=99,
            )

    def test_not_a_table(self, tools_env):
        tools, sid = tools_env
        # 添加一个文本框而非表格
        tools.add_text(sid, 0, "not a table", position="custom")
        slides = tools.list_slides(sid)
        shape_index = slides["slides"][0]["shape_count"] - 1

        with pytest.raises(ValueError, match="不是表格"):
            tools.format_table_cell(
                session_id=sid,
                slide_index=0,
                shape_index=shape_index,
                row=0,
                col=0,
            )

    def test_invalid_fill_color(self, table_env):
        tools, sid, si = table_env
        with pytest.raises(ValueError, match="无效的颜色格式"):
            tools.format_table_cell(
                session_id=sid,
                slide_index=0,
                shape_index=si,
                row=0,
                col=0,
                fill_color="GGGGGG",
            )


# ===== _parse_hex_color 静态方法测试 =====

class TestParseHexColor:
    def test_valid_colors(self):
        rgb = PptxTools._parse_hex_color("FF0000")
        assert str(rgb) == "FF0000"

        rgb = PptxTools._parse_hex_color("00ff00")
        assert str(rgb) == "00FF00"

    def test_with_hash_prefix(self):
        rgb = PptxTools._parse_hex_color("#0000FF")
        assert str(rgb) == "0000FF"

    def test_invalid_length(self):
        with pytest.raises(ValueError):
            PptxTools._parse_hex_color("FFF")

    def test_invalid_chars(self):
        with pytest.raises(ValueError):
            PptxTools._parse_hex_color("ZZZZZZ")

    def test_non_string(self):
        with pytest.raises(ValueError):
            PptxTools._parse_hex_color(123)
