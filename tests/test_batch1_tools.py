"""批次1新增工具测试：add_shape, add_chart, manage_text, format_table_cell"""
import pytest

from security.session import SessionManager
from tools.manager import PptxTools


@pytest.fixture
def setup_tools(tmp_path):
    manager = SessionManager()
    tools = PptxTools(manager, str(tmp_path))
    session_id = tools.create("Test Deck")["session_id"]
    tools.add_slide(session_id, layout_index=6)
    return tools, session_id


class TestAddShape:
    def test_add_rectangle(self, setup_tools):
        tools, sid = setup_tools
        result = tools.add_shape(sid, 0, "rectangle", 1.0, 1.0, 3.0, 2.0)
        assert result["message"] == "形状已添加"
        assert result["shape_type"] == "rectangle"

    def test_add_with_text_and_color(self, setup_tools):
        tools, sid = setup_tools
        result = tools.add_shape(
            sid,
            0,
            "oval",
            1.0,
            1.0,
            3.0,
            2.0,
            text="Hello",
            fill_color="#FF5733",
            line_color="#000000",
        )
        assert result["has_text"] is True

    def test_invalid_shape_type(self, setup_tools):
        tools, sid = setup_tools
        with pytest.raises(ValueError, match="不支持的形状类型"):
            tools.add_shape(sid, 0, "invalid_shape", 1.0, 1.0, 3.0, 2.0)

    def test_invalid_color(self, setup_tools):
        tools, sid = setup_tools
        with pytest.raises(ValueError, match="颜色"):
            tools.add_shape(sid, 0, "rectangle", 1.0, 1.0, 3.0, 2.0, fill_color="#XYZ")

    def test_arrow_alias(self, setup_tools):
        tools, sid = setup_tools
        result = tools.add_shape(sid, 0, "arrow", 0.5, 0.5, 2.0, 1.0)
        assert result["message"] == "形状已添加"


class TestAddChart:
    def test_add_column_chart(self, setup_tools):
        tools, sid = setup_tools
        result = tools.add_chart(
            sid,
            0,
            "column",
            categories=["Q1", "Q2", "Q3"],
            series_data=[{"name": "Sales", "values": [100, 200, 150]}],
            title="季度销售",
        )
        assert result["message"] == "图表已添加"
        assert result["has_title"] is True

    def test_invalid_chart_type(self, setup_tools):
        tools, sid = setup_tools
        with pytest.raises(ValueError, match="不支持的图表类型"):
            tools.add_chart(sid, 0, "donut", ["A"], [{"name": "s", "values": [1]}])

    def test_values_length_mismatch(self, setup_tools):
        tools, sid = setup_tools
        with pytest.raises(ValueError, match="长度"):
            tools.add_chart(
                sid,
                0,
                "line",
                categories=["A", "B", "C"],
                series_data=[{"name": "S", "values": [1, 2]}],
            )

    def test_empty_categories(self, setup_tools):
        tools, sid = setup_tools
        with pytest.raises(ValueError, match="categories"):
            tools.add_chart(sid, 0, "bar", [], [{"name": "S", "values": []}])


class TestManageText:
    def test_add_textbox(self, setup_tools):
        tools, sid = setup_tools
        result = tools.manage_text(sid, "add", slide_index=0, text="Hello World", font_size=24)
        assert result["operation"] == "add"
        assert result["message"] == "文本框已添加"

    def test_extract_all(self, setup_tools):
        tools, sid = setup_tools
        result = tools.manage_text(sid, "extract")
        assert result["operation"] == "extract"
        assert "total_slides" in result

    def test_add_missing_text(self, setup_tools):
        tools, sid = setup_tools
        with pytest.raises(ValueError, match="text"):
            tools.manage_text(sid, "add", slide_index=0)

    def test_format_missing_shape_index(self, setup_tools):
        tools, sid = setup_tools
        with pytest.raises(ValueError, match="shape_index"):
            tools.manage_text(sid, "format", slide_index=0)

    def test_invalid_operation(self, setup_tools):
        tools, sid = setup_tools
        with pytest.raises(ValueError, match="不支持的操作"):
            tools.manage_text(sid, "delete", slide_index=0)


class TestFormatTableCell:
    def test_format_cell(self, setup_tools):
        tools, sid = setup_tools
        tools.add_table(
            sid,
            0,
            rows=3,
            cols=3,
            data=[["A", "B", "C"], ["D", "E", "F"], ["G", "H", "I"]],
        )
        result = tools.format_table_cell(
            sid,
            0,
            0,
            row=0,
            col=0,
            text="Updated",
            font_size=14,
            bold=True,
            fill_color="#FFFF00",
            alignment="center",
        )
        assert result["message"] == "表格单元格已格式化"

    def test_shape_not_table(self, setup_tools):
        tools, sid = setup_tools
        tools.manage_text(sid, "add", slide_index=0, text="not a table")
        with pytest.raises(ValueError, match="不是表格"):
            tools.format_table_cell(sid, 0, 0, row=0, col=0, text="x")

    def test_row_out_of_range(self, setup_tools):
        tools, sid = setup_tools
        tools.add_table(sid, 0, rows=2, cols=2)
        with pytest.raises(ValueError, match="row"):
            tools.format_table_cell(sid, 0, 0, row=99, col=0)
