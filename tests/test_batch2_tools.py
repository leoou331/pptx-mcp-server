"""
Batch2 新增工具测试：
- pptx_manage_hyperlinks
- pptx_add_connector
- pptx_manage_slide_transitions
- pptx_set_core_properties
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


# ===== pptx_manage_hyperlinks =====

class TestManageHyperlinks:
    @pytest.fixture
    def text_env(self, tools_env):
        """创建带文本框的幻灯片。"""
        tools, sid = tools_env
        tools.add_text(sid, 0, "Click here for more info", position="custom")
        slides = tools.list_slides(sid)
        shape_index = slides["slides"][0]["shape_count"] - 1
        return tools, sid, shape_index

    def test_add_hyperlink_to_all_runs(self, text_env):
        tools, sid, si = text_env
        result = tools.manage_hyperlinks(
            session_id=sid,
            slide_index=0,
            shape_index=si,
            operation="add",
            url="https://example.com",
        )
        assert result["operation"] == "add"
        assert result["url"] == "https://example.com"
        assert result["added_count"] >= 1
        assert "已为" in result["message"]

    def test_add_hyperlink_with_text(self, text_env):
        tools, sid, si = text_env
        result = tools.manage_hyperlinks(
            session_id=sid,
            slide_index=0,
            shape_index=si,
            operation="add",
            url="https://example.com",
            text="Visit Example",
        )
        assert result["added_count"] == 1

    def test_list_hyperlinks(self, text_env):
        tools, sid, si = text_env
        # 先添加超链接
        tools.manage_hyperlinks(
            session_id=sid, slide_index=0, shape_index=si,
            operation="add", url="https://example.com", text="Link",
        )
        result = tools.manage_hyperlinks(
            session_id=sid, slide_index=0, shape_index=si,
            operation="list",
        )
        assert result["operation"] == "list"
        assert result["count"] >= 1
        assert any(h["url"] == "https://example.com" for h in result["hyperlinks"])

    def test_remove_hyperlinks(self, text_env):
        tools, sid, si = text_env
        # 先添加
        tools.manage_hyperlinks(
            session_id=sid, slide_index=0, shape_index=si,
            operation="add", url="https://example.com",
        )
        # 再移除
        result = tools.manage_hyperlinks(
            session_id=sid, slide_index=0, shape_index=si,
            operation="remove",
        )
        assert result["operation"] == "remove"
        assert "已移除" in result["message"]

    def test_update_hyperlinks(self, text_env):
        tools, sid, si = text_env
        # 先添加
        tools.manage_hyperlinks(
            session_id=sid, slide_index=0, shape_index=si,
            operation="add", url="https://old.com",
        )
        # 再更新
        result = tools.manage_hyperlinks(
            session_id=sid, slide_index=0, shape_index=si,
            operation="update", url="https://new.com",
        )
        assert result["operation"] == "update"
        assert result["url"] == "https://new.com"

    def test_invalid_operation(self, text_env):
        tools, sid, si = text_env
        with pytest.raises(ValueError, match="无效的 operation"):
            tools.manage_hyperlinks(
                session_id=sid, slide_index=0, shape_index=si,
                operation="delete",
            )

    def test_add_missing_url(self, text_env):
        tools, sid, si = text_env
        with pytest.raises(ValueError, match="需要提供 url"):
            tools.manage_hyperlinks(
                session_id=sid, slide_index=0, shape_index=si,
                operation="add",
            )

    def test_invalid_session_id(self, tools_env):
        tools, _ = tools_env
        with pytest.raises(KeyError):
            tools.manage_hyperlinks(
                session_id="nonexistent",
                slide_index=0,
                shape_index=0,
                operation="list",
            )

    def test_shape_without_text_frame(self, tools_env):
        tools, sid = tools_env
        # 添加表格（没有 text_frame）
        tools.add_table(sid, 0, rows=2, cols=2)
        slides = tools.list_slides(sid)
        shape_index = slides["slides"][0]["shape_count"] - 1
        with pytest.raises(ValueError, match="没有文本框"):
            tools.manage_hyperlinks(
                session_id=sid, slide_index=0, shape_index=shape_index,
                operation="list",
            )


# ===== pptx_add_connector =====

class TestAddConnector:
    def test_add_basic_connector(self, tools_env):
        tools, sid = tools_env
        result = tools.add_connector(
            session_id=sid,
            slide_index=0,
            start_x=2.0,
            start_y=3.0,
            end_x=10.0,
            end_y=8.0,
        )
        assert result["message"] == "连接线已添加"
        assert result["slide_index"] == 0
        assert result["arrow_end"] is True

    def test_add_connector_with_color_and_width(self, tools_env):
        tools, sid = tools_env
        result = tools.add_connector(
            session_id=sid,
            slide_index=0,
            start_x=1.0,
            start_y=1.0,
            end_x=5.0,
            end_y=5.0,
            line_color="FF0000",
            line_width=2.5,
        )
        assert result["message"] == "连接线已添加"

    def test_add_connector_double_arrow(self, tools_env):
        tools, sid = tools_env
        result = tools.add_connector(
            session_id=sid,
            slide_index=0,
            start_x=0.0,
            start_y=0.0,
            end_x=10.0,
            end_y=0.0,
            arrow_start=True,
            arrow_end=True,
        )
        assert result["arrow_start"] is True
        assert result["arrow_end"] is True

    def test_add_connector_no_arrows(self, tools_env):
        tools, sid = tools_env
        result = tools.add_connector(
            session_id=sid,
            slide_index=0,
            start_x=1.0,
            start_y=1.0,
            end_x=5.0,
            end_y=5.0,
            arrow_start=False,
            arrow_end=False,
        )
        assert result["arrow_start"] is False
        assert result["arrow_end"] is False

    def test_add_connector_invalid_color(self, tools_env):
        tools, sid = tools_env
        with pytest.raises(ValueError, match="无效的颜色格式"):
            tools.add_connector(
                session_id=sid,
                slide_index=0,
                start_x=1.0, start_y=1.0,
                end_x=5.0, end_y=5.0,
                line_color="ZZZZZZ",
            )

    def test_add_connector_invalid_width(self, tools_env):
        tools, sid = tools_env
        with pytest.raises(ValueError, match="line_width 必须大于 0"):
            tools.add_connector(
                session_id=sid,
                slide_index=0,
                start_x=1.0, start_y=1.0,
                end_x=5.0, end_y=5.0,
                line_width=-1,
            )

    def test_add_connector_non_numeric_coord(self, tools_env):
        tools, sid = tools_env
        with pytest.raises(TypeError, match="start_x 必须是数字"):
            tools.add_connector(
                session_id=sid,
                slide_index=0,
                start_x="abc", start_y=1.0,
                end_x=5.0, end_y=5.0,
            )

    def test_invalid_session_id(self, tools_env):
        tools, _ = tools_env
        with pytest.raises(KeyError):
            tools.add_connector(
                session_id="nonexistent",
                slide_index=0,
                start_x=1.0, start_y=1.0,
                end_x=5.0, end_y=5.0,
            )


# ===== pptx_manage_slide_transitions =====

class TestManageSlideTransitions:
    def test_set_fade_transition(self, tools_env):
        tools, sid = tools_env
        result = tools.manage_slide_transitions(
            session_id=sid,
            slide_index=0,
            transition_type="fade",
        )
        assert result["message"] == "幻灯片过渡效果已设置"
        assert result["transition_type"] == "fade"

    def test_set_transition_with_duration(self, tools_env):
        tools, sid = tools_env
        result = tools.manage_slide_transitions(
            session_id=sid,
            slide_index=0,
            transition_type="wipe",
            duration=1.5,
        )
        assert result["duration_seconds"] == 1.5

    def test_set_transition_with_advance(self, tools_env):
        tools, sid = tools_env
        result = tools.manage_slide_transitions(
            session_id=sid,
            slide_index=0,
            transition_type="push",
            advance_after=5.0,
        )
        assert result["advance_after_seconds"] == 5.0

    def test_set_transition_with_all_options(self, tools_env):
        tools, sid = tools_env
        result = tools.manage_slide_transitions(
            session_id=sid,
            slide_index=0,
            transition_type="dissolve",
            duration=2.0,
            advance_after=10.0,
        )
        assert result["transition_type"] == "dissolve"
        assert result["duration_seconds"] == 2.0
        assert result["advance_after_seconds"] == 10.0

    def test_all_transition_types(self, tools_env):
        tools, sid = tools_env
        types = ["fade", "push", "wipe", "split", "zoom", "fly", "appear",
                 "dissolve", "cut", "wheel", "strips", "checker", "blinds",
                 "box", "random"]
        for tt in types:
            result = tools.manage_slide_transitions(
                session_id=sid,
                slide_index=0,
                transition_type=tt,
            )
            assert result["transition_type"] == tt

    def test_replace_existing_transition(self, tools_env):
        tools, sid = tools_env
        # 设置第一次
        tools.manage_slide_transitions(
            session_id=sid, slide_index=0, transition_type="fade",
        )
        # 替换为新的
        result = tools.manage_slide_transitions(
            session_id=sid, slide_index=0, transition_type="wipe",
        )
        assert result["transition_type"] == "wipe"

    def test_invalid_transition_type(self, tools_env):
        tools, sid = tools_env
        with pytest.raises(ValueError, match="不支持的 transition_type"):
            tools.manage_slide_transitions(
                session_id=sid,
                slide_index=0,
                transition_type="nonexistent",
            )

    def test_negative_duration(self, tools_env):
        tools, sid = tools_env
        with pytest.raises(ValueError, match="duration 不能为负数"):
            tools.manage_slide_transitions(
                session_id=sid,
                slide_index=0,
                transition_type="fade",
                duration=-1.0,
            )

    def test_invalid_session_id(self, tools_env):
        tools, _ = tools_env
        with pytest.raises(KeyError):
            tools.manage_slide_transitions(
                session_id="nonexistent",
                slide_index=0,
                transition_type="fade",
            )


# ===== pptx_set_core_properties =====

class TestSetCoreProperties:
    def test_set_title(self, tools_env):
        tools, sid = tools_env
        result = tools.set_core_properties(
            session_id=sid,
            title="My Presentation",
        )
        assert result["field_count"] == 1
        assert result["updated_fields"]["title"] == "My Presentation"
        assert "已更新" in result["message"]

    def test_set_multiple_properties(self, tools_env):
        tools, sid = tools_env
        result = tools.set_core_properties(
            session_id=sid,
            title="Test Title",
            subject="Test Subject",
            author="Test Author",
            keywords="test, pptx",
            comments="This is a test",
            category="Testing",
        )
        assert result["field_count"] == 6
        assert result["updated_fields"]["title"] == "Test Title"
        assert result["updated_fields"]["author"] == "Test Author"

    def test_set_author_only(self, tools_env):
        tools, sid = tools_env
        result = tools.set_core_properties(
            session_id=sid,
            author="John Doe",
        )
        assert result["field_count"] == 1
        assert result["updated_fields"]["author"] == "John Doe"

    def test_properties_persist(self, tools_env):
        tools, sid = tools_env
        tools.set_core_properties(
            session_id=sid,
            title="Persistent Title",
            author="Persistent Author",
        )
        # 验证属性确实被设置了
        session = tools.sessions.get(sid)
        with session.lock:
            props = session.presentation.core_properties
            assert props.title == "Persistent Title"
            assert props.author == "Persistent Author"

    def test_no_properties_provided(self, tools_env):
        tools, sid = tools_env
        with pytest.raises(ValueError, match="至少需要设置一个属性"):
            tools.set_core_properties(session_id=sid)

    def test_non_string_title(self, tools_env):
        tools, sid = tools_env
        with pytest.raises(TypeError, match="title 必须是字符串"):
            tools.set_core_properties(
                session_id=sid,
                title=123,
            )

    def test_invalid_session_id(self, tools_env):
        tools, _ = tools_env
        with pytest.raises(KeyError):
            tools.set_core_properties(
                session_id="nonexistent",
                title="Test",
            )
