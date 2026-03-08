import threading
from pathlib import Path
from tempfile import TemporaryDirectory

import pytest
from PIL import Image

from security.session import SessionManager
from tools.manager import PptxTools


class DummyPresentation:
    def __init__(self):
        self.slides = []

    def save(self, path):
        Path(path).write_bytes(b"replacement")


class DummySession:
    def __init__(self):
        self.presentation = DummyPresentation()
        self.lock = threading.Lock()
        self.dirty = True


class DummySessions:
    def __init__(self, session):
        self._session = session

    def get(self, session_id):
        assert session_id == "session-1"
        return self._session


def test_save_is_atomic_when_presentation_save_fails(tmp_path):
    target = tmp_path / "deck.pptx"
    target.write_bytes(b"original")

    session = DummySession()
    tools = PptxTools(DummySessions(session), str(tmp_path))

    def broken_save(path):
        Path(path).write_bytes(b"partial")
        raise RuntimeError("save failed")

    session.presentation.save = broken_save

    with pytest.raises(RuntimeError, match="save failed"):
        tools.save("session-1", str(target))

    assert target.read_bytes() == b"original"


def test_add_image_accepts_height_only_dimension(tmp_path):
    manager = SessionManager()
    tools = PptxTools(manager, str(tmp_path))
    session_id = tools.create("Deck")["session_id"]
    tools.add_slide(session_id, layout_index=0)

    image_path = tmp_path / "sample.png"
    Image.new("RGB", (32, 32), color="white").save(image_path)

    result = tools.add_image(
        session_id=session_id,
        slide_index=0,
        image_path=str(image_path),
        left=1.0,
        top=1.0,
        height=2.0,
    )

    assert result["slide_index"] == 0
    assert result["message"] == "图片已添加"


def test_open_rejects_absolute_paths_outside_allowed_directories(tmp_path):
    manager = SessionManager()
    tools = PptxTools(manager, str(tmp_path))

    with TemporaryDirectory() as outside_dir:
        outside_path = Path(outside_dir) / "external.pptx"

        with pytest.raises(ValueError, match="路径不在允许目录内"):
            tools.open(str(outside_path))
