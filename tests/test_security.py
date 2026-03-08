import os
from pathlib import Path

import pytest

from security.tempfile import temp_manager
from security.validator import safe_path_in_dirs


def test_safe_path_in_dirs_allows_nested_relative_paths(tmp_path):
    resolved = Path(safe_path_in_dirs(str(tmp_path), "nested/deck.pptx"))

    assert resolved == (tmp_path / "nested" / "deck.pptx").resolve()


def test_safe_path_in_dirs_rejects_absolute_path_outside_allowlist(tmp_path):
    outside_dir = tmp_path.parent / "outside"
    outside_dir.mkdir(exist_ok=True)
    outside_path = outside_dir / "secret.pptx"

    with pytest.raises(ValueError, match="路径不在允许目录内"):
        safe_path_in_dirs(str(tmp_path), str(outside_path), temp_manager.temp_dir)


@pytest.mark.skipif(not hasattr(os, "symlink"), reason="symlink not supported on this platform")
def test_safe_path_in_dirs_rejects_symlink_escape(tmp_path):
    outside_dir = tmp_path.parent / "symlink-outside"
    outside_dir.mkdir(exist_ok=True)
    linked_dir = tmp_path / "linked"
    os.symlink(outside_dir, linked_dir, target_is_directory=True)

    with pytest.raises(ValueError, match="路径不在允许目录内"):
        safe_path_in_dirs(str(tmp_path), "linked/secret.pptx")
