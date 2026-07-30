import os

from utils.file_utils import ensure_dir_writable, is_file_locked, save_env_dict_atomically


def test_ensure_dir_writable(tmp_path):
    target_dir = os.path.join(tmp_path, "writable_dir")
    assert ensure_dir_writable(target_dir) is True
    assert os.path.exists(target_dir)


def test_is_file_locked_nonexistent():
    assert is_file_locked("nonexistent_file_path_12345.txt") is False


def test_save_env_dict_atomically(tmp_path):
    env_file = os.path.join(tmp_path, ".env")
    data = {"KEY": "VAL"}
    save_env_dict_atomically(data, env_file)

    assert os.path.exists(env_file)
    with open(env_file, "r", encoding="utf-8") as f:
        content = f.read()
    assert "KEY=VAL" in content
