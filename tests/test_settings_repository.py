import os

from repositories.settings_repository import load_env_dict


def test_load_env_dict(tmp_path):
    env_file = os.path.join(tmp_path, ".env")
    with open(env_file, "w", encoding="utf-8") as f:
        f.write("# comment\n")
        f.write("KEY1=VAL1\n")
        f.write("KEY2 = VAL2 \n")

    res = load_env_dict(env_file)
    assert res == {"KEY1": "VAL1", "KEY2": "VAL2"}


def test_load_env_dict_nonexistent():
    res = load_env_dict("nonexistent_path_.env")
    assert res == {}
