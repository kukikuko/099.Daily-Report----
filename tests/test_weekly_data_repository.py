import json
import os

from repositories.weekly_data_repository import (
    normalize_weekly_data,
    save_json_atomically,
)


def test_normalize_weekly_data():
    raw = {
        "locations": ["본사", "연구소"],
        "holiday_indices": [1, 3, 3, 9, "bad"],
    }
    res = normalize_weekly_data(raw, default_loc="본사")
    assert len(res["locations"]) == 5
    assert res["locations"][0] == "본사"
    assert res["locations"][1] == "연구소"
    assert res["locations"][2] == "본사"
    assert res["holiday_indices"] == [1, 3]


def test_save_json_atomically(tmp_path):
    dest = os.path.join(tmp_path, "sub", "test.json")
    data = {"key": "value"}

    save_json_atomically(data, dest)

    assert os.path.isfile(dest)
    with open(dest, "r", encoding="utf-8") as f:
        loaded = json.load(f)
    assert loaded == data
