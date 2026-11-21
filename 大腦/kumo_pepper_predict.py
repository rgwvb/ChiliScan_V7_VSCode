"""Use NVIDIA Kumo Tabular to classify pepper spectra.

Put reference.csv and new_samples.csv in the same folder as this script.
Both CSV files must have the same sample_id and wavelength-feature columns.
reference.csv also needs a quality column with labels: 新鮮, 蟲害, or 腐爛.
If there are more than 498 wavelength features, the request evenly samples 498
to stay within Kumo's 500-column limit (including sample_id and quality).

From the project root, run these commands in PowerShell:
    cd .\\大腦
    python .\\kumo_pepper_predict.py

The script asks for the rotated NVIDIA API key using hidden terminal input.
"""

import csv
import getpass
import json
import math
import sys
import urllib.error
import urllib.request
from pathlib import Path

API_URL = "https://ai.api.nvidia.com/v1/structured-data/nvidia/kumo-tabular/predictions"
ID_COLUMN = "sample_id"
TARGET_COLUMN = "quality"
CLASSES = ["新鮮", "蟲害", "腐爛"]
BASE_DIR = Path(__file__).resolve().parent
REFERENCE_FILE = BASE_DIR / "reference.csv"
PREDICT_FILE = BASE_DIR / "new_samples.csv"
MAX_RESPONSE_BYTES = 2_000_000
MAX_FEATURE_COLUMNS = 498


def read_csv(path: Path):
    if not path.exists():
        raise FileNotFoundError(f"找不到檔案：{path.name}")
    with path.open("r", encoding="utf-8-sig", newline="") as file:
        reader = csv.DictReader(file)
        if not reader.fieldnames:
            raise ValueError(f"{path.name} 沒有欄位標題")
        rows = list(reader)
        return reader.fieldnames, rows


def to_number(row, column, row_number):
    value = (row.get(column) or "").strip()
    if not value:
        raise ValueError(f"{row_number} 列的 {column} 是空白")
    try:
        number = float(value.replace(",", ""))
    except ValueError as error:
        raise ValueError(f"{row_number} 列的 {column} 不是數字：{value}") from error
    if not math.isfinite(number):
        raise ValueError(f"{row_number} 列的 {column} 必須是有限數字")
    return number


class NoRedirectHandler(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, request, response, code, message, headers, new_url):
        return None


def main():
    api_key = getpass.getpass("請輸入已輪替的 NVIDIA API 金鑰（輸入不會顯示）：").strip()
    if not api_key:
        raise EnvironmentError("未輸入 NVIDIA API 金鑰。")

    ref_headers, reference_rows = read_csv(REFERENCE_FILE)
    query_headers, query_rows = read_csv(PREDICT_FILE)

    if not reference_rows:
        raise ValueError("reference.csv 至少要有一筆已標記樣本")
    if not query_rows:
        raise ValueError("new_samples.csv 至少要有一筆待判斷樣本")
    if ID_COLUMN not in ref_headers or TARGET_COLUMN not in ref_headers:
        raise ValueError("reference.csv 必須包含 sample_id 和 quality 欄位")
    if ID_COLUMN not in query_headers:
        raise ValueError("new_samples.csv 必須包含 sample_id 欄位")
    if len(set(ref_headers)) != len(ref_headers) or len(set(query_headers)) != len(query_headers):
        raise ValueError("CSV 欄位名稱不可重複")

    feature_columns = [
        name for name in ref_headers if name not in (ID_COLUMN, TARGET_COLUMN)
    ]
    query_feature_columns = [name for name in query_headers if name != ID_COLUMN]
    if not feature_columns:
        raise ValueError("找不到波長強度欄位")
    if feature_columns != query_feature_columns:
        raise ValueError(
            "reference.csv 和 new_samples.csv 的波長欄位或順序不一致"
        )
    if len(feature_columns) > MAX_FEATURE_COLUMNS:
        original_count = len(feature_columns)
        last_index = original_count - 1
        selected_indices = [
            round(index * last_index / (MAX_FEATURE_COLUMNS - 1))
            for index in range(MAX_FEATURE_COLUMNS)
        ]
        feature_columns = [feature_columns[index] for index in selected_indices]
        print(
            f"Kumo 每表最多 {MAX_FEATURE_COLUMNS + 2} 欄；"
            f"已將 {original_count} 個波長特徵均勻取樣為 {len(feature_columns)} 個。"
        )

    context_rows = []
    context_ids = set()
    context_labels = set()
    for row_number, row in enumerate(reference_rows, start=2):
        sample_id = (row.get(ID_COLUMN) or "").strip()
        label = (row.get(TARGET_COLUMN) or "").strip()
        if not sample_id:
            raise ValueError(f"reference.csv 第 {row_number} 列缺少 sample_id")
        if sample_id in context_ids:
            raise ValueError(f"reference.csv 第 {row_number} 列的 sample_id 重複")
        if label not in CLASSES:
            raise ValueError(
                f"reference.csv 第 {row_number} 列的 quality 必須是："
                + "、".join(CLASSES)
            )
        values = [to_number(row, column, row_number) for column in feature_columns]
        context_rows.append([sample_id, *values, label])
        context_ids.add(sample_id)
        context_labels.add(label)
    if context_labels != set(CLASSES):
        raise ValueError("reference.csv 必須至少包含新鮮、蟲害、腐爛各一筆參考樣本")

    prediction_rows = []
    for row_number, row in enumerate(query_rows, start=2):
        sample_id = (row.get(ID_COLUMN) or "").strip()
        if not sample_id:
            raise ValueError(f"new_samples.csv 第 {row_number} 列缺少 sample_id")
        if sample_id in context_ids:
            raise ValueError(f"new_samples.csv 第 {row_number} 列的 sample_id 與參考樣本重複")
        values = [to_number(row, column, row_number) for column in feature_columns]
        prediction_rows.append([sample_id, *values])
        context_ids.add(sample_id)

    schema_columns = {
        ID_COLUMN: {"dtype": "string", "stype": "ID", "nullable": False}
    }
    for column in feature_columns:
        schema_columns[column] = {"dtype": "float64", "stype": "numerical"}
    schema_columns[TARGET_COLUMN] = {
        "dtype": "string",
        "stype": "categorical",
    }

    payload = {
        "model": "kumo-tabular",
        "task": {
            "kind": "classification",
            "target": {
                "column_name": TARGET_COLUMN,
                "dtype": "string",
                "classes": CLASSES,
            },
        },
        "schema": {
            "instance_table": {
                "columns": schema_columns,
                "primary_key": ID_COLUMN,
            },
            "related_tables": {},
            "relationships": [],
        },
        "context": {
            "instance_table": {
                "format": "arrays",
                "columns": [ID_COLUMN, *feature_columns, TARGET_COLUMN],
                "rows": context_rows,
            },
            "related_tables": {},
        },
        "predict": {
            "instance_table": {
                "format": "arrays",
                "columns": [ID_COLUMN, *feature_columns],
                "rows": prediction_rows,
            },
            "related_tables": {},
        },
        "output": {"fields": ["prediction", "probabilities"]},
    }

    request = urllib.request.Request(
        API_URL,
        data=json.dumps(payload, ensure_ascii=False, allow_nan=False).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        },
        method="POST",
    )

    print(f"傳送 {len(context_rows)} 筆參考樣本，預測 {len(prediction_rows)} 筆資料…")
    opener = urllib.request.build_opener(NoRedirectHandler())
    try:
        with opener.open(request, timeout=120) as response:
            body = response.read(MAX_RESPONSE_BYTES + 1)
            if len(body) > MAX_RESPONSE_BYTES:
                print("NVIDIA API 回傳內容過大。", file=sys.stderr)
                return 1
            result = json.loads(body.decode("utf-8"))
    except urllib.error.HTTPError as error:
        detail = error.read(MAX_RESPONSE_BYTES).decode("utf-8", errors="replace")
        print(f"NVIDIA API 錯誤（HTTP {error.code}）：\n{detail}", file=sys.stderr)
        return 1
    except urllib.error.URLError as error:
        print(f"連線 NVIDIA API 失敗：{error.reason}", file=sys.stderr)
        return 1
    except (UnicodeDecodeError, json.JSONDecodeError):
        print("NVIDIA API 回傳無效 JSON。", file=sys.stderr)
        return 1

    print("\n分類結果：")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (FileNotFoundError, ValueError, EnvironmentError) as error:
        print(f"錯誤：{error}", file=sys.stderr)
        raise SystemExit(1)
