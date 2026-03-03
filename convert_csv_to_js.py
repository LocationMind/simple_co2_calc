"""
CSVファイルをJavaScriptファイルに変換するスクリプト

使い方:
    python convert_csv_to_js.py

説明:
    型式一覧フォルダー内のCSVファイルを読み込み、
    型式と燃費値のペアをJavaScriptオブジェクトとして出力します。

    入力: 型式一覧/output_XXXX.csv
    出力: 型式一覧/output_XXXX.js
"""

import csv
import json
import os

# CSVファイルのリスト
csv_files = [
    "output_WLTC.csv",
    "output_JC08.csv",
    "output_10-15.csv",
    "output_JH25.csv",
    "output_JH15.csv",
]

print("=" * 60)
print("CSV→JavaScript 変換ツール")
print("=" * 60)
print()

for csv_file in csv_files:
    csv_path = os.path.join("型式一覧", csv_file)

    if not os.path.exists(csv_path):
        print(f"[WARN] ファイルが見つかりません: {csv_path}")
        continue

    data = {}

    try:
        with open(csv_path, "r", encoding="utf-8") as f:
            reader = csv.DictReader(f)

            for row in reader:
                model_type = row.get("型式", "").strip()
                fuel_efficiency = row.get("燃費値（km/L）", "").strip()
                fuel_type = row.get("燃料種別", "").strip()

                if model_type and fuel_efficiency:
                    try:
                        efficiency_value = float(fuel_efficiency)

                        # 型式ごとに配列で管理（複数の燃料種別がある場合に対応）
                        if model_type not in data:
                            data[model_type] = []

                        # 既に同じ燃料種別と燃費値のデータがあるかチェック
                        existing = False
                        for item in data[model_type]:
                            if (
                                item["fuel"] == fuel_type
                                and item["efficiency"] == efficiency_value
                            ):
                                existing = True
                                break

                        # 同じ燃料種別でも燃費値が異なる場合は追加する
                        if not existing:
                            data[model_type].append(
                                {"fuel": fuel_type, "efficiency": efficiency_value}
                            )
                    except ValueError:
                        continue

        # JavaScriptファイルとして出力
        js_filename = csv_file.replace(".csv", ".js")
        js_path = os.path.join("型式一覧", js_filename)

        with open(js_path, "w", encoding="utf-8") as f:
            # ファイル名から変数名を生成（ハイフンやドットをアンダースコアに変換）
            var_name = csv_file.replace(".csv", "").replace("-", "_").replace("・", "_")
            f.write(f"const csvData_{var_name} = ")
            f.write(json.dumps(data, ensure_ascii=False, indent=2))
            f.write(";\n")

        print(f"[OK] 変換完了: {csv_file} -> {js_filename}")
        print(f"  データ数: {len(data):,}件")

    except Exception as e:
        print(f"[ERROR] エラー: {csv_file} - {str(e)}")

print()
print("=" * 60)
print("すべての変換が完了しました！")
print("=" * 60)
