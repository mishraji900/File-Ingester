from flask import Flask, request, jsonify
import os
from converter import convert_to_xlsx
import openpyxl

app = Flask(__name__)

OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "converted")
os.makedirs(OUTPUT_DIR, exist_ok=True)


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


@app.route("/convert", methods=["POST"])
def convert():
    data = request.get_json()
    file_paths = data.get("files", [])
    results = []

    for fp in file_paths:
        try:
            converted_path = convert_to_xlsx(fp, OUTPUT_DIR)
            results.append({
                "original": os.path.basename(fp),
                "converted": converted_path,
                "status": "ok"
            })
        except Exception as e:
            results.append({
                "original": os.path.basename(fp),
                "converted": None,
                "status": "error",
                "error": str(e)
            })

    return jsonify({"files": results})


@app.route("/sheets", methods=["POST"])
def get_sheets():
    data = request.get_json()
    file_path = data.get("file")

    try:
        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
        sheet_names = wb.sheetnames
        wb.close()
        return jsonify({"sheets": sheet_names})
    except Exception as e:
        return jsonify({"error": str(e)}), 400


@app.route("/preview", methods=["POST"])
def preview():
    data = request.get_json()
    file_path = data.get("file")
    sheet_name = data.get("sheet")
    max_rows = data.get("max_rows", 20)

    try:
        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
        ws = wb[sheet_name]

        rows_iter = ws.iter_rows(values_only=True)
        header = next(rows_iter, None)
        if header is None:
            wb.close()
            return jsonify({"columns": [], "rows": []})

        columns = [str(c) if c is not None else f"col_{i}" for i, c in enumerate(header)]

        preview_rows = []
        for i, row in enumerate(rows_iter):
            if i >= max_rows:
                break
            row_dict = {}
            for col_name, val in zip(columns, row):
                row_dict[col_name] = val
            preview_rows.append(row_dict)

        wb.close()
        return jsonify({"columns": columns, "rows": preview_rows})
    except Exception as e:
        return jsonify({"error": str(e)}), 400


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5123)