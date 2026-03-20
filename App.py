from flask import Flask, request, jsonify, render_template, send_file
from Extractor import process_folder
from Exporter import export_to_excel
import os

# templates folder is lowercase: templates/index.html
app = Flask(__name__, template_folder="templates")

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/extract", methods=["POST"])
def extract():
    data        = request.json
    folder_path = data.get("folder_path", "").strip()
    fields      = data.get("fields", [])

    if not folder_path:
        return jsonify({"error": "No folder path provided"}), 400
    if not fields:
        return jsonify({"error": "No fields provided"}), 400

    # Allow relative paths (relative to App.py location)
    if not os.path.isabs(folder_path):
        base = os.path.dirname(os.path.abspath(__file__))
        folder_path = os.path.join(base, folder_path)

    if not os.path.exists(folder_path):
        return jsonify({"error": f"Folder not found: {folder_path}"}), 404

    results = process_folder(folder_path, fields)
    return jsonify(results)

@app.route("/export", methods=["POST"])
def export():
    data   = request.json
    rows   = data.get("data", [])
    fields = data.get("fields", [])

    # Save next to App.py
    out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output.xlsx")
    path = export_to_excel(rows, fields, output_path=out_path)
    return send_file(path, as_attachment=True, download_name="extracted_data.xlsx")

if __name__ == "__main__":
    print("\n  ID-RSS — HackIndia 2026")
    print("  Running at http://localhost:5000\n")
    app.run(host="0.0.0.0", port=5000, debug=True)
