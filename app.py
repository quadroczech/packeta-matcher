from flask import Flask, request, send_file, jsonify, render_template
from io import BytesIO
import os
from excel_processor import process_excel

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024  # max 10 MB

API_PASSWORD = os.environ.get("ZASILKOVNA_API_PASSWORD", "")

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/process", methods=["POST"])
def process():
    if "file" not in request.files:
        return jsonify({"error": "Soubor nebyl nahrán."}), 400

    file = request.files["file"]
    if not file.filename.endswith((".xlsx", ".xls")):
        return jsonify({"error": "Prosím nahrajte soubor ve formátu .xlsx"}), 400

    save_folder = request.form.get("save_folder", "").strip()
    file_bytes = file.read()

    try:
        result_bytes, stats = process_excel(file_bytes, API_PASSWORD)
    except ValueError as e:
        return jsonify({"error": str(e)}), 422
    except Exception as e:
        return jsonify({"error": f"Chyba při zpracování: {str(e)}"}), 500

    output_filename = file.filename.rsplit(".", 1)[0] + "_objednavky.xlsx"
    saved_path = ""

    if save_folder and os.path.isdir(save_folder):
        saved_path = os.path.join(save_folder, output_filename)
        with open(saved_path, "wb") as f:
            f.write(result_bytes)

    response = send_file(
        BytesIO(result_bytes),
        as_attachment=True,
        download_name=output_filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response.headers["X-Stats-Total"] = stats["total"]
    response.headers["X-Stats-Found"] = stats["found"]
    response.headers["X-Stats-NotFound"] = stats["not_found"]
    response.headers["X-Saved-Path"] = saved_path
    response.headers["Access-Control-Expose-Headers"] = (
        "X-Stats-Total, X-Stats-Found, X-Stats-NotFound, X-Saved-Path"
    )
    return response

if __name__ == "__main__":
    app.run(debug=True, port=5000)
