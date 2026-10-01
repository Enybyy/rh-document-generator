"""Local application. Run with python app.py, or waitress-serve app:app."""
from io import BytesIO
from pathlib import Path
import json
import os

from flask import Flask, jsonify, request, send_file, send_from_directory
from werkzeug.exceptions import RequestEntityTooLarge

from generator import InputError, prepare, build_zip
import samples
from receipts import extract_receipt

ROOT = Path(__file__).resolve().parent
app = Flask(__name__, static_folder=None)
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024


@app.get("/")
def index():
    return send_from_directory(ROOT, "index.html")


@app.get("/assets/<path:name>")
def assets(name):
    return send_from_directory(ROOT / "assets", name)


@app.get("/api/health")
def health():
    return jsonify(mode="local", max_rows=500)


def input_files():
    result = []
    for name, extension, required in [("excel", "xlsx", True), ("template", "docx", True), ("personal", "xlsx", False)]:
        f = request.files.get(name)
        if not f or not f.filename:
            if required:
                raise InputError("Carga un Excel XLSX y una plantilla Word DOCX.")
            result.append(None)
            continue
        if not f.filename.lower().endswith("." + extension):
            raise InputError(f"{name}: usa un archivo .{extension}.")
        result.append(f.read())
    return result


@app.post("/api/preview")
def preview():
    excel, template, personal = input_files()
    return jsonify(prepare(excel, template, personal))


@app.post("/api/receipt")
def receipt():
    file = request.files.get("pdf")
    if not file:
        raise InputError("Selecciona un PDF de recibo.")
    return jsonify(extract_receipt(file.read()))


@app.post("/api/generate")
def generate():
    excel, template, personal = input_files()
    prepared = prepare(excel, template, personal)
    try:
        selected = json.loads(request.form.get("selected", "[]"))
        if not isinstance(selected, list) or any(type(i) is not int or i < 0 or i >= len(prepared["rows"]) for i in selected):
            raise ValueError
    except (ValueError, TypeError):
        raise InputError("La selección de registros no es válida.")
    data = build_zip(template, prepared, set(selected))
    return send_file(BytesIO(data), mimetype="application/zip", as_attachment=True, download_name="documentos-rh.zip")


@app.get("/api/samples/<name>")
def sample(name):
    if name == "personal.xlsx":
        data, mime = samples.excel(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    elif name == "plantilla.docx":
        data, mime = samples.template(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    else:
        return jsonify(error="Ejemplo no encontrado."), 404
    return send_file(BytesIO(data), mimetype=mime, as_attachment=True, download_name=name)


@app.errorhandler(InputError)
def input_error(error):
    return jsonify(error=str(error)), 400


@app.errorhandler(RequestEntityTooLarge)
def too_large(error):
    return jsonify(error="La carga total excede 25 MB. Cada archivo puede pesar hasta 10 MB."), 413


@app.after_request
def headers(response):
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["Referrer-Policy"] = "no-referrer"
    response.headers["Content-Security-Policy"] = "default-src 'self'; style-src 'self'; script-src 'self'; img-src 'self' data:; connect-src 'self'; frame-ancestors 'none'; base-uri 'self'"
    if request.path.startswith("/api/"):
        response.headers["Cache-Control"] = "no-store"
    return response


if __name__ == "__main__":
    from waitress import serve
    serve(app, host="127.0.0.1", port=int(os.environ.get("PORT", "5081")))
