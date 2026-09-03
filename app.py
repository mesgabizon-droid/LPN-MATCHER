import json
import os
import shutil
import tempfile
import threading
import uuid

from flask import (
    Flask,
    after_this_request,
    flash,
    redirect,
    render_template,
    request,
    send_file,
    url_for,
)
from werkzeug.utils import secure_filename

from matcher import run_matching

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-key-cambia-esto")
app.config["MAX_CONTENT_LENGTH"] = 40 * 1024 * 1024  # 40 MB

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")
JOBS_DIR = os.path.join(BASE_DIR, "jobs")
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(JOBS_DIR, exist_ok=True)

ALLOWED_EXCEL = {".xls", ".xlsx"}
ALLOWED_PDF = {".pdf"}


def _ext(filename):
    return os.path.splitext(filename)[1].lower()


def _job_path(job_id):
    return os.path.join(JOBS_DIR, f"{job_id}.json")


def _write_job(job_id, data):
    path = _job_path(job_id)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f)
    os.replace(tmp, path)


def _read_job(job_id):
    try:
        with open(_job_path(job_id), "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return None


def _procesar_job(job_id, pedido_path, etiquetas_path, workdir):
    try:
        def progress_cb(pagina, total):
            _write_job(
                job_id,
                {
                    "status": "procesando",
                    "paso": f"Leyendo etiqueta {pagina} de {total} del PDF...",
                    "pagina": pagina,
                    "total": total,
                },
            )

        token = uuid.uuid4().hex[:12]
        out_xlsx_name = f"pedido_LPN_{token}.xlsx"
        out_pdf_name = f"etiquetas_CODIGO_{token}.pdf"
        out_xlsx_path = os.path.join(OUTPUT_DIR, out_xlsx_name)
        out_pdf_path = os.path.join(OUTPUT_DIR, out_pdf_name)

        info = run_matching(
            pedido_path, etiquetas_path, out_xlsx_path, out_pdf_path, workdir,
            progress_cb=progress_cb,
        )

        _write_job(
            job_id,
            {
                "status": "listo",
                "n_pdf": info["n_pdf"],
                "n_excel": info["n_excel"],
                "resumen": info["resumen"],
                "archivo_xlsx": out_xlsx_name,
                "archivo_pdf": out_pdf_name,
            },
        )
    except FileNotFoundError as e:
        _write_job(
            job_id,
            {
                "status": "error",
                "mensaje": f"Falta una herramienta del sistema (poppler-utils o tesseract-ocr). Detalle: {e}",
            },
        )
    except Exception as e:
        _write_job(job_id, {"status": "error", "mensaje": f"Ocurrio un error procesando los archivos: {e}"})
    finally:
        shutil.rmtree(workdir, ignore_errors=True)


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/procesar", methods=["POST"])
def procesar():
    pedido = request.files.get("pedido")
    etiquetas = request.files.get("etiquetas")

    if not pedido or pedido.filename == "":
        flash("Debes subir el archivo Excel del pedido.")
        return redirect(url_for("index"))
    if not etiquetas or etiquetas.filename == "":
        flash("Debes subir el PDF de etiquetas.")
        return redirect(url_for("index"))

    if _ext(pedido.filename) not in ALLOWED_EXCEL:
        flash("El pedido debe ser un archivo .xls o .xlsx")
        return redirect(url_for("index"))
    if _ext(etiquetas.filename) not in ALLOWED_PDF:
        flash("Las etiquetas deben ser un archivo .pdf")
        return redirect(url_for("index"))

    workdir = tempfile.mkdtemp(prefix="lpn_")
    pedido_path = os.path.join(workdir, secure_filename(pedido.filename))
    etiquetas_path = os.path.join(workdir, secure_filename(etiquetas.filename))
    pedido.save(pedido_path)
    etiquetas.save(etiquetas_path)

    job_id = uuid.uuid4().hex
    _write_job(job_id, {"status": "procesando", "paso": "Iniciando...", "pagina": 0, "total": 0})

    hilo = threading.Thread(
        target=_procesar_job, args=(job_id, pedido_path, etiquetas_path, workdir), daemon=True
    )
    hilo.start()

    return redirect(url_for("estado", job_id=job_id))


@app.route("/estado/<job_id>")
def estado(job_id):
    job = _read_job(job_id)
    if job is None:
        flash("No se encontro ese proceso, intenta de nuevo.")
        return redirect(url_for("index"))

    if job["status"] == "error":
        flash(job["mensaje"])
        return redirect(url_for("index"))

    if job["status"] == "listo":
        return render_template(
            "resultado.html",
            n_pdf=job["n_pdf"],
            n_excel=job["n_excel"],
            resumen=job["resumen"],
            archivo_xlsx=job["archivo_xlsx"],
            archivo_pdf=job["archivo_pdf"],
        )

    return render_template("procesando.html", job_id=job_id, paso=job.get("paso", "Procesando..."))


@app.route("/descargar/<archivo>")
def descargar(archivo):
    safe_name = secure_filename(archivo)
    path = os.path.join(OUTPUT_DIR, safe_name)
    if not os.path.isfile(path):
        flash("El archivo ya no esta disponible, vuelve a procesarlo.")
        return redirect(url_for("index"))

    @after_this_request
    def cleanup(response):
        try:
            os.remove(path)
        except OSError:
            pass
        return response

    ext = _ext(safe_name)
    download_name = "etiquetas_con_codigo.pdf" if ext == ".pdf" else "pedido_con_LPN.xlsx"
    return send_file(path, as_attachment=True, download_name=download_name)


@app.errorhandler(413)
def too_large(e):
    flash("El archivo es demasiado grande (limite 40 MB).")
    return redirect(url_for("index"))


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
