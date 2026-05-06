"""
app.py
------
Servidor web Flask para o pipeline de Pedidos de Compra.
Expõe uma única página HTML com upload de .msg, progresso em tempo real
via Server-Sent Events (SSE) e download do relatório Excel.

Uso:
    python app.py
    → Abre http://localhost:5000 automaticamente no navegador.
"""

from __future__ import annotations

import json
import logging
import sys
import time
import threading
import webbrowser
from pathlib import Path

from flask import Flask, Response, jsonify, render_template, request, send_file

# ─── Paths ────────────────────────────────────────────────────────────────────
BASE_DIR = Path(__file__).parent
sys.path.insert(0, str(BASE_DIR))

INPUT_DIR   = BASE_DIR / "input_msgs"
TEMP_DIR    = BASE_DIR / "temp_pdfs"
OUTPUT_DIR  = BASE_DIR / "output"
OUTPUT_CSV  = OUTPUT_DIR / "mapeamento_produtos.csv"
OUTPUT_XLSX = OUTPUT_DIR / "consolidado_pedidos.xlsx"

for _d in [INPUT_DIR, TEMP_DIR, OUTPUT_DIR]:
    _d.mkdir(parents=True, exist_ok=True)

logging.basicConfig(level=logging.WARNING)

app = Flask(__name__)


# ─── Helpers ──────────────────────────────────────────────────────────────────

def _sse(data: dict) -> str:
    return f"data: {json.dumps(data, ensure_ascii=False)}\n\n"


def _clear_dir(directory: Path, pattern: str) -> None:
    for f in directory.glob(pattern):
        try:
            f.unlink()
        except OSError:
            pass


# ─── Routes ───────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/files")
def list_files():
    files = sorted(f.name for f in INPUT_DIR.glob("*.msg"))
    return jsonify({"files": files, "total": len(files)})


@app.route("/upload", methods=["POST"])
def upload():
    uploaded = request.files.getlist("files")
    saved: list[str] = []
    errors: list[str] = []

    for f in uploaded:
        if not f.filename:
            continue
        name = Path(f.filename).name
        if not name.lower().endswith(".msg"):
            errors.append(f"{name}: somente .msg são aceitos")
            continue
        dest = INPUT_DIR / name
        f.save(str(dest))
        saved.append(name)

    total = len(list(INPUT_DIR.glob("*.msg")))
    return jsonify({"saved": saved, "errors": errors, "total": total})


@app.route("/remove", methods=["POST"])
def remove():
    name = (request.json or {}).get("name", "")
    target = INPUT_DIR / Path(name).name  # prevents path traversal
    if target.exists() and target.suffix.lower() == ".msg":
        target.unlink()
    total = len(list(INPUT_DIR.glob("*.msg")))
    return jsonify({"total": total})


@app.route("/process")
def process():
    def generate():
        msg_files = list(INPUT_DIR.glob("*.msg"))
        if not msg_files:
            yield _sse({"type": "error", "message": "Nenhum arquivo .msg encontrado."})
            return

        yield _sse({"type": "start", "total_files": len(msg_files)})

        # Limpa saídas anteriores para não misturar lotes
        _clear_dir(TEMP_DIR, "*.pdf")

        # ── Etapa 1: Extração de PDFs dos .msg ───────────────────────────────
        yield _sse({"type": "stage", "stage": 1, "status": "running",
                    "message": "Extraindo anexos PDF dos emails..."})
        t0 = time.time()
        try:
            from extractor_msg import extract_pdf_attachments
            pdfs = extract_pdf_attachments(INPUT_DIR, TEMP_DIR)
            elapsed = f"{time.time() - t0:.1f}s"
            yield _sse({"type": "stage", "stage": 1, "status": "done",
                        "message": f"{len(pdfs)} PDF(s) extraído(s)",
                        "elapsed": elapsed, "count": len(pdfs)})
        except Exception as exc:
            yield _sse({"type": "stage", "stage": 1, "status": "error",
                        "message": str(exc)})
            yield _sse({"type": "done", "success": False})
            return

        # ── Etapa 2: Mapeamento ID → Nome ────────────────────────────────────
        yield _sse({"type": "stage", "stage": 2, "status": "running",
                    "message": "Mapeando IDs de produto para nomes..."})
        t0 = time.time()
        try:
            from extractor_ids import build_product_mapping
            mapping = build_product_mapping(TEMP_DIR, OUTPUT_CSV)
            elapsed = f"{time.time() - t0:.1f}s"
            yield _sse({"type": "stage", "stage": 2, "status": "done",
                        "message": f"{len(mapping)} produto(s) mapeado(s)",
                        "elapsed": elapsed, "count": len(mapping)})
        except Exception as exc:
            yield _sse({"type": "stage", "stage": 2, "status": "error",
                        "message": str(exc)})
            yield _sse({"type": "done", "success": False})
            return

        # ── Etapa 3: Consolidação e geração do Excel ─────────────────────────
        yield _sse({"type": "stage", "stage": 3, "status": "running",
                    "message": "Consolidando quantidades e gerando Excel..."})
        t0 = time.time()
        try:
            from extractor_quantidades import build_consolidation_report
            build_consolidation_report(TEMP_DIR, OUTPUT_CSV, OUTPUT_XLSX)
            elapsed = f"{time.time() - t0:.1f}s"
            yield _sse({"type": "stage", "stage": 3, "status": "done",
                        "message": "Relatório Excel gerado com sucesso",
                        "elapsed": elapsed})
        except Exception as exc:
            yield _sse({"type": "stage", "stage": 3, "status": "error",
                        "message": str(exc)})
            yield _sse({"type": "done", "success": False})
            return

        yield _sse({"type": "done", "success": True})

    return Response(
        generate(),
        mimetype="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.route("/download")
def download():
    if not OUTPUT_XLSX.exists():
        return jsonify({"error": "Relatório ainda não gerado"}), 404
    return send_file(
        str(OUTPUT_XLSX.absolute()),
        as_attachment=True,
        download_name="consolidado_pedidos.xlsx",
    )


@app.route("/reset", methods=["POST"])
def reset():
    _clear_dir(INPUT_DIR, "*.msg")
    _clear_dir(TEMP_DIR, "*.pdf")
    if OUTPUT_XLSX.exists():
        OUTPUT_XLSX.unlink(missing_ok=True)
    if OUTPUT_CSV.exists():
        OUTPUT_CSV.unlink(missing_ok=True)
    return jsonify({"status": "ok"})


# ─── Entry point ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    def _open_browser():
        time.sleep(1.2)
        webbrowser.open("http://localhost:5000")

    threading.Thread(target=_open_browser, daemon=True).start()
    app.run(host="127.0.0.1", port=5000, debug=False, threaded=True, use_reloader=False)
