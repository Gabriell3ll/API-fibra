#!/usr/bin/env python3 
# calcular_rutas_full.py (versión con integración mejorada para n8n)
# Requisitos: pip install openpyxl flask gunicorn requests

import json
import math
import shutil
import os
import sys
import io
import requests
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from flask import Flask, request, jsonify, send_from_directory, make_response
from werkzeug.utils import secure_filename

# ====================
# CONFIG (entorno)
# ====================
PLANTILLA = os.environ.get("PLANTILLA", "Plantilla.xlsx")
OUTPUT_PREFIX = os.environ.get("OUTPUT_PREFIX", "Resultado")
OUTPUT_DIR = os.environ.get("OUTPUT_DIR", "/tmp")
# tamaño máximo para un JSON subido (bytes)
MAX_UPLOAD_BYTES = int(os.environ.get("MAX_UPLOAD_BYTES", 8 * 1024 * 1024))

# Asegurar directorio de salida
os.makedirs(OUTPUT_DIR, exist_ok=True)

# --------------------
# (--- AQUI VA TODA TU LÓGICA DE CÁLCULO TAL CUAL LA TIENES ---)
# Para no duplicar en el mensaje, asumo que las funciones:
# NivelKV, TipoDesdeNivel, distancia_m, angle_deg, leer_thresholds,
# mapear_tramos, sumar_tramo, absorber, FILL_KV y procesar_rutas(rutas)
# están definidas arriba *exactamente* como tú las compartiste.
#
# (En tu fichero, deja intactas todas esas funciones y el cuerpo de procesar_rutas)
# --------------------

# --- Para este fragmento de ejemplo las funciones están arriba tal cual.
# (Si pegas este bloque, asegúrate de mantener la definición completa de procesar_rutas tal cual.)

# -------------------
# Flask API (mejorada)
# -------------------
app = Flask(__name__)

def _parse_json_from_uploaded_file(file_storage):
    """Lee y parsea JSON desde un archivo subido (Form-Data 'file')."""
    raw = file_storage.stream.read(MAX_UPLOAD_BYTES + 1)
    if len(raw) > MAX_UPLOAD_BYTES:
        raise ValueError("Archivo demasiado grande")
    # intentar decodificar
    text = raw.decode('utf-8')
    return json.loads(text)

def _extract_rutas_from_payload(data):
    """
    Flexible extractor para aceptar varios formatos que podría enviar n8n:
    - JSON directo: lista de rutas
    - {"rutas": [ ... ] }
    - [ {"rutas": [...] } ]  (envuelto en array)
    - {"items": [{"json":{"rutas": [...]}} , ... ] }  (n8n workflow full)
    """
    # caso n8n full payload con items
    if isinstance(data, dict) and "items" in data and isinstance(data["items"], list):
        # intentar extraer rutas desde items[0].json.rutas o items[*].json
        for it in data["items"]:
            try:
                j = it.get("json", {})
                if "rutas" in j and isinstance(j["rutas"], list):
                    return j["rutas"]
            except Exception:
                continue
        # fallback: buscar cualquier .json.rutas
        for it in data["items"]:
            try:
                j = it.get("json", {})
                if isinstance(j, list):
                    return j
            except Exception:
                continue

    # caso {'rutas': [...]}
    if isinstance(data, dict) and "rutas" in data and isinstance(data["rutas"], list):
        return data["rutas"]

    # caso lista directa
    if isinstance(data, list):
        # si es lista de items que contienen key 'rutas' en cada item
        if len(data) > 0 and isinstance(data[0], dict) and "rutas" in data[0]:
            # tomar data[0]["rutas"]
            first = data[0].get("rutas")
            if isinstance(first, list):
                return first
        # si la lista parece ya ser la lista de rutas, devolverla
        # comprobación simple: item tiene 'puntos' o 'branch' keys
        if len(data) > 0 and isinstance(data[0], dict) and ("puntos" in data[0] or "branch" in data[0]):
            return data

    # caso {'body': {...}} o {'data': {...}}
    for key in ("body","data","payload"):
        if isinstance(data, dict) and key in data:
            return _extract_rutas_from_payload(data[key])

    # no se pudo extraer
    return None

@app.after_request
def add_cors_headers(response):
    # Cabeceras CORS básicas para que n8n o webhooks puedan consumir (ajusta en producción)
    response.headers['Access-Control-Allow-Origin'] = '*'
    response.headers['Access-Control-Allow-Headers'] = 'Content-Type, Authorization'
    response.headers['Access-Control-Allow-Methods'] = 'GET, POST, OPTIONS'
    return response

@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status":"ok"}), 200

@app.route("/procesar", methods=["POST", "OPTIONS"])
def api_procesar():
    """
    Endpoint principal:
    - acepta JSON en body
    - acepta Form-Data con archivo JSON en campo 'file'
    - si se pasa callback_url (query param o en el JSON), el servidor hará POST sincronizado con los resultados
    - devuelve JSON con resultados + URLs de descarga
    """
    # permitir preflight
    if request.method == "OPTIONS":
        return make_response(("", 204))

    try:
        rutas = None
        # 1) si viene un archivo multipart 'file'
        if request.files and 'file' in request.files:
            f = request.files['file']
            try:
                data = _parse_json_from_uploaded_file(f)
            except Exception as e:
                return jsonify({"error": f"Error leyendo archivo JSON: {str(e)}"}), 400
            rutas = _extract_rutas_from_payload(data)

        else:
            # 2) intentar JSON directo
            try:
                data = request.get_json(force=True, silent=True)
            except Exception:
                data = None

            # 3) si no hay JSON, intentar form field 'payload' (stringified JSON)
            if data is None and 'payload' in request.form:
                try:
                    data = json.loads(request.form['payload'])
                except Exception:
                    data = None

            # 4) extraer rutas
            if data is not None:
                rutas = _extract_rutas_from_payload(data)

        # 5) validación
        if rutas is None:
            return jsonify({"error":"No se pudo extraer la lista de rutas. Envíe JSON con 'rutas' o una lista de rutas."}), 400

        if not isinstance(rutas, list):
            return jsonify({"error":"El JSON debe contener una lista de rutas."}), 400

        # opcional: obtener callback_url (query param o dentro del payload)
        callback_url = request.args.get("callback_url") or (data.get("callback_url") if isinstance(data, dict) else None)

        # llamos a la lógica de cálculo (sin cambios)
        resultados = procesar_rutas(rutas)

        # montar download URLs (host_url puede incluir puerto)
        host = request.host_url.rstrip('/')
        for r in resultados:
            if "file" in r:
                r["download_url"] = f"{host}/download/{r['file']}"

        resp_body = {"status":"ok", "results": resultados}

        # Si viene callback_url: enviar resultado de forma síncrona (blocking).
        callback_status = None
        if callback_url:
            try:
                cb = requests.post(callback_url, json=resp_body, timeout=20)
                callback_status = {"status_code": cb.status_code, "ok": cb.ok}
            except Exception as e:
                callback_status = {"error": str(e)}
            resp_body["callback_status"] = callback_status

        return jsonify(resp_body), 200

    except FileNotFoundError as e:
        return jsonify({"error": str(e)}), 500
    except Exception as e:
        # registrar en consola y devolver 500
        print("Error en /procesar:", str(e))
        return jsonify({"error": str(e)}), 500

@app.route("/download/<filename>", methods=["GET"])
def download_file(filename):
    # servir solo desde OUTPUT_DIR
    safe_name = secure_filename(filename)
    file_path = os.path.join(OUTPUT_DIR, safe_name)
    if not os.path.exists(file_path):
        return jsonify({"error":"file not found"}), 404
    return send_from_directory(OUTPUT_DIR, safe_name, as_attachment=True)

if __name__ == "__main__":
    # En producción usar gunicorn: e.g. gunicorn -w 4 "calcular_rutas_full:app"
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 10000)))
