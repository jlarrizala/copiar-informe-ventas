
from flask import Flask, request, jsonify
import io, os, base64
import requests
import msal
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from copy import copy

app = Flask(__name__)
GRAPH_BASE = "https://graph.microsoft.com/v1.0"

def get_token():
    tenant_id = os.environ["TENANT_ID"]
    client_id = os.environ["CLIENT_ID"]
    client_secret = os.environ["CLIENT_SECRET"]
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app_msal = msal.ConfidentialClientApplication(
        client_id, authority=authority, client_credential=client_secret
    )
    result = app_msal.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise RuntimeError(f"MSAL error: {result}")
    return result["access_token"]

def encode_sharing_url(url):
    b64 = base64.b64encode(url.encode()).decode()
    return "u!" + b64.rstrip("=").replace("/", "_").replace("+", "-")

def graph_get(url, token, stream=False):
    return requests.get(url, headers={"Authorization": f"Bearer {token}"}, allow_redirects=True, stream=stream)

def graph_put(url, token, content):
    return requests.put(url, headers={
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/octet-stream"
    }, data=content)

def copy_cell(src, dst):
    dst.value = src.value  # mantiene fórmula si hay
    if src.has_style:
        dst.font = copy(src.font)
        dst.border = copy(src.border)
        dst.fill = copy(src.fill)
        dst.number_format = copy(src.number_format)
        dst.protection = copy(src.protection)
        dst.alignment = copy(src.alignment)

@app.route("/copy-range", methods=["POST"])
def copy_range():
    data = request.get_json(force=True)

    # Parámetros requeridos
    required = ["source_sharing_url", "source_sheet", "source_range",
                "dest_sharing_url", "dest_sheet", "search_col_letter",
                "search_text", "offset_rows"]
    missing = [k for k in required if k not in data]
    if missing:
        return jsonify({"error": f"Faltan parámetros: {', '.join(missing)}"}), 400

    token = get_token()

    # --- Descargar ORIGEN ---
    src_encoded = encode_sharing_url(data["source_sharing_url"])
    src_resp = graph_get(f"{GRAPH_BASE}/shares/{src_encoded}/driveItem/content", token)
    if src_resp.status_code >= 400:
        return jsonify({"error": f"Descarga origen: {src_resp.status_code}"}), 400
    wb_src = load_workbook(io.BytesIO(src_resp.content), data_only=False)
    ws_src = wb_src[data["source_sheet"]]

    # Parseo del rango B10:AA34 (ejemplo)
    start_cell, end_cell = data["source_range"].split(":")
    start_col = column_index_from_string(''.join(filter(str.isalpha, start_cell)))
    start_row = int(''.join(filter(str.isdigit, start_cell)))
    end_col = column_index_from_string(''.join(filter(str.isalpha, end_cell)))
    end_row = int(''.join(filter(str.isdigit, end_cell)))
    col_count = end_col - start_col + 1
    row_count = end_row - start_row + 1

    # --- Descargar DESTINO ---
    dst_encoded = encode_sharing_url(data["dest_sharing_url"])
    meta = graph_get(f"{GRAPH_BASE}/shares/{dst_encoded}/driveItem", token).json()
    drive_id = meta["parentReference"]["driveId"]
    item_id = meta["id"]

    dst_resp = graph_get(f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/content", token)
    if dst_resp.status_code >= 400:
        return jsonify({"error": f"Descarga destino: {dst_resp.status_code}"}), 400
    wb_dst = load_workbook(io.BytesIO(dst_resp.content), data_only=False)
    ws_dst = wb_dst[data["dest_sheet"]]

    # --- Buscar texto en columna C (o la que indiques) ---
    col_idx = column_index_from_string(data["search_col_letter"])
    found_row = None
    for r in range(1, ws_dst.max_row + 1):
        val = ws_dst.cell(r, col_idx).value
        if isinstance(val, str) and val.strip() == data["search_text"].strip():
            found_row = r
            break
    if not found_row:
        return jsonify({"error": f'No se encontró "{data["search_text"]}" en columna {data["search_col_letter"]}'}), 400

    paste_row = found_row + int(data["offset_rows"])
    paste_col = col_idx

    # --- Copiar celdas con estilos ---
    for r_off in range(row_count):
        for c_off in range(col_count):
            src_cell = ws_src.cell(start_row + r_off, start_col + c_off)
            dst_cell = ws_dst.cell(paste_row + r_off, paste_col + c_off)
            copy_cell(src_cell, dst_cell)

    # (Opcional) Ajustar anchos de columna según origen: B..AA -> C..*
    # Si quieres, te agrego esta parte en la siguiente iteración.

    # --- Guardar y subir al mismo item (sobrescribe) ---
    out = io.BytesIO()
    wb_dst.save(out)
    out.seek(0)
    up_resp = graph_put(f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/content", token, out.read())
    if up_resp.status_code >= 400:
        return jsonify({"error": f"Upload destino: {up_resp.status_code}"}), 400

    return jsonify({"status": "ok", "paste_start": f"{get_column_letter(paste_col)}{paste_row}",
                    "rows": row_count, "cols": col_count})
