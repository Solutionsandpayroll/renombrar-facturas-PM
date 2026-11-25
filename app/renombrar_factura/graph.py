import base64
import io
import os
import re
import zipfile
import time
from pathlib import Path
import requests


def _cargar_env_local():
    """Carga un archivo .env (KEY=VALUE) desde la raiz del proyecto o de app/,
    sin dependencias externas. No pisa variables ya definidas en el entorno."""
    candidatos = [
        Path(__file__).resolve().parent.parent.parent / ".env",  # raiz del repo
        Path(__file__).resolve().parent.parent / ".env",          # carpeta app/
    ]
    for env_file in candidatos:
        if env_file.exists():
            for linea in env_file.read_text(encoding="utf-8").splitlines():
                linea = linea.strip()
                if not linea or linea.startswith("#") or "=" not in linea:
                    continue
                clave, valor = linea.split("=", 1)
                os.environ.setdefault(clave.strip(), valor.strip())


_cargar_env_local()

TENANT = os.environ.get("GRAPH_TENANT_ID", "")
CLIENT = os.environ.get("GRAPH_CLIENT_ID", "")
SECRET = os.environ.get("GRAPH_CLIENT_SECRET", "")

# El buzón donde llegan las facturas (cuenta compartida o de un usuario específico)
MAILBOX = os.environ.get("GRAPH_MAILBOX", "")

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# ── Cache simple de token en memoria (evita pedir uno nuevo en cada request) ──
_token_cache = {"token": None, "expira": 0}


def _obtener_token():
    """Devuelve un token válido, reutilizando el cacheado si no ha expirado."""
    if not (TENANT and CLIENT and SECRET and MAILBOX):
        raise RuntimeError(
            "Faltan credenciales de Microsoft Graph. Configure las variables "
            "GRAPH_TENANT_ID, GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET y GRAPH_MAILBOX "
            "(por ejemplo copiando .env.example a .env)."
        )

    ahora = time.time()
    if _token_cache["token"] and ahora < _token_cache["expira"] - 60:
        return _token_cache["token"]

    url = f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/token"
    datos = {
        "grant_type": "client_credentials",
        "client_id": CLIENT,
        "client_secret": SECRET,
        "scope": "https://graph.microsoft.com/.default",
    }
    resp = requests.post(url, data=datos, timeout=30)
    resp.raise_for_status()
    payload = resp.json()

    _token_cache["token"] = payload["access_token"]
    _token_cache["expira"] = ahora + payload.get("expires_in", 3600)
    return _token_cache["token"]


def _headers():
    return {
        "Authorization": f"Bearer {_obtener_token()}",
        "Content-Type": "application/json",
    }


def _buscar_mensajes_por_asunto(asunto, top=5):
    """
    Busca mensajes en MAILBOX cuyo asunto contenga el texto dado.
    Devuelve la lista de mensajes (id, hasAttachments) ordenados por fecha desc.
    """
    url = f"{GRAPH_BASE}/users/{MAILBOX}/messages"
    params = {
        "$search": f'"subject:{asunto}"',
        "$top": top,
        "$select": "id,subject,hasAttachments,receivedDateTime",
    }
    # $search requiere este header especial en Graph
    headers = _headers()
    headers["ConsistencyLevel"] = "eventual"

    resp = requests.get(url, headers=headers, params=params, timeout=60)

    if resp.status_code != 200:
        raise RuntimeError(f"Error buscando mensajes: {resp.status_code} {resp.text}")

    mensajes = resp.json().get("value", [])
    mensajes.sort(key=lambda m: m.get("receivedDateTime", ""), reverse=True)
    return mensajes


def _obtener_adjuntos(mensaje_id):
    """Devuelve la lista de adjuntos (con contentBytes) de un mensaje."""
    url = f"{GRAPH_BASE}/users/{MAILBOX}/messages/{mensaje_id}/attachments"
    resp = requests.get(url, headers=_headers(), timeout=60)

    if resp.status_code != 200:
        raise RuntimeError(f"Error obteniendo adjuntos: {resp.status_code} {resp.text}")

    return resp.json().get("value", [])


def buscar_pdf_por_asunto(asunto):
    """
    Equivalente a _buscar_pdf_por_asunto() pero usando Graph API en vez de Power Automate.
    Devuelve (pdf_bytes, pdf_nombre, error).
    """
    try:
        mensajes = _buscar_mensajes_por_asunto(asunto)
    except Exception as e:
        return None, None, str(e)

    if not mensajes:
        return None, None, "No se encontraron correos para ese asunto."

    errores = []

    for msg in mensajes:
        if not msg.get("hasAttachments"):
            continue

        try:
            adjuntos = _obtener_adjuntos(msg["id"])
        except Exception as e:
            errores.append(f"mensaje {msg['id']}: {e}")
            continue

        pdf_bytes, pdf_nombre = _extraer_pdf_de_adjuntos(adjuntos, errores)
        if pdf_bytes is not None:
            return pdf_bytes, pdf_nombre, None

    error = "No se encontró ningún PDF entre los adjuntos (ni directo ni dentro de un ZIP)."
    if errores:
        error += " Detalles: " + "; ".join(errores[:5])
    return None, None, error


def _asunto_contiene_factura(asunto, facturas):
    """
    Verifica si el asunto contiene alguno de los numeros de factura como
    numero completo (limitado por digitos), con o sin prefijo de letras.
    Ej: 'AXBC137396' coincide con '137396'; 'MP704166' coincide con '704166'.
    """
    if not asunto:
        return False
    for f in facturas:
        if not f:
            continue
        if re.search(rf"(?<!\d){re.escape(f)}(?!\d)", asunto, re.IGNORECASE):
            return True
    return False


def _escanear_mensajes_nit(criterio_search, facturas, max_mensajes=300, errores=None):
    """
    Pagina mensajes que coinciden con el criterio de busqueda dado y devuelve
    el primer PDF cuyo asunto contenga alguna de las facturas.
    Devuelve (pdf_bytes, pdf_nombre, vistos, candidatos, error).
    """
    if errores is None:
        errores = []

    url = f"{GRAPH_BASE}/users/{MAILBOX}/messages"
    params = {
        "$search": criterio_search,
        "$top": 100,
        "$select": "id,subject,hasAttachments,receivedDateTime",
    }
    headers = _headers()
    headers["ConsistencyLevel"] = "eventual"

    try:
        resp = requests.get(url, headers=headers, params=params, timeout=60)
    except Exception as e:
        return None, None, 0, 0, str(e)

    if resp.status_code != 200:
        return None, None, 0, 0, f"Error buscando mensajes por NIT: {resp.status_code} {resp.text}"

    data = resp.json()
    vistos = 0
    candidatos_totales = 0

    while True:
        pagina = data.get("value", [])
        vistos += len(pagina)

        candidatos = [
            m for m in pagina
            if m.get("hasAttachments") and _asunto_contiene_factura(m.get("subject", ""), facturas)
        ]
        candidatos_totales += len(candidatos)

        candidatos.sort(key=lambda m: m.get("receivedDateTime", ""), reverse=True)
        for msg in candidatos:
            try:
                adjuntos = _obtener_adjuntos(msg["id"])
            except Exception as e:
                errores.append(f"mensaje {msg['id']}: {e}")
                continue

            pdf_bytes, pdf_nombre = _extraer_pdf_de_adjuntos(adjuntos, errores)
            if pdf_bytes is not None:
                return pdf_bytes, pdf_nombre, vistos, candidatos_totales, None

        next_link = data.get("@odata.nextLink")
        if not next_link or vistos >= max_mensajes:
            break
        try:
            resp = requests.get(next_link, headers=headers, timeout=60)
        except Exception as e:
            return None, None, vistos, candidatos_totales, str(e)
        if resp.status_code != 200:
            break
        data = resp.json()

    return None, None, vistos, candidatos_totales, None


def buscar_pdf_por_nit_y_factura(nit, facturas, ventana_fechas=None, max_mensajes=300):
    """
    Estrategia alternativa: busca correos del proveedor por NIT y filtra
    localmente por numero de factura (maneja prefijos tipo AXBC, FE, MP, etc.).

    ventana_fechas: opcional (datetime_desde, datetime_hasta). Si se informa,
    primero se busca solo en ese rango (mas rapido); si no hay resultado,
    se reintenta en todo el historial para no perder correos fuera de ventana.

    Devuelve (pdf_bytes, pdf_nombre, error).
    """
    if not nit:
        return None, None, "NIT vacio, no se pudo buscar por proveedor."

    errores = []

    if ventana_fechas and all(ventana_fechas):
        desde, hasta = ventana_fechas
        d = desde.strftime("%Y-%m-%dT%H:%M:%SZ")
        h = hasta.strftime("%Y-%m-%dT%H:%M:%SZ")
        criterio = f'"subject:{nit} AND received:{d}..{h}"'
        pdf, nombre, _, _, err = _escanear_mensajes_nit(criterio, facturas, max_mensajes, errores)
        if err:
            return None, None, err
        if pdf is not None:
            return pdf, nombre, None

    pdf, nombre, vistos, cand, err = _escanear_mensajes_nit(f'"subject:{nit}"', facturas, max_mensajes, errores)
    if err:
        return None, None, err
    if pdf is not None:
        return pdf, nombre, None

    if cand == 0:
        return None, None, f"Se revisaron {vistos} correos del NIT {nit} y ninguno tiene la factura en el asunto."

    error = "Correo encontrado pero sin PDF en adjuntos (ni directo ni dentro de un ZIP)."
    if errores:
        error += " Detalles: " + "; ".join(errores[:5])
    return None, None, error


def _extraer_pdf_de_adjuntos(adjuntos, errores):
    """
    Recorre los adjuntos de Graph (tipo fileAttachment) buscando:
    - un PDF directo, o
    - un ZIP que contenga un PDF adentro.
    Ignora adjuntos inline (firmas, imágenes embebidas).
    """
    for archivo in adjuntos:
        # Ignorar adjuntos inline (firmas, logos embebidos, etc.)
        if archivo.get("isInline"):
            continue

        # Solo nos interesan los fileAttachment (traen contentBytes en base64)
        if archivo.get("@odata.type") != "#microsoft.graph.fileAttachment":
            continue

        nombre = archivo.get("name", "adjunto")
        contenido_b64 = archivo.get("contentBytes", "")

        if not contenido_b64:
            errores.append(f"{nombre}: sin contentBytes")
            continue

        try:
            contenido_bytes = base64.b64decode(contenido_b64)
        except Exception as e:
            errores.append(f"{nombre}: {e}")
            continue

        if nombre.lower().endswith(".pdf"):
            return contenido_bytes, nombre

        if nombre.lower().endswith(".zip"):
            try:
                with zipfile.ZipFile(io.BytesIO(contenido_bytes)) as zf:
                    pdfs_internos = [n for n in zf.namelist() if n.lower().endswith(".pdf")]
                    if not pdfs_internos:
                        errores.append(f"{nombre}: el ZIP no contiene ningún PDF")
                        continue
                    nombre_interno = pdfs_internos[0]
                    return zf.read(nombre_interno), nombre_interno
            except zipfile.BadZipFile:
                errores.append(f"{nombre}: no es un ZIP válido")
                continue

    return None, None