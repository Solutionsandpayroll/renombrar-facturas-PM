import base64
import io
import re
import zipfile
import time
import requests

TENANT = "a104dec4-8857-4cf3-b96c-53796c1dab1e"
CLIENT = "945b1b53-72a6-42d8-8cd5-a4a6efa4c162"
SECRET = "h1e8Q~elYjZDpyloye~-H2DofcDJtZv3pSIU8ba."

# El buzón donde llegan las facturas (cuenta compartida o de un usuario específico)
MAILBOX = "servicioalcliente@puntomedical.com.co"  # <-- ajusta esto

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# ── Cache simple de token en memoria (evita pedir uno nuevo en cada request) ──
_token_cache = {"token": None, "expira": 0}


def _obtener_token():
    """Devuelve un token válido, reutilizando el cacheado si no ha expirado."""
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


def _listar_mensajes_por_nit(nit, max_mensajes=300):
    """
    Busca mensajes cuyo asunto contenga el NIT como token (los asuntos de
    facturas tienen formato NIT;PROVEEDOR;PREFIJO+FACTURA;...).
    Pagina resultados hasta max_mensajes.
    """
    url = f"{GRAPH_BASE}/users/{MAILBOX}/messages"
    params = {
        "$search": f'"subject:{nit}"',
        "$top": 100,
        "$select": "id,subject,hasAttachments,receivedDateTime",
    }
    headers = _headers()
    headers["ConsistencyLevel"] = "eventual"

    todos = []
    resp = requests.get(url, headers=headers, params=params, timeout=60)
    if resp.status_code != 200:
        raise RuntimeError(f"Error buscando mensajes por NIT: {resp.status_code} {resp.text}")

    data = resp.json()
    todos.extend(data.get("value", []))
    next_link = data.get("@odata.nextLink")

    while next_link and len(todos) < max_mensajes:
        resp = requests.get(next_link, headers=headers, timeout=60)
        if resp.status_code != 200:
            break
        data = resp.json()
        todos.extend(data.get("value", []))
        next_link = data.get("@odata.nextLink")

    todos = todos[:max_mensajes]
    todos.sort(key=lambda m: m.get("receivedDateTime", ""), reverse=True)
    return todos


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


def buscar_pdf_por_nit_y_factura(nit, facturas):
    """
    Estrategia alternativa: busca correos del proveedor por NIT y filtra
    localmente por numero de factura (maneja prefijos tipo AXBC, FE, MP, etc.).
    Devuelve (pdf_bytes, pdf_nombre, error).
    """
    if not nit:
        return None, None, "NIT vacio, no se pudo buscar por proveedor."

    try:
        mensajes = _listar_mensajes_por_nit(nit)
    except Exception as e:
        return None, None, str(e)

    candidatos = [m for m in mensajes if m.get("hasAttachments") and _asunto_contiene_factura(m.get("subject", ""), facturas)]

    if not candidatos:
        return None, None, f"Se buscaron {len(mensajes)} correos del NIT {nit} y ninguno tiene la factura en el asunto."

    errores = []
    for msg in candidatos:
        try:
            adjuntos = _obtener_adjuntos(msg["id"])
        except Exception as e:
            errores.append(f"mensaje {msg['id']}: {e}")
            continue

        pdf_bytes, pdf_nombre = _extraer_pdf_de_adjuntos(adjuntos, errores)
        if pdf_bytes is not None:
            return pdf_bytes, pdf_nombre, None

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