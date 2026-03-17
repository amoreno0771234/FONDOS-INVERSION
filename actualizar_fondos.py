"""
actualizar_fondos.py  v6
────────────────────────
Busca cada fondo DIRECTAMENTE por ISIN en Morningstar.
No depende de IDs hardcodeados que pueden estar desactualizados.
"""

import os, time, datetime, smtplib, logging
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from io import BytesIO

import requests
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─── CONFIGURACIÓN ────────────────────────────────────────────────────────
EMAIL_REMITENTE = os.environ["EMAIL_REMITENTE"]
EMAIL_PASSWORD  = os.environ["EMAIL_PASSWORD"]
EMAIL_DESTINO   = os.environ["EMAIL_DESTINO"]
EXCEL_PATH      = "fondos_inversion.xlsx"

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/122.0 Safari/537.36",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "es-ES,es;q=0.9",
    "Referer": "https://www.morningstar.es/",
}


def _f(val):
    if val is None or val == "":
        return None
    try:
        return float(str(val).replace(",", ".").replace("%", "").strip())
    except (ValueError, TypeError):
        return None


# ─── PASO 1: BUSCAR MS_ID POR ISIN ───────────────────────────────────────
def buscar_id_por_isin(isin: str) -> tuple:
    """
    Busca el fondo en Morningstar por ISIN.
    Devuelve (ms_id, nombre) o (None, None).
    """
    url = (
        "https://www.morningstar.es/es/util/SecuritySearch.ashx"
        f"?q={isin}&limit=5&universe=FOESP%24%24ALL%7CFOEUR%24%24ALL"
        f"%7CFOIRL%24%24ALL%7CFOGBR%24%24ALL%7CFOEAA%24%24ALL"
        "&lang=es-ES&fmt=json"
    )
    try:
        r = requests.get(url, headers=HEADERS, timeout=15)
        if r.status_code != 200:
            return None, None
        data = r.json()
        resultados = data.get("r", [])
        if resultados:
            mejor = resultados[0]
            return mejor.get("id"), mejor.get("n", "")
    except Exception as e:
        log.debug(f"Búsqueda ISIN {isin}: {e}")
    return None, None


# ─── PASO 2: OBTENER DATOS POR MS_ID ─────────────────────────────────────
def obtener_datos_por_id(ms_id: str) -> dict:
    """
    Obtiene NAV, rentabilidades, duración, YTM y estrellas
    usando el endpoint de ficha de Morningstar.
    """
    datos = {}

    # Endpoint de rendimiento
    url = (
        f"https://www.morningstar.es/es/funds/snapshot/snapshot.aspx"
        f"?id={ms_id}&tab=1"
    )
    try:
        r = requests.get(url, headers=HEADERS, timeout=20)
        if r.status_code != 200:
            return datos

        from bs4 import BeautifulSoup
        soup = BeautifulSoup(r.text, "html.parser")

        # NAV desde bloque de precio superior
        for sel in ["div.snapshot-header span.price",
                    "span.price", "td.price"]:
            tag = soup.select_one(sel)
            if tag:
                v = _f(tag.get_text(strip=True).replace(".", "").replace(",", "."))
                if v and 0.01 < v < 999999:
                    datos["nav"] = round(v, 4)
                    break

        # Rentabilidades desde tablas
        for tabla in soup.find_all("table"):
            for fila in tabla.find_all("tr"):
                celdas = [td.get_text(strip=True) for td in fila.find_all("td")]
                if len(celdas) < 2:
                    continue
                label = celdas[0].lower()
                raw   = celdas[1].replace(",", ".").replace("%", "").strip()
                v     = _f(raw)
                if v is None:
                    continue
                if any(x in label for x in ["año en curso", "ytd", "acum. año"]):
                    datos["ytd"] = v
                elif "1 mes" in label:
                    datos["1m"] = v
                elif "3 mes" in label:
                    datos["3m"] = v
                elif any(x in label for x in ["1 año", "12 mes"]):
                    datos["1y"] = v
                elif "duraci" in label:
                    datos["duracion"] = v
                elif "ytm" in label or "rendim" in label:
                    datos["ytm"] = v

        # Estrellas
        for img in soup.find_all("img"):
            alt = (img.get("alt") or img.get("title") or "").lower()
            if "estrell" in alt or "star" in alt:
                for ch in alt:
                    if ch.isdigit():
                        datos["estrellas"] = int(ch)
                        break

    except Exception as e:
        log.warning(f"Ficha {ms_id}: {e}")

    return datos


# ─── CACHÉ DE IDs PARA NO BUSCAR CADA DÍA ────────────────────────────────
# Se construye en la primera ejecución y se reutiliza
_CACHE_IDS = {}

def obtener_datos_fondo(isin: str, nombre: str) -> dict:
    """Busca el ID si no lo tenemos y luego obtiene los datos."""
    global _CACHE_IDS

    ms_id = _CACHE_IDS.get(isin)
    if not ms_id:
        ms_id, nombre_ms = buscar_id_por_isin(isin)
        if ms_id:
            _CACHE_IDS[isin] = ms_id
            log.info(f"  ISIN {isin} → ID {ms_id} ({nombre_ms[:40]})")
        else:
            log.warning(f"  ISIN {isin} no encontrado en Morningstar ({nombre})")
            return {}
        time.sleep(0.5)

    datos = obtener_datos_por_id(ms_id)
    return datos


# ─── ACTUALIZAR EXCEL ─────────────────────────────────────────────────────
AZUL_CL = "D6E4F0"
BLANCO  = "FFFFFF"
AZUL_H  = "1F4E79"
VERDE   = "E2EFDA"
ROJO    = "FCE4D6"
thin = Border(
    left=Side(style="thin", color="BFBFBF"),
    right=Side(style="thin", color="BFBFBF"),
    top=Side(style="thin", color="BFBFBF"),
    bottom=Side(style="thin", color="BFBFBF"),
)

COL_DURACION = 6
COL_YTM      = 7
COL_YTD      = 25
COL_ESTRELLAS= 34
COL_NAV      = 36
COL_1M       = 37
COL_3M       = 38
COL_1Y       = 39
COL_ACTUALIZ = 40


def celda(ws, row, col, valor, bg, fmt=None, bold=False):
    c = ws.cell(row=row, column=col, value=valor)
    c.font      = Font(name="Arial", size=8, bold=bold, color="000000")
    c.fill      = PatternFill("solid", start_color=bg)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border    = thin
    if fmt:
        c.number_format = fmt
    return c


def bg_rent(v):
    if v is None: return BLANCO
    return VERDE if v >= 0 else ROJO


def actualizar_excel():
    wb  = load_workbook(EXCEL_PATH)
    ws  = wb["Fondos - Datos Completos"]
    hoy = datetime.date.today().strftime("%d/%m/%Y")

    # Cabeceras
    nuevas = {
        COL_NAV:     "NAV\nÚltimo",
        COL_1M:      "Rent.\n1 Mes",
        COL_3M:      "Rent.\n3 Meses",
        COL_1Y:      "Rent.\n1 Año",
        COL_ACTUALIZ:"Última\nActualiz.",
    }
    for col, titulo in nuevas.items():
        c = ws.cell(row=2, column=col, value=titulo)
        c.font      = Font(bold=True, color="FFFFFF", size=8, name="Arial")
        c.fill      = PatternFill("solid", start_color=AZUL_H)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border    = thin
        ws.column_dimensions[get_column_letter(col)].width = 10

    fondos_ok  = 0
    sin_datos  = []

    for row in range(3, ws.max_row + 1):
        isin   = ws.cell(row=row, column=2).value
        nombre = ws.cell(row=row, column=1).value or ""
        if not isin or str(isin).strip() == "":
            continue
        isin   = str(isin).strip()
        bg_row = AZUL_CL if row % 2 == 1 else BLANCO

        log.info(f"Procesando: {isin} — {str(nombre)[:35]}")
        datos = obtener_datos_fondo(isin, str(nombre))
        time.sleep(1.2)

        if not datos:
            sin_datos.append(str(nombre)[:40])
            celda(ws, row, COL_ACTUALIZ, hoy, bg_row)
            continue

        if datos.get("nav"):
            celda(ws, row, COL_NAV, datos["nav"], bg_row, fmt="#,##0.0000")
        if datos.get("ytd") is not None:
            celda(ws, row, COL_YTD, datos["ytd"]/100,
                  bg_rent(datos["ytd"]), fmt="0.00%", bold=abs(datos["ytd"])>5)
        for col, key in [(COL_1M,"1m"),(COL_3M,"3m"),(COL_1Y,"1y")]:
            if datos.get(key) is not None:
                celda(ws, row, col, datos[key]/100, bg_rent(datos[key]), fmt="0.00%")
        if datos.get("duracion") is not None:
            celda(ws, row, COL_DURACION, datos["duracion"], bg_row, fmt="0.00")
        if datos.get("ytm") is not None:
            celda(ws, row, COL_YTM, datos["ytm"]/100, bg_row, fmt="0.00%")
        if datos.get("estrellas"):
            celda(ws, row, COL_ESTRELLAS, "⭐"*datos["estrellas"], bg_row)

        celda(ws, row, COL_ACTUALIZ, hoy, bg_row)
        fondos_ok += 1
        log.info(
            f"  ✓ NAV={datos.get('nav')} "
            f"YTD={datos.get('ytd')}% "
            f"1M={datos.get('1m')}% "
            f"3M={datos.get('3m')}% "
            f"1Y={datos.get('1y')}%"
        )

    # Historial
    if "Historial" not in wb.sheetnames:
        wlog = wb.create_sheet("Historial")
        for col, txt in [(1,"Fecha"),(2,"Actualizados"),(3,"Sin datos"),(4,"Notas")]:
            wlog.cell(1, col, txt).font = Font(bold=True, name="Arial")
    else:
        wlog = wb["Historial"]
    nr = wlog.max_row + 1
    wlog.cell(nr, 1, hoy)
    wlog.cell(nr, 2, fondos_ok)
    wlog.cell(nr, 3, ", ".join(sin_datos) if sin_datos else "—")
    wlog.cell(nr, 4, "v6 - búsqueda por ISIN")

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    wb.save(EXCEL_PATH)
    log.info(f"✅ {fondos_ok} fondos actualizados, {len(sin_datos)} sin datos.")
    return buf, fondos_ok, sin_datos


def enviar_email(excel_bytes, fondos_ok, sin_datos):
    hoy_str     = datetime.date.today().strftime("%d/%m/%Y")
    nombre_arch = f"fondos_inversion_{datetime.date.today():%Y%m%d}.xlsx"

    msg            = MIMEMultipart()
    msg["From"]    = EMAIL_REMITENTE
    msg["To"]      = EMAIL_DESTINO
    msg["Subject"] = f"📊 Fondos de Inversión — {hoy_str} ({fondos_ok} actualizados)"

    sin_html = ""
    if sin_datos:
        items = "".join(f"<li>{f}</li>" for f in sin_datos)
        sin_html = f"<p style='color:#888;font-size:11px'>⚠️ Sin datos en Morningstar:<ul>{items}</ul></p>"

    cuerpo = f"""
    <html><body style="font-family:Arial,sans-serif;color:#333">
    <h2 style="color:#1F4E79">📊 Fondos de Inversión — {hoy_str}</h2>
    <p><strong>{fondos_ok} fondos</strong> actualizados con NAV, YTD, 1M, 3M, 1Y.</p>
    {sin_html}
    <p style="color:#888;font-size:11px">Fuente: Morningstar ES · GitHub Actions v6</p>
    </body></html>
    """
    msg.attach(MIMEText(cuerpo, "html"))

    parte = MIMEBase("application", "octet-stream")
    parte.set_payload(excel_bytes.read())
    encoders.encode_base64(parte)
    parte.add_header("Content-Disposition", f'attachment; filename="{nombre_arch}"')
    msg.attach(parte)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as srv:
        srv.login(EMAIL_REMITENTE, EMAIL_PASSWORD)
        srv.sendmail(EMAIL_REMITENTE, EMAIL_DESTINO, msg.as_string())
    log.info(f"Email enviado a {EMAIL_DESTINO}")


if __name__ == "__main__":
    log.info("═══ Inicio actualización v6 ═══")
    excel_buf, fondos_ok, sin_datos = actualizar_excel()
    enviar_email(excel_buf, fondos_ok, sin_datos)
    log.info("═══ Proceso completado ═══")
