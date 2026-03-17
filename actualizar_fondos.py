"""
actualizar_fondos.py  v2
────────────────────────
Obtiene NAV, rentabilidades (YTD, 1M, 3M, 1Y), rating y duración
de fondos de inversión europeos usando Morningstar como fuente principal.
"""

import os, time, datetime, smtplib, logging
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from io import BytesIO

import requests
from bs4 import BeautifulSoup
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
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/122.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "es-ES,es;q=0.9",
}

# ─── ISIN → ID MORNINGSTAR ────────────────────────────────────────────────
ISIN_A_MS = {
    "FR0010830885": "F00000XNKY",
    "FR0013399633": "F00001BSCK",
    "LU1882449801": "F00001BHQK",
    "LU0830528907": "F0GBR063CP",
    "LU0906524193": "F00000YM0O",
    "LU1706854152": "F00001BSIO",
    "FR0011365212": "F00000VXBK",
    "LU0658026603": "F00000NCI5",
    "LU0151324422": "F0GBR04GYK",
    "LU0252128276": "F0GBR04GYL",
    "FR0014008AI4": "F00001DLBY",
    "FR0014000J453":"F00001EO44",
    "LU0336084032": "F0GBR04QEK",
    "FR0010149120": "F0GBR04QEJ",
    "LU1966822956": "F00001DTXY",
    "LU1694789451": "F00001AK5D",
    "LU0343303002": "F0GBR06BXC",
    "LU1161527038": "F00000VXYT",
    "FR0010334495": "F0GBR04QEO",
    "LU0512127621": "F0GBR06BXX",
    "IE0000BTZZE8": "F00001GBA5",
    "IE000WFO9P18": "F00001HC6T",
    "IE0000M5MJ59": "F00001EM4X",
    "IE0000M53C66": "F00001GMLT",
    "LU0300ZRW254": "F0GBR06BXR",
    "IE00B84J9L26": "F00000OUQB",
    "ES0173829047": "F0GBR04IFH",
    "LU2917874104": "F00001N9KY",
    "LU1551754515": "F00001AX9D",
    "LU0243957239": "F0GBR04QEG",
    "LU0694789451": "F00001AK5D",
}


def _to_float(val):
    if val is None:
        return None
    try:
        return float(str(val).replace(",", ".").replace("%", "").strip())
    except (ValueError, TypeError):
        return None


def buscar_ms_id(isin):
    """Busca el ID de Morningstar por ISIN si no está en el diccionario."""
    url = (
        f"https://www.morningstar.es/es/util/SecuritySearch.ashx"
        f"?q={isin}&limit=1&universe=FOESP%24%24ALL%7CFOEUR%24%24ALL"
        f"&lang=es-ES&fmt=json"
    )
    try:
        r = requests.get(url, headers=HEADERS, timeout=10)
        data = r.json()
        results = data.get("r", [])
        if results:
            return results[0].get("id")
    except Exception as e:
        log.debug(f"Búsqueda MS {isin}: {e}")
    return None


def obtener_datos_ms(ms_id):
    """Obtiene NAV y rentabilidades de Morningstar ES por scraping."""
    datos = {}

    # — Pestaña de rendimiento (tab=1) —
    url = (
        f"https://www.morningstar.es/es/funds/snapshot/snapshot.aspx"
        f"?id={ms_id}&tab=1"
    )
    try:
        r = requests.get(url, headers=HEADERS, timeout=20)
        if r.status_code != 200:
            return datos
        soup = BeautifulSoup(r.text, "html.parser")

        # NAV — buscar en el bloque de precio
        for tag in soup.find_all(["span", "td", "div"]):
            cls = " ".join(tag.get("class", []))
            if "price" in cls or "nav" in cls.lower():
                v = _to_float(tag.get_text(strip=True))
                if v and 0.1 < v < 100000:
                    datos["nav"] = v
                    break

        # Rentabilidades — buscar en tablas
        for tabla in soup.find_all("table"):
            for fila in tabla.find_all("tr"):
                celdas = fila.find_all("td")
                if len(celdas) < 2:
                    continue
                label = celdas[0].get_text(strip=True).lower()
                val   = _to_float(
                    celdas[1].get_text(strip=True)
                    .replace(",", ".").replace("%", "")
                )
                if val is None:
                    continue
                if "1 mes" in label or "mes 1" in label:
                    datos["1m"] = val
                elif "3 mes" in label or "mes 3" in label:
                    datos["3m"] = val
                elif "1 año" in label or "12 mes" in label:
                    datos["1y"] = val
                elif "año en curso" in label or "ytd" in label:
                    datos["ytd"] = val
                elif "duraci" in label:
                    datos["duracion"] = val
                elif "ytm" in label or "rendim" in label:
                    datos["ytm"] = val

        # Estrellas Morningstar
        for img in soup.find_all("img"):
            alt = img.get("alt", "").lower()
            if "estrell" in alt or "star" in alt:
                for ch in alt:
                    if ch.isdigit():
                        datos["estrellas"] = int(ch)
                        break

    except Exception as e:
        log.warning(f"Scraping MS {ms_id}: {e}")

    # — Intentar también API de datos de rendimiento —
    try:
        api_url = (
            f"https://www.morningstar.es/es/funds/snapshot/snapshot.aspx"
            f"?id={ms_id}&tab=0"
        )
        r2 = requests.get(api_url, headers=HEADERS, timeout=15)
        if r2.status_code == 200:
            soup2 = BeautifulSoup(r2.text, "html.parser")
            # Buscar NAV en página principal
            spans = soup2.find_all("span", class_=lambda c: c and "price" in c)
            for sp in spans:
                v = _to_float(sp.get_text(strip=True))
                if v and 0.1 < v < 100000 and "nav" not in datos:
                    datos["nav"] = v
                    break
    except Exception:
        pass

    return datos


# ─── ACTUALIZAR EXCEL ─────────────────────────────────────────────────────
AZUL_CL = "D6E4F0"
BLANCO  = "FFFFFF"
AZUL_H  = "1F4E79"
thin = Border(
    left=Side(style="thin", color="BFBFBF"),
    right=Side(style="thin", color="BFBFBF"),
    top=Side(style="thin", color="BFBFBF"),
    bottom=Side(style="thin", color="BFBFBF"),
)

COL_DURACION = 6
COL_YTM      = 7
COL_RATING   = 8
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


def actualizar_excel():
    wb  = load_workbook(EXCEL_PATH)
    ws  = wb["Fondos - Datos Completos"]
    hoy = datetime.date.today().strftime("%d/%m/%Y")

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

    fondos_ok = 0
    for row in range(3, ws.max_row + 1):
        isin = ws.cell(row=row, column=2).value
        if not isin or str(isin).strip() == "":
            continue
        isin   = str(isin).strip()
        bg_row = AZUL_CL if row % 2 == 1 else BLANCO

        ms_id = ISIN_A_MS.get(isin) or buscar_ms_id(isin)
        if not ms_id:
            log.warning(f"Sin ID Morningstar: {isin}")
            celda(ws, row, COL_ACTUALIZ, hoy, bg_row)
            continue

        log.info(f"  {isin} → {ms_id}")
        datos = obtener_datos_ms(ms_id)
        time.sleep(1.5)

        if datos.get("nav"):
            celda(ws, row, COL_NAV, datos["nav"], bg_row, fmt="#,##0.0000")
        if datos.get("ytd") is not None:
            celda(ws, row, COL_YTD, datos["ytd"] / 100, bg_row, fmt="0.00%",
                  bold=datos["ytd"] > 5)
        for col, key in [(COL_1M, "1m"), (COL_3M, "3m"), (COL_1Y, "1y")]:
            if datos.get(key) is not None:
                celda(ws, row, col, datos[key] / 100, bg_row, fmt="0.00%")
        if datos.get("duracion") is not None:
            celda(ws, row, COL_DURACION, datos["duracion"], bg_row, fmt="0.00")
        if datos.get("ytm") is not None:
            celda(ws, row, COL_YTM, datos["ytm"] / 100, bg_row, fmt="0.00%")
        if datos.get("estrellas"):
            celda(ws, row, COL_ESTRELLAS, "⭐" * datos["estrellas"], bg_row)

        celda(ws, row, COL_ACTUALIZ, hoy, bg_row)
        fondos_ok += 1

    if "Historial" not in wb.sheetnames:
        wlog = wb.create_sheet("Historial")
        for col, txt in [(1,"Fecha"),(2,"Fondos actualizados"),(3,"Notas")]:
            wlog.cell(1, col, txt).font = Font(bold=True, name="Arial")
    else:
        wlog = wb["Historial"]
    nr = wlog.max_row + 1
    wlog.cell(nr, 1, hoy)
    wlog.cell(nr, 2, fondos_ok)
    wlog.cell(nr, 3, "Actualización automática diaria v2")

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    wb.save(EXCEL_PATH)
    log.info(f"Excel guardado — {fondos_ok} fondos actualizados.")
    return buf, fondos_ok


def enviar_email(excel_bytes, fondos_ok):
    hoy_str     = datetime.date.today().strftime("%d/%m/%Y")
    nombre_arch = f"fondos_inversion_{datetime.date.today():%Y%m%d}.xlsx"

    msg            = MIMEMultipart()
    msg["From"]    = EMAIL_REMITENTE
    msg["To"]      = EMAIL_DESTINO
    msg["Subject"] = f"📊 Fondos de Inversión — Actualización {hoy_str}"

    cuerpo = f"""
    <html><body style="font-family:Arial,sans-serif;color:#333">
    <h2 style="color:#1F4E79">📊 Fondos de Inversión — {hoy_str}</h2>
    <p>Se adjunta el Excel con <strong>{fondos_ok} fondos actualizados</strong>
    con NAV, rentabilidades (YTD, 1M, 3M, 1Y), duración y YTM.</p>
    <p style="color:#888;font-size:12px">
      Fuente: Morningstar ES · GitHub Actions
    </p>
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
    log.info("═══ Inicio actualización v2 ═══")
    excel_buf, fondos_ok = actualizar_excel()
    enviar_email(excel_buf, fondos_ok)
    log.info("═══ Proceso completado ═══")
