"""
actualizar_fondos.py  v8.1
────────────────────────
Fuente principal: API pública del buscador (screener) de Morningstar, consultada por ISIN.
Respaldo: página de snapshot de Morningstar ES (método de la v7).
Si un fondo no devuelve datos, el log de GitHub indica qué respondió Morningstar.
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

EMAIL_REMITENTE = os.environ.get("EMAIL_REMITENTE", "")
EMAIL_PASSWORD  = os.environ.get("EMAIL_PASSWORD", "")
EMAIL_DESTINO   = os.environ.get("EMAIL_DESTINO", "")
SIN_EMAIL       = os.environ.get("SIN_EMAIL", "0").strip() == "1"
EXCEL_PATH      = "fondos_inversion.xlsx"

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/122.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "es-ES,es;q=0.9",
}

# IDs de Morningstar ES (solo para el método de respaldo)
ISIN_A_MS = {
    "FR0010830885": "F00000XNKY",
    "LU1882473264": "F00000T1ZR",
    "LU1882449801": "F00001BHQK",
    "LU0906524193": "F00000YM0O",
    "LU1880401523": "F0GBR04R1K",
    "LU1706854152": "F00001BSIO",
    "FR0011365212": "F00000VXBK",
    "LU0658026603": "F00000NCI5",
    "LU0151324422": "F0GBR04GYK",
    "LU0252128276": "F0GBR04GYL",
    "LU0336084032": "F0GBR04QEK",
    "FR0010149120": "F0GBR04QEJ",
    "LU1966822956": "F00001DTXY",
    "LU1694789451": "F00001AK5D",
    "LU1161527038": "F00000VXYT",
    "LU0512127621": "F0GBR06BXX",
    "IE00B84J9L26": "F00000OUQB",
    "ES0173829047": "F0GBR04IFH",
    "LU2917874104": "F00001N9KY",
    "LU1551754515": "F00001AX9D",
    "LU0243957239": "F0GBR04QEG",
    "IE00B65YMK29": "F00000OUQC",
}


def _f(val):
    if val is None or val == "":
        return None
    try:
        s = str(val).replace("\xa0", "").replace(" ", "").replace(",", ".").replace("%", "").strip()
        return float(s)
    except (ValueError, TypeError):
        return None


SCREENER_HOSTS = [
    "https://lt.morningstar.com/api/rest.svc/klr5zyak8x/security/screener",
]

# ISIN erróneos en el Excel → ISIN correcto (se corrigen solos en la columna B)
CORRECCIONES_ISIN = {
    "FR0013399633": "FR0011399633",   # Amundi Euro Liquidity Short Term Govies E (C)
    "LU0830528907": "LU0839528907",   # Amundi Funds Euro Corporate Bond Select A2 EUR (C)
    "FR0014008AI4": "FR001400KAV4",   # Carmignac Credit 2029 A EUR Acc
    "FR0014000J453": "FR001400U4S3",  # Carmignac Credit 2031 A EUR Acc
    "LU0343303002": "LU0034353002",   # DWS Floating Rate Notes LC
    "FR0010334495": "FR0011034495",   # EdR SICAV Financial Bonds A EUR
    "IE0000BTZZE8": "IE00008TZ2E8",   # Man Dynamic Income D H EUR
    "IE000WFO9P18": "IE000WFG9PJ8",   # Man Dynamic Income DB H EUR
    "IE0000M5MJ59": "IE000JM1MJ59",   # Man Euro Corporate Bond D EUR
    "IE0000M53C66": "IE000MI53C66",   # Man Global Investment Grade Opportunities D H EUR
    "LU0300ZRW254": "IE00BDZRWZ54",   # Neuberger Berman Short Duration EMD EUR A Acc
}
SCREENER_UNIVERSOS = ["FOESP$$ALL", "FOEUR$$ALL", "FOGBR$$ALL"]
SCREENER_CAMPOS = "SecId|Name|ClosePrice|ClosePriceDate|ReturnM0|ReturnM1|ReturnM3|ReturnM12|StarRatingM255"


def obtener_datos_screener(isin: str) -> dict:
    """Consulta el screener de Morningstar filtrando por ISIN. Devuelve NAV, YTD, 1M, 3M, 1A y estrellas."""
    for host in SCREENER_HOSTS:
        for univ in SCREENER_UNIVERSOS:
            params = {
                "page": 1, "pageSize": 5, "sortOrder": "LegalName asc", "outputType": "json",
                "version": 1, "languageId": "es-ES", "currencyId": "EUR",
                "universeIds": univ, "securityDataPoints": SCREENER_CAMPOS, "term": isin,
            }
            try:
                r = requests.get(host, params=params, headers={**HEADERS, "Accept": "application/json"}, timeout=20)
                if r.status_code != 200:
                    log.info(f"    screener {host.split('/')[2]} {univ}: HTTP {r.status_code}")
                    continue
                filas = (r.json() or {}).get("rows") or []
                if not filas:
                    continue
                f = filas[0]
                datos = {}
                if _f(f.get("ClosePrice")):
                    datos["nav"] = round(_f(f.get("ClosePrice")), 4)
                for clave, campo in [("ytd", "ReturnM0"), ("1m", "ReturnM1"), ("3m", "ReturnM3"), ("1y", "ReturnM12")]:
                    v = _f(f.get(campo))
                    if v is not None:
                        datos[clave] = v
                est = _f(f.get("StarRatingM255"))
                if est:
                    datos["estrellas"] = int(est)
                if datos:
                    log.info(f"    screener OK ({univ}): {f.get('Name', '')[:40]} · fecha NAV {f.get('ClosePriceDate', '')}")
                    return datos
            except Exception as e:
                log.info(f"    screener {host.split('/')[2]} {univ}: error {type(e).__name__}: {str(e)[:120]}")
            time.sleep(0.3)
    return {}


def obtener_datos_snapshot(ms_id: str) -> dict:
    """Obtiene datos del snapshot de Morningstar ES (tab=0 y tab=1)."""
    datos = {}
    for tab in [0, 1]:
        url = f"https://www.morningstar.es/es/funds/snapshot/snapshot.aspx?id={ms_id}&tab={tab}"
        try:
            r = requests.get(url, headers=HEADERS, timeout=20)
            if r.status_code != 200:
                log.info(f"    snapshot tab={tab}: HTTP {r.status_code}")
                continue
            soup = BeautifulSoup(r.text, "html.parser")
            if tab == 0:
                for sel in ["div.snapshotTitleBox span.price", "span.price", "div.overviewKeyStatsTable td.value"]:
                    for tag in soup.select(sel):
                        txt = tag.get_text(strip=True).replace(".", "").replace(",", ".")
                        v = _f(txt)
                        if v and 0.01 < v < 999999:
                            datos["nav"] = round(v, 4)
                            break
                    if "nav" in datos:
                        break
                star_span = soup.find("span", class_=lambda c: c and "star" in c.lower())
                if star_span:
                    txt = star_span.get("title", "") or star_span.get_text()
                    for ch in txt:
                        if ch.isdigit():
                            datos["estrellas"] = int(ch)
                            break
            elif tab == 1:
                for tabla in soup.find_all("table"):
                    for fila in tabla.find_all("tr"):
                        celdas = [td.get_text(strip=True) for td in fila.find_all("td")]
                        if len(celdas) < 2:
                            continue
                        label = celdas[0].lower()
                        v = _f(celdas[1])
                        if v is None:
                            continue
                        if any(x in label for x in ["año en curso", "ytd", "acum"]):
                            if "ytd" not in datos:
                                datos["ytd"] = v
                        elif "1 mes" in label and "1m" not in datos:
                            datos["1m"] = v
                        elif "3 mes" in label and "3m" not in datos:
                            datos["3m"] = v
                        elif ("1 año" in label or "12 mes" in label) and "1y" not in datos:
                            datos["1y"] = v
                        elif "duraci" in label and "duracion" not in datos:
                            datos["duracion"] = v
                        elif ("ytm" in label or "rendim" in label) and "ytm" not in datos:
                            datos["ytm"] = v
        except Exception as e:
            log.debug(f"Snapshot {ms_id} tab={tab}: {e}")
        time.sleep(0.5)
    return datos


# ─── EXCEL ────────────────────────────────────────────────────────────────
AZUL_CL="D6E4F0"; BLANCO="FFFFFF"; AZUL_H="1F4E79"
VERDE="E2EFDA"; ROJO="FCE4D6"
thin = Border(left=Side(style="thin",color="BFBFBF"),right=Side(style="thin",color="BFBFBF"),
              top=Side(style="thin",color="BFBFBF"),bottom=Side(style="thin",color="BFBFBF"))

COL_DURACION=6; COL_YTM=7; COL_YTD=25; COL_ESTRELLAS=34
COL_NAV=36; COL_1M=37; COL_3M=38; COL_1Y=39; COL_ACTUALIZ=40

def celda(ws, row, col, valor, bg, fmt=None, bold=False):
    c = ws.cell(row=row, column=col, value=valor)
    c.font = Font(name="Arial", size=8, bold=bold, color="000000")
    c.fill = PatternFill("solid", start_color=bg)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = thin
    if fmt: c.number_format = fmt
    return c

def bg_rent(v):
    if v is None: return BLANCO
    return VERDE if v >= 0 else ROJO

def actualizar_excel():
    wb = load_workbook(EXCEL_PATH)
    ws = wb["Fondos - Datos Completos"]
    hoy = datetime.date.today().strftime("%d/%m/%Y")

    nuevas = {COL_NAV:"NAV\nÚltimo", COL_1M:"Rent.\n1 Mes",
              COL_3M:"Rent.\n3 Meses", COL_1Y:"Rent.\n1 Año", COL_ACTUALIZ:"Última\nActualiz."}
    for col, titulo in nuevas.items():
        c = ws.cell(row=2, column=col, value=titulo)
        c.font = Font(bold=True, color="FFFFFF", size=8, name="Arial")
        c.fill = PatternFill("solid", start_color=AZUL_H)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = thin
        ws.column_dimensions[get_column_letter(col)].width = 10

    fondos_ok = 0
    sin_datos = []

    for row in range(3, ws.max_row + 1):
        isin = ws.cell(row=row, column=2).value
        nombre = str(ws.cell(row=row, column=1).value or "")
        if not isin or str(isin).strip() == "":
            continue
        isin = str(isin).strip()
        if isin in CORRECCIONES_ISIN:
            nuevo = CORRECCIONES_ISIN[isin]
            log.info(f"ISIN corregido en Excel: {isin} → {nuevo} ({nombre[:35]})")
            ws.cell(row=row, column=2, value=nuevo)
            isin = nuevo
        bg_row = AZUL_CL if row % 2 == 1 else BLANCO

        log.info(f"Obteniendo: {isin} ({nombre[:35]})")
        datos = obtener_datos_screener(isin)

        # Respaldo: método antiguo (snapshot)
        if not datos and ISIN_A_MS.get(isin):
            datos = obtener_datos_snapshot(ISIN_A_MS[isin])
        time.sleep(1.0)

        if not datos:
            sin_datos.append(nombre[:40])
            celda(ws, row, COL_ACTUALIZ, hoy, bg_row)
            continue

        if datos.get("nav"):
            celda(ws, row, COL_NAV, datos["nav"], bg_row, fmt="#,##0.0000")
        if datos.get("ytd") is not None:
            celda(ws, row, COL_YTD, datos["ytd"]/100, bg_rent(datos["ytd"]), fmt="0.00%", bold=abs(datos["ytd"])>5)
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
        log.info(f"  ✓ NAV={datos.get('nav')} YTD={datos.get('ytd')}% 1M={datos.get('1m')}% 1Y={datos.get('1y')}%")

    if "Historial" not in wb.sheetnames:
        wlog = wb.create_sheet("Historial")
        for col, txt in [(1,"Fecha"),(2,"Actualizados"),(3,"Sin datos"),(4,"Notas")]:
            wlog.cell(1, col, txt).font = Font(bold=True, name="Arial")
    else:
        wlog = wb["Historial"]
    nr = wlog.max_row + 1
    wlog.cell(nr, 1, hoy); wlog.cell(nr, 2, fondos_ok)
    wlog.cell(nr, 3, ", ".join(sin_datos) if sin_datos else "—")
    wlog.cell(nr, 4, "v8.1")

    buf = BytesIO()
    wb.save(buf); buf.seek(0); wb.save(EXCEL_PATH)
    log.info(f"✅ {fondos_ok} fondos actualizados, {len(sin_datos)} sin datos.")
    return buf, fondos_ok, sin_datos


def enviar_email(excel_bytes, fondos_ok, sin_datos):
    hoy_str = datetime.date.today().strftime("%d/%m/%Y")
    nombre_arch = f"fondos_inversion_{datetime.date.today():%Y%m%d}.xlsx"
    msg = MIMEMultipart()
    msg["From"] = EMAIL_REMITENTE; msg["To"] = EMAIL_DESTINO
    msg["Subject"] = f"📊 Fondos de Inversión — {hoy_str} ({fondos_ok} actualizados)"
    sin_html = ""
    if sin_datos:
        items = "".join(f"<li>{f}</li>" for f in sin_datos)
        sin_html = f"<p style='color:#888;font-size:11px'>⚠️ Sin datos:<ul>{items}</ul></p>"
    cuerpo = f"""<html><body style="font-family:Arial,sans-serif;color:#333">
    <h2 style="color:#1F4E79">📊 Fondos de Inversión — {hoy_str}</h2>
    <p><strong>{fondos_ok} fondos</strong> actualizados con NAV, YTD, 1M, 3M, 1Y.</p>
    {sin_html}
    <p style="color:#888;font-size:11px">Fuente: Morningstar · GitHub Actions v8.1</p>
    </body></html>"""
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
    log.info("═══ Inicio actualización v8.1 ═══")
    excel_buf, fondos_ok, sin_datos = actualizar_excel()
    if fondos_ok == 0:
        log.warning("⚠️ Ningún fondo ha devuelto datos: revisa las líneas 'screener'/'snapshot' de arriba.")
    if SIN_EMAIL:
        log.info("Modo prueba (SIN_EMAIL=1): Excel actualizado, no se envía correo.")
    else:
        enviar_email(excel_buf, fondos_ok, sin_datos)
    log.info("═══ Proceso completado ═══")
