"""
actualizar_fondos.py  v5
────────────────────────
Corrige:
- NAV obtenido correctamente (campo correcto del screener)
- Fondos F0GBR incluidos (universo FOGBR$$ALL añadido)
- Consulta por lotes para mayor fiabilidad
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

# Todos los universos: España, Europa, Irlanda, UK, Global
UNIVERSOS = "FOESP%24%24ALL%7CFOEUR%24%24ALL%7CFOIRL%24%24ALL%7CFOGBR%24%24ALL%7CFOEAA%24%24ALL"

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
    "FR0014000J453": "F00001EO44",
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
}


def _f(val):
    if val is None or val == "":
        return None
    try:
        return float(str(val).replace(",", ".").replace("%", "").strip())
    except (ValueError, TypeError):
        return None


def obtener_datos_lote(ids: list) -> dict:
    """
    Consulta el screener de Morningstar para un lote de IDs.
    Devuelve dict {ms_id: {nav, ytd, 1m, 3m, 1y, duracion, ytm, estrellas}}
    """
    if not ids:
        return {}

    # Campos a obtener — incluye Nav, NavOfDay y Price por si acaso
    campos = (
        "SecId|LegalName|StarRating"
        "|Nav|NavOfDay|ClosePrice"
        "|GBRReturnM0|GBRReturnM1|GBRReturnM3|GBRReturnM12"
        "|EffectiveDuration|YieldToMaturity"
    )

    filtro = "SecId%3AIN%3A" + "%7C".join(ids)

    url = (
        "https://tools.morningstar.es/api/rest.svc/klr5zyak8x/security/screener"
        f"?page=1&pageSize={len(ids)}&sortOrder=LegalName+asc&outputType=json"
        f"&version=1&languageId=es-ES&currencyId=EUR"
        f"&universeIds={UNIVERSOS}"
        f"&securityDataPoints={campos.replace('|', '%7C')}"
        f"&filters={filtro}"
    )

    try:
        r = requests.get(url, headers=HEADERS, timeout=25)
        if r.status_code != 200:
            log.warning(f"Screener lote: status {r.status_code}")
            return {}
        data = r.json()
        rows = data.get("rows", [])
    except Exception as e:
        log.warning(f"Screener lote: {e}")
        return {}

    resultado = {}
    for row in rows:
        sec_id = row.get("SecId")
        if not sec_id:
            continue

        datos = {}

        # NAV — intentar varios campos
        for campo_nav in ["Nav", "NavOfDay", "ClosePrice"]:
            v = _f(row.get(campo_nav))
            if v and v > 0:
                datos["nav"] = round(v, 4)
                break

        # Rentabilidades
        for key, campo in [
            ("ytd", "GBRReturnM0"),
            ("1m",  "GBRReturnM1"),
            ("3m",  "GBRReturnM3"),
            ("1y",  "GBRReturnM12"),
        ]:
            v = _f(row.get(campo))
            if v is not None:
                datos[key] = v

        # Duración y YTM
        v = _f(row.get("EffectiveDuration"))
        if v is not None:
            datos["duracion"] = v
        v = _f(row.get("YieldToMaturity"))
        if v is not None:
            datos["ytm"] = v

        # Estrellas
        stars = row.get("StarRating")
        if stars:
            try:
                datos["estrellas"] = int(stars)
            except (ValueError, TypeError):
                pass

        resultado[sec_id] = datos
        nombre = row.get("LegalName", "")[:40]
        log.info(
            f"  ✓ {sec_id} ({nombre}) "
            f"NAV={datos.get('nav')} "
            f"YTD={datos.get('ytd')}% "
            f"1M={datos.get('1m')}% "
            f"1Y={datos.get('1y')}%"
        )

    return resultado


# ─── COLUMNAS EXCEL ───────────────────────────────────────────────────────
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

    # Recoger todos los ISINs y sus filas
    filas = {}
    for row in range(3, ws.max_row + 1):
        isin = ws.cell(row=row, column=2).value
        if not isin or str(isin).strip() == "":
            continue
        isin = str(isin).strip()
        ms_id = ISIN_A_MS.get(isin)
        if ms_id:
            filas[ms_id] = row
        else:
            nombre = ws.cell(row=row, column=1).value
            log.warning(f"Sin ID MS: {isin} ({nombre})")
            bg_row = AZUL_CL if row % 2 == 1 else BLANCO
            celda(ws, row, COL_ACTUALIZ, hoy, bg_row)

    # Consultar en lotes de 10
    todos_ids  = list(filas.keys())
    todos_datos = {}
    for i in range(0, len(todos_ids), 10):
        lote = todos_ids[i:i+10]
        log.info(f"Consultando lote {i//10+1}: {lote}")
        res = obtener_datos_lote(lote)
        todos_datos.update(res)
        time.sleep(1.0)

    # Escribir en Excel
    fondos_ok = 0
    sin_datos = []
    for ms_id, row in filas.items():
        bg_row = AZUL_CL if row % 2 == 1 else BLANCO
        datos  = todos_datos.get(ms_id, {})

        if not datos:
            nombre = ws.cell(row=row, column=1).value
            sin_datos.append(str(nombre)[:40])
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
    wlog.cell(nr, 4, "v5")

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
        sin_html = f"<p style='color:#888;font-size:11px'>⚠️ Sin datos: <ul>{items}</ul></p>"

    cuerpo = f"""
    <html><body style="font-family:Arial,sans-serif;color:#333">
    <h2 style="color:#1F4E79">📊 Fondos de Inversión — {hoy_str}</h2>
    <p><strong>{fondos_ok} fondos</strong> actualizados con NAV, YTD, 1M, 3M, 1Y, duración y YTM.</p>
    {sin_html}
    <p style="color:#888;font-size:11px">Fuente: Morningstar ES · GitHub Actions v5</p>
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
    log.info("═══ Inicio actualización v5 ═══")
    excel_buf, fondos_ok, sin_datos = actualizar_excel()
    enviar_email(excel_buf, fondos_ok, sin_datos)
    log.info("═══ Proceso completado ═══")
