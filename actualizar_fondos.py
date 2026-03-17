"""
actualizar_fondos.py  v4
────────────────────────
Usa el screener de Morningstar que devuelve directamente
rentabilidades, NAV, duración y YTM sin necesidad de series temporales.
Funciona para fondos de cualquier domicilio (FR, LU, IE, ES...).
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
    """Convierte a float de forma segura."""
    if val is None or val == "":
        return None
    try:
        return float(str(val).replace(",", ".").replace("%", "").strip())
    except (ValueError, TypeError):
        return None


# ─── OBTENER DATOS VÍA SCREENER DE MORNINGSTAR ───────────────────────────
def obtener_datos(ms_id: str) -> dict:
    """
    Usa el screener de Morningstar ES que devuelve en una sola llamada:
    NAV, YTD, 1M, 3M, 1Y, estrellas, duración efectiva y YTM.
    Funciona para todos los dominios de fondo (FR, LU, IE, ES...).
    """
    campos = (
        "SecId|LegalName|StarRating|Nav|NavDate"
        "|GBRReturnM0|GBRReturnM1|GBRReturnM3|GBRReturnM12"
        "|EffectiveDuration|YieldToMaturity|CreditRating"
    )
    url = (
        "https://tools.morningstar.es/api/rest.svc/klr5zyak8x/security/screener"
        "?page=1&pageSize=1&sortOrder=LegalName+asc&outputType=json"
        "&version=1&languageId=es-ES&currencyId=EUR"
        "&universeIds=FOESP%24%24ALL%7CFOEUR%24%24ALL%7CFOIRL%24%24ALL%7CFOGBR%24%24ALL"
        f"&securityDataPoints={campos.replace('|', '%7C')}"
        f"&filters=SecId%3AIN%3A{ms_id}"
    )
    try:
        r = requests.get(url, headers=HEADERS, timeout=20)
        if r.status_code != 200:
            log.warning(f"Screener {ms_id}: status {r.status_code}")
            return {}
        data = r.json()
        rows = data.get("rows", [])
        if not rows:
            log.warning(f"Screener {ms_id}: sin resultados")
            return {}

        row = rows[0]
        datos = {}

        nav = _f(row.get("Nav"))
        if nav:
            datos["nav"] = nav
        if row.get("NavDate"):
            datos["fecha_nav"] = str(row["NavDate"])[:10]

        for key, campo in [
            ("ytd", "GBRReturnM0"),
            ("1m",  "GBRReturnM1"),
            ("3m",  "GBRReturnM3"),
            ("1y",  "GBRReturnM12"),
        ]:
            v = _f(row.get(campo))
            if v is not None:
                datos[key] = v

        dur = _f(row.get("EffectiveDuration"))
        if dur is not None:
            datos["duracion"] = dur

        ytm = _f(row.get("YieldToMaturity"))
        if ytm is not None:
            datos["ytm"] = ytm

        stars = row.get("StarRating")
        if stars:
            try:
                datos["estrellas"] = int(stars)
            except (ValueError, TypeError):
                pass

        return datos

    except Exception as e:
        log.warning(f"Screener {ms_id}: {e}")
        return {}


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


def color_rentabilidad(valor):
    """Verde si positivo, rojo si negativo."""
    if valor is None:
        return BLANCO
    return VERDE if valor >= 0 else ROJO


def actualizar_excel():
    wb  = load_workbook(EXCEL_PATH)
    ws  = wb["Fondos - Datos Completos"]
    hoy = datetime.date.today().strftime("%d/%m/%Y")

    # Cabeceras columnas nuevas
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
    fondos_sin_datos = []

    for row in range(3, ws.max_row + 1):
        isin = ws.cell(row=row, column=2).value
        nombre = ws.cell(row=row, column=1).value
        if not isin or str(isin).strip() == "":
            continue
        isin   = str(isin).strip()
        bg_row = AZUL_CL if row % 2 == 1 else BLANCO

        ms_id = ISIN_A_MS.get(isin)
        if not ms_id:
            log.warning(f"Sin ID Morningstar: {isin} ({nombre})")
            fondos_sin_datos.append(f"{nombre} (sin ID)")
            celda(ws, row, COL_ACTUALIZ, hoy, bg_row)
            continue

        log.info(f"  {isin} → {ms_id}")
        datos = obtener_datos(ms_id)
        time.sleep(0.8)

        if not datos:
            fondos_sin_datos.append(f"{nombre}")
            celda(ws, row, COL_ACTUALIZ, hoy, bg_row)
            continue

        # NAV
        if datos.get("nav"):
            celda(ws, row, COL_NAV, datos["nav"], bg_row, fmt="#,##0.0000")

        # YTD — en columna 25, con color
        if datos.get("ytd") is not None:
            bg_ytd = color_rentabilidad(datos["ytd"])
            celda(ws, row, COL_YTD, datos["ytd"] / 100, bg_ytd, fmt="0.00%",
                  bold=abs(datos["ytd"]) > 5)

        # 1M, 3M, 1Y — con color verde/rojo
        for col, key in [(COL_1M,"1m"),(COL_3M,"3m"),(COL_1Y,"1y")]:
            if datos.get(key) is not None:
                bg_r = color_rentabilidad(datos[key])
                celda(ws, row, col, datos[key] / 100, bg_r, fmt="0.00%")

        # Duración
        if datos.get("duracion") is not None:
            celda(ws, row, COL_DURACION, datos["duracion"], bg_row, fmt="0.00")

        # YTM
        if datos.get("ytm") is not None:
            celda(ws, row, COL_YTM, datos["ytm"] / 100, bg_row, fmt="0.00%")

        # Estrellas
        if datos.get("estrellas"):
            celda(ws, row, COL_ESTRELLAS, "⭐" * datos["estrellas"], bg_row)

        celda(ws, row, COL_ACTUALIZ, hoy, bg_row)
        fondos_ok += 1
        log.info(
            f"    ✓ NAV={datos.get('nav')} "
            f"YTD={datos.get('ytd')}% "
            f"1M={datos.get('1m')}% "
            f"3M={datos.get('3m')}% "
            f"1Y={datos.get('1y')}%"
        )

    if fondos_sin_datos:
        log.warning(f"Fondos sin datos ({len(fondos_sin_datos)}): {', '.join(fondos_sin_datos)}")

    # Historial
    if "Historial" not in wb.sheetnames:
        wlog = wb.create_sheet("Historial")
        for col, txt in [(1,"Fecha"),(2,"Fondos actualizados"),(3,"Sin datos"),(4,"Notas")]:
            wlog.cell(1, col, txt).font = Font(bold=True, name="Arial")
    else:
        wlog = wb["Historial"]
    nr = wlog.max_row + 1
    wlog.cell(nr, 1, hoy)
    wlog.cell(nr, 2, fondos_ok)
    wlog.cell(nr, 3, ", ".join(fondos_sin_datos) if fondos_sin_datos else "—")
    wlog.cell(nr, 4, "Actualización automática v4")

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    wb.save(EXCEL_PATH)
    log.info(f"✅ Excel guardado — {fondos_ok} fondos actualizados, {len(fondos_sin_datos)} sin datos.")
    return buf, fondos_ok, fondos_sin_datos


# ─── ENVIAR EMAIL ─────────────────────────────────────────────────────────
def enviar_email(excel_bytes, fondos_ok, fondos_sin_datos):
    hoy_str     = datetime.date.today().strftime("%d/%m/%Y")
    nombre_arch = f"fondos_inversion_{datetime.date.today():%Y%m%d}.xlsx"

    msg            = MIMEMultipart()
    msg["From"]    = EMAIL_REMITENTE
    msg["To"]      = EMAIL_DESTINO
    msg["Subject"] = f"📊 Fondos de Inversión — {hoy_str} ({fondos_ok} actualizados)"

    sin_datos_html = ""
    if fondos_sin_datos:
        items = "".join(f"<li>{f}</li>" for f in fondos_sin_datos)
        sin_datos_html = f"""
        <p style="color:#888;font-size:12px">
        ⚠️ Fondos sin datos disponibles en Morningstar:<br>
        <ul style="font-size:11px">{items}</ul>
        </p>
        """

    cuerpo = f"""
    <html><body style="font-family:Arial,sans-serif;color:#333">
    <h2 style="color:#1F4E79">📊 Fondos de Inversión — {hoy_str}</h2>
    <p>Excel actualizado con <strong>{fondos_ok} fondos</strong>.</p>
    <table style="border-collapse:collapse;margin:16px 0">
      <tr style="background:#1F4E79;color:white">
        <th style="padding:8px 16px">Dato</th><th style="padding:8px 16px">Estado</th>
      </tr>
      <tr><td style="padding:6px 16px">💰 NAV / Precio</td><td>✅ Actualizado</td></tr>
      <tr style="background:#f2f2f2"><td style="padding:6px 16px">📈 YTD</td><td>✅ Actualizado</td></tr>
      <tr><td style="padding:6px 16px">📈 Rent. 1M / 3M / 1Y</td><td>✅ Actualizado</td></tr>
      <tr style="background:#f2f2f2"><td style="padding:6px 16px">⭐ Estrellas Morningstar</td><td>✅ Actualizado</td></tr>
      <tr><td style="padding:6px 16px">⏱️ Duración y YTM</td><td>✅ Actualizado</td></tr>
    </table>
    {sin_datos_html}
    <p style="color:#888;font-size:11px">Fuente: Morningstar ES · GitHub Actions v4</p>
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


# ─── MAIN ─────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    log.info("═══ Inicio actualización v4 ═══")
    excel_buf, fondos_ok, sin_datos = actualizar_excel()
    enviar_email(excel_buf, fondos_ok, sin_datos)
    log.info("═══ Proceso completado ═══")
