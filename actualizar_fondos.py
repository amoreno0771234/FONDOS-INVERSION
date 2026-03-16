"""
actualizar_fondos.py
────────────────────
Actualiza diariamente los fondos de inversión:
  • NAV / Precio (Yahoo Finance)
  • Rentabilidad YTD, 1M, 3M, 1Y (Yahoo Finance / cálculo)
  • Rating crediticio (Morningstar scraping)
  • Duración y YTM (Morningstar scraping)

Luego envía el Excel por email.
"""

import os
import time
import datetime
import smtplib
import logging
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from io import BytesIO

import requests
import yfinance as yf
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────────────────────────
#  CONFIGURACIÓN  (se leen desde variables de entorno en GitHub)
# ─────────────────────────────────────────────────────────────────
EMAIL_REMITENTE  = os.environ["EMAIL_REMITENTE"]   # tu Gmail
EMAIL_PASSWORD   = os.environ["EMAIL_PASSWORD"]    # contraseña de app Gmail
EMAIL_DESTINO    = os.environ["EMAIL_DESTINO"]     # donde recibes el Excel
EXCEL_PATH       = "fondos_inversion.xlsx"         # archivo en el repositorio

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

# ─────────────────────────────────────────────────────────────────
#  ISIN → TICKER de Yahoo Finance
#  (Yahoo no soporta ISIN directamente; mapeamos manualmente los
#   fondos que tienen ticker conocido; el resto se obtienen
#   vía Morningstar)
# ─────────────────────────────────────────────────────────────────
ISIN_A_TICKER = {
    "FR0010830885": "CSH2.PA",    # Amundi Enhanced Ultra ST Bond
    "FR0013399633": "C3M.PA",     # Amundi Euro Liquidity ST Govies
    "LU1882449801": "AFEM.PA",    # Amundi EM Fund
    "LU0906524193": "AGBD.PA",    # Amundi Global Corporate Bond
    "LU1706854152": "ASFD.PA",    # Amundi SF Diversified
    "FR0010149120": "CSM.PA",     # Carmignac Sécurité
    "LU0336084032": "CFPB.PA",    # Carmignac Portfolio Flexible Bond
    "LU0151324422": "CBCO.PA",    # Candriam Bonds Credit Opp.
    "LU0694789451": "DNCA.PA",    # DNCA Alpha Bonds
    "LU1161527038": "EDRA.PA",    # EdR Fund Bond Allocation
}

# ─────────────────────────────────────────────────────────────────
#  ISIN → ID Morningstar  (para obtener datos de ficha de fondo)
#  Se construye la URL: https://www.morningstar.es/es/funds/snapshot/snapshot.aspx?id=<ID>
# ─────────────────────────────────────────────────────────────────
ISIN_A_MORNINGSTAR = {
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
}

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "es-ES,es;q=0.9",
}

# ─────────────────────────────────────────────────────────────────
#  1.  NAV + RENTABILIDADES  vía Yahoo Finance
# ─────────────────────────────────────────────────────────────────
def obtener_datos_yahoo(ticker: str) -> dict:
    """Devuelve NAV, YTD%, 1M%, 3M%, 1Y% para un ticker de Yahoo."""
    try:
        t = yf.Ticker(ticker)
        hist = t.history(period="2y")
        if hist.empty:
            return {}

        precio_hoy = round(hist["Close"].iloc[-1], 4)
        hoy = hist.index[-1]

        def pct(dias_atras: int) -> float | None:
            fecha_ref = hoy - datetime.timedelta(days=dias_atras)
            pasado = hist[hist.index <= fecha_ref]
            if pasado.empty:
                return None
            p0 = pasado["Close"].iloc[-1]
            return round((precio_hoy / p0 - 1) * 100, 2)

        # YTD: desde 1 enero del año en curso
        inicio_anio = datetime.datetime(hoy.year, 1, 1)
        hist_ytd = hist[hist.index >= str(inicio_anio)]
        ytd = None
        if not hist_ytd.empty:
            p0_ytd = hist_ytd["Close"].iloc[0]
            ytd = round((precio_hoy / p0_ytd - 1) * 100, 2)

        return {
            "nav":   precio_hoy,
            "ytd":   ytd,
            "1m":    pct(30),
            "3m":    pct(90),
            "1y":    pct(365),
            "fecha": hoy.strftime("%d/%m/%Y"),
        }
    except Exception as e:
        log.warning(f"Yahoo {ticker}: {e}")
        return {}


# ─────────────────────────────────────────────────────────────────
#  2.  DATOS DE FICHA  vía Morningstar ES
#      (rating, duración, YTM)
# ─────────────────────────────────────────────────────────────────
def obtener_morningstar(ms_id: str) -> dict:
    """Scraping básico de la ficha de Morningstar España."""
    url = f"https://www.morningstar.es/es/funds/snapshot/snapshot.aspx?id={ms_id}"
    try:
        r = requests.get(url, headers=HEADERS, timeout=15)
        if r.status_code != 200:
            return {}
        soup = BeautifulSoup(r.text, "html.parser")

        datos = {}

        # Rating Morningstar (estrellas)
        rating_div = soup.find("span", {"class": "ratingValue"})
        if rating_div:
            datos["estrellas"] = rating_div.get_text(strip=True)

        # Buscar tabla de datos clave
        for row in soup.find_all("tr"):
            celdas = row.find_all("td")
            if len(celdas) >= 2:
                etiqueta = celdas[0].get_text(strip=True).lower()
                valor    = celdas[1].get_text(strip=True)
                if "duraci" in etiqueta:
                    datos["duracion"] = valor
                elif "ytm" in etiqueta or "rendimiento" in etiqueta:
                    datos["ytm"] = valor
                elif "rating" in etiqueta and "credit" in etiqueta:
                    datos["rating_credito"] = valor

        return datos
    except Exception as e:
        log.warning(f"Morningstar {ms_id}: {e}")
        return {}


# ─────────────────────────────────────────────────────────────────
#  3.  ACTUALIZAR EXCEL
# ─────────────────────────────────────────────────────────────────
# Mapeo de columnas en la hoja "Fondos - Datos Completos"
COL_ISIN        = 2
COL_DURACION    = 6
COL_YTM         = 7
COL_RATING      = 8
COL_RENT_2023   = 24   # Rentabilidad 2023
COL_YTD         = 25   # YTD 2026
COL_ESTRELLAS   = 34
COL_NAV         = 36   # Nueva columna que añadiremos
COL_1M          = 37
COL_3M          = 38
COL_1Y          = 39
COL_ACTUALIZ    = 40

AZUL_HDR  = "1F4E79"
NARANJA   = "F4B942"
AZUL_CL   = "D6E4F0"
BLANCO    = "FFFFFF"
thin = Border(
    left  =Side(style="thin", color="BFBFBF"),
    right =Side(style="thin", color="BFBFBF"),
    top   =Side(style="thin", color="BFBFBF"),
    bottom=Side(style="thin", color="BFBFBF"),
)

def estilo_dato(cell, bg=BLANCO, fmt=None, bold=False, align="center"):
    cell.font      = Font(name="Arial", size=8, bold=bold,
                          color="000000" if bg != AZUL_HDR else "FFFFFF")
    cell.fill      = PatternFill("solid", start_color=bg)
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
    cell.border    = thin
    if fmt:
        cell.number_format = fmt


def actualizar_excel() -> BytesIO:
    wb = load_workbook(EXCEL_PATH)
    ws = wb["Fondos - Datos Completos"]

    today = datetime.date.today().strftime("%d/%m/%Y")

    # ── Añadir cabeceras extra si no existen ──────────────────────
    nuevas_cols = {
        COL_NAV:     "NAV\nÚltimo",
        COL_1M:      "Rent.\n1 Mes",
        COL_3M:      "Rent.\n3 Meses",
        COL_1Y:      "Rent.\n1 Año",
        COL_ACTUALIZ:"Última\nActualiz.",
    }
    for col, titulo in nuevas_cols.items():
        cell = ws.cell(row=2, column=col, value=titulo)
        cell.font      = Font(bold=True, color="FFFFFF", size=8, name="Arial")
        cell.fill      = PatternFill("solid", start_color=AZUL_HDR)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = thin
        ws.column_dimensions[get_column_letter(col)].width = 10

    # ── Recorrer filas de datos (fila 3 en adelante) ──────────────
    for row in range(3, ws.max_row + 1):
        isin_cell = ws.cell(row=row, column=COL_ISIN)
        isin = isin_cell.value
        if not isin or str(isin).strip() == "":
            continue

        isin = str(isin).strip()
        bg_row = AZUL_CL if row % 2 == 1 else BLANCO
        log.info(f"  Procesando {isin} (fila {row})...")

        # — Yahoo Finance —
        ticker = ISIN_A_TICKER.get(isin)
        ydata  = obtener_datos_yahoo(ticker) if ticker else {}

        if ydata:
            # NAV
            c = ws.cell(row=row, column=COL_NAV, value=ydata.get("nav"))
            estilo_dato(c, bg=bg_row, fmt="#,##0.0000")
            # YTD
            ytd = ydata.get("ytd")
            if ytd is not None:
                c = ws.cell(row=row, column=COL_YTD, value=ytd / 100)
                estilo_dato(c, bg=bg_row, fmt="0.00%",
                            bold=(ytd > 5))   # negrita si >5%
            # 1M
            m1 = ydata.get("1m")
            if m1 is not None:
                c = ws.cell(row=row, column=COL_1M, value=m1 / 100)
                estilo_dato(c, bg=bg_row, fmt="0.00%")
            # 3M
            m3 = ydata.get("3m")
            if m3 is not None:
                c = ws.cell(row=row, column=COL_3M, value=m3 / 100)
                estilo_dato(c, bg=bg_row, fmt="0.00%")
            # 1Y
            y1 = ydata.get("1y")
            if y1 is not None:
                c = ws.cell(row=row, column=COL_1Y, value=y1 / 100)
                estilo_dato(c, bg=bg_row, fmt="0.00%")

        # — Morningstar —
        ms_id = ISIN_A_MORNINGSTAR.get(isin)
        if ms_id:
            mdata = obtener_morningstar(ms_id)
            time.sleep(1.5)   # pausa cortés para no saturar el servidor

            if mdata.get("duracion"):
                try:
                    dur = float(mdata["duracion"].replace(",", "."))
                    c = ws.cell(row=row, column=COL_DURACION, value=dur)
                    estilo_dato(c, bg=bg_row, fmt="0.00")
                except ValueError:
                    pass

            if mdata.get("ytm"):
                try:
                    ytm_str = mdata["ytm"].replace("%", "").replace(",", ".").strip()
                    ytm_val = float(ytm_str) / 100
                    c = ws.cell(row=row, column=COL_YTM, value=ytm_val)
                    estilo_dato(c, bg=bg_row, fmt="0.00%")
                except ValueError:
                    pass

            if mdata.get("rating_credito"):
                c = ws.cell(row=row, column=COL_RATING, value=mdata["rating_credito"])
                estilo_dato(c, bg=bg_row)

            if mdata.get("estrellas"):
                try:
                    n = int(mdata["estrellas"])
                    c = ws.cell(row=row, column=COL_ESTRELLAS, value="⭐" * n)
                    estilo_dato(c, bg=bg_row)
                except ValueError:
                    pass

        # — Fecha de actualización —
        c = ws.cell(row=row, column=COL_ACTUALIZ, value=today)
        estilo_dato(c, bg=bg_row)

    # ── Añadir pestaña de log de actualizaciones ──────────────────
    if "Historial" not in wb.sheetnames:
        wlog = wb.create_sheet("Historial")
        wlog.cell(1, 1, "Fecha").font      = Font(bold=True, name="Arial")
        wlog.cell(1, 2, "Fondos OK").font  = Font(bold=True, name="Arial")
        wlog.cell(1, 3, "Notas").font      = Font(bold=True, name="Arial")
    else:
        wlog = wb["Historial"]

    next_row = wlog.max_row + 1
    wlog.cell(next_row, 1, today)
    wlog.cell(next_row, 2, ws.max_row - 2)
    wlog.cell(next_row, 3, "Actualización automática diaria")

    # ── Guardar en memoria y devolver ─────────────────────────────
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    # También sobreescribir el archivo en disco (para el repositorio)
    wb.save(EXCEL_PATH)
    log.info("Excel guardado correctamente.")
    return buf


# ─────────────────────────────────────────────────────────────────
#  4.  ENVIAR EMAIL
# ─────────────────────────────────────────────────────────────────
def enviar_email(excel_bytes: BytesIO):
    today_str = datetime.date.today().strftime("%d/%m/%Y")
    nombre_archivo = f"fondos_inversion_{datetime.date.today():%Y%m%d}.xlsx"

    msg = MIMEMultipart()
    msg["From"]    = EMAIL_REMITENTE
    msg["To"]      = EMAIL_DESTINO
    msg["Subject"] = f"📊 Fondos de Inversión — Actualización {today_str}"

    cuerpo = f"""
    <html><body style="font-family:Arial,sans-serif;color:#333">
    <h2 style="color:#1F4E79">📊 Fondos de Inversión — {today_str}</h2>
    <p>Se adjunta el Excel con <strong>36 fondos</strong> actualizado automáticamente con:</p>
    <ul>
      <li>💰 <strong>NAV / Precio</strong> al cierre de ayer</li>
      <li>📈 <strong>Rentabilidad</strong> YTD, 1 mes, 3 meses, 1 año</li>
      <li>⭐ <strong>Rating Morningstar</strong> actualizado</li>
      <li>⏱️ <strong>Duración y YTM</strong> (cuando disponibles)</li>
    </ul>
    <p style="color:#888;font-size:12px">
      Fuentes: Yahoo Finance · Morningstar ES<br>
      Actualización automática vía GitHub Actions
    </p>
    </body></html>
    """

    msg.attach(MIMEText(cuerpo, "html"))

    parte = MIMEBase("application", "octet-stream")
    parte.set_payload(excel_bytes.read())
    encoders.encode_base64(parte)
    parte.add_header(
        "Content-Disposition",
        f'attachment; filename="{nombre_archivo}"',
    )
    msg.attach(parte)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(EMAIL_REMITENTE, EMAIL_PASSWORD)
        server.sendmail(EMAIL_REMITENTE, EMAIL_DESTINO, msg.as_string())

    log.info(f"Email enviado a {EMAIL_DESTINO}")


# ─────────────────────────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    log.info("═══ Inicio de actualización ═══")
    excel_buf = actualizar_excel()
    enviar_email(excel_buf)
    log.info("═══ Proceso completado ═══")
