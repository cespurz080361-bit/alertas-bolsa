"""
Alertas de precios de bolsa por Telegram — v2 adaptada a GitHub Actions.

Modos de funcionamiento:
  1) GitHub Actions (recomendado; funciona con el PC apagado):
       python alertas_bolsa_telegram.py --una-vez
     Ejecuta UN ciclo de vigilancia. La repetición cada 15 minutos la hace el
     cron del workflow (.github/workflows/alertas.yml) y el estado (precios de
     referencia) se guarda automáticamente en un commit del repositorio.

  2) En tu PC (como la versión anterior):
       python alertas_bolsa_telegram.py
     Bucle continuo consultando cada INTERVALO_MINUTOS mientras el PC esté
     encendido.

Requisitos:
    pip install yfinance pandas requests xlrd openpyxl

Credenciales de Telegram (por este orden de prioridad):
    a) Variables de entorno TELEGRAM_TOKEN y TELEGRAM_CHAT_ID.
       En GitHub se configuran como "Secrets" del repositorio (recomendado).
    b) Constantes pegadas más abajo (solo para uso local; NO subas al
       repositorio una versión con el token escrito).

Compatibilidad con OpenOffice:
    Se lee tanto .xlsx como el .xls clásico (mediante xlrd), de modo que
    puedes seguir editando alertas_bolsa.xls en OpenOffice y subiendo el
    mismo archivo al repositorio cuando cambies objetivos o activos.
    (La plantilla de reserva, si no existiera ningún Excel, se crea en
     .xlsx porque pandas ya no puede escribir formato .xls; en este
     proyecto no llega a ejecutarse nunca, ya que alertas_bolsa.xls
     viaja en el repositorio.)

Notas:
- yfinance no garantiza cotizaciones en tiempo real.
- La zona horaria para la ventana horaria y las horas mostradas es
  configurable (ZONA_HORARIA). Los servidores de GitHub van en UTC, por
  eso se fija Europe/Madrid por defecto.
"""

from __future__ import annotations

import html
import json
import logging
import os
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo

import pandas as pd
import requests
import yfinance as yf


# ---------------------- CONFIGURACIÓN ----------------------
INTERVALO_MINUTOS = 15  # Solo se usa en modo bucle (en GitHub programa el workflow).

# Datos de Telegram. Primero se intentan las variables de entorno (Secrets
# de GitHub); si no existen, se usan los valores escritos aquí (solo útil
# si ejecutas el script en tu propio PC).
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN", "")      # ej: "123456:ABC-DEF..."
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "")  # ej: "123456789"

# Zona horaria para la ventana horaria y para las horas de los mensajes.
ZONA_HORARIA = os.environ.get("TZ_ALERTAS", "Europe/Madrid")

HORA_INICIO = "00:00"  # Hora local (según ZONA_HORARIA), formato HH:MM
HORA_FIN = "23:59"     # Hora local (según ZONA_HORARIA), formato HH:MM

ENVIAR_AVISO_INICIO = False  # Mensaje de prueba al arrancar (solo modo bucle).
# ----------------------------------------------------------

BASE_DIR = Path(__file__).resolve().parent
EXCEL_XLSX = BASE_DIR / "alertas_bolsa.xlsx"  # Formato moderno (opcional)
EXCEL_XLS = BASE_DIR / "alertas_bolsa.xls"    # Formato OpenOffice (el de este proyecto)
ESTADO_PATH = BASE_DIR / "estado_alertas.json"
LOG_PATH = BASE_DIR / "alertas_bolsa.log"


def _excel_existente() -> Path:
    """Devuelve el Excel que exista (.xlsx preferente) o la ruta .xlsx por defecto."""
    if EXCEL_XLSX.exists():
        return EXCEL_XLSX
    if EXCEL_XLS.exists():
        return EXCEL_XLS
    return EXCEL_XLSX


EXCEL_PATH = _excel_existente()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[
        logging.FileHandler(LOG_PATH, encoding="utf-8"),
        logging.StreamHandler(),
    ],
)

COLUMNAS = ["Nombre", "Símbolo", "Objetivo 1", "Objetivo 2", "Activo"]


def _tz() -> Any:
    """Devuelve la zona horaria configurada (con reserva segura)."""
    try:
        return ZoneInfo(ZONA_HORARIA)
    except Exception:
        return datetime.now().astimezone().tzinfo


def ahora_local() -> datetime:
    """Fecha/hora actual en la zona horaria configurada."""
    return datetime.now(_tz())


def crear_excel_si_no_existe() -> None:
    """Crea una plantilla de reserva si no existe ningún Excel.

    Nota: la plantilla se genera en .xlsx (pandas ya no puede escribir el
    formato .xls). En este proyecto no debería ejecutarse nunca, porque
    alertas_bolsa.xls viaja en el repositorio.
    """
    if EXCEL_XLSX.exists() or EXCEL_XLS.exists():
        return

    df = pd.DataFrame(
        [
            {"Nombre": "Inditex (ejemplo)", "Símbolo": "ITX.MC", "Objetivo 1": 45, "Objetivo 2": 55, "Activo": "NO"},
            {"Nombre": "Apple (ejemplo)", "Símbolo": "AAPL", "Objetivo 1": 230, "Objetivo 2": 270, "Activo": "NO"},
        ],
        columns=COLUMNAS,
    )
    df.to_excel(EXCEL_XLSX, index=False)
    logging.info("Plantilla creada: %s", EXCEL_XLSX)


def cargar_estado() -> dict[str, Any]:
    if not ESTADO_PATH.exists():
        return {"precios_previos": {}, "alertas_pendientes": []}
    try:
        with ESTADO_PATH.open("r", encoding="utf-8") as f:
            data = json.load(f)
        data.setdefault("precios_previos", {})
        data.setdefault("alertas_pendientes", [])
        return data
    except (OSError, json.JSONDecodeError) as exc:
        logging.warning("No se pudo leer el estado (%s). Se inicia estado nuevo.", exc)
        return {"precios_previos": {}, "alertas_pendientes": []}


def guardar_estado(estado: dict[str, Any]) -> None:
    temp = ESTADO_PATH.with_suffix(".tmp")
    with temp.open("w", encoding="utf-8") as f:
        json.dump(estado, f, ensure_ascii=False, indent=2)
    temp.replace(ESTADO_PATH)


def hora_dentro_de_ventana() -> bool:
    ahora = ahora_local().strftime("%H:%M")
    inicio = HORA_INICIO
    fin = HORA_FIN
    if inicio <= fin:
        return inicio <= ahora <= fin
    return ahora >= inicio or ahora <= fin


def cargar_alertas() -> pd.DataFrame:
    crear_excel_si_no_existe()
    try:
        df = pd.read_excel(EXCEL_PATH)
    except Exception as exc:
        raise RuntimeError(
            f"No se pudo leer {EXCEL_PATH.name}: {exc}. "
            "El formato .xls requiere el paquete xlrd (incluido en requirements.txt). "
            "Comprueba también que el archivo no esté abierto/bloqueado por OpenOffice."
        ) from exc
    for col in COLUMNAS:
        if col not in df.columns:
            raise ValueError(f"Falta la columna '{col}' en {EXCEL_PATH.name}")
    df["Símbolo"] = df["Símbolo"].fillna("").astype(str).str.strip()
    df["Activo"] = df["Activo"].fillna("NO").astype(str).str.strip().str.upper()
    return df


def obtener_precio(symbol: str) -> tuple[float | None, str | None]:
    """
    Devuelve el último cierre intradía disponible y su hora.
    No garantiza que el dato sea en tiempo real; mercados cerrados pueden devolver
    la última vela disponible.
    """
    try:
        hist = yf.Ticker(symbol).history(period="5d", interval="1m", auto_adjust=False, actions=False)
        if hist is None or hist.empty or "Close" not in hist.columns:
            logging.warning("%s: Yahoo Finance no devolvió cotizaciones.", symbol)
            return None, None
        close = pd.to_numeric(hist["Close"], errors="coerce").dropna()
        if close.empty:
            return None, None
        timestamp = close.index[-1]
        price = float(close.iloc[-1])
        if price <= 0:
            return None, None
        timestamp_text = timestamp.strftime("%Y-%m-%d %H:%M %Z")
        return price, timestamp_text
    except Exception as exc:
        logging.warning("%s: error consultando cotización: %s", symbol, exc)
        return None, None


def detectar_cruces(
    nombre: str,
    symbol: str,
    precio_anterior: float | None,
    precio_actual: float,
    objetivos: list[float],
    hora_dato: str,
) -> list[dict[str, Any]]:
    """Detecta el cruce de cada objetivo en cualquiera de los dos sentidos."""
    if precio_anterior is None:
        return []

    eventos = []
    for objetivo in objetivos:
        if precio_anterior < objetivo <= precio_actual:
            direccion = "HA SUBIDO y ha cruzado al alza"
        elif precio_anterior > objetivo >= precio_actual:
            direccion = "HA BAJADO y ha cruzado a la baja"
        else:
            continue

        eventos.append({
            "nombre": nombre,
            "simbolo": symbol,
            "objetivo": objetivo,
            "anterior": precio_anterior,
            "actual": precio_actual,
            "direccion": direccion,
            "hora_dato": hora_dato,
            "creado": ahora_local().isoformat(timespec="seconds"),
        })
    return eventos


def enviar_telegram(mensaje: str) -> bool:
    """
    Envía un mensaje por Telegram con el Bot API oficial.
    Devuelve True si el envío se completó correctamente.
    """
    if not TELEGRAM_TOKEN or not TELEGRAM_CHAT_ID:
        logging.error(
            "Faltan TELEGRAM_TOKEN y/o TELEGRAM_CHAT_ID. En GitHub, crea los "
            "'Secrets' del repositorio con esos nombres. En local, exporta las "
            "variables de entorno o pégalos en la zona de CONFIGURACIÓN del script."
        )
        return False

    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    try:
        respuesta = requests.post(
            url,
            data={
                "chat_id": TELEGRAM_CHAT_ID,
                "text": mensaje,
                "parse_mode": "HTML",  # permite <b>negrita</b>, <code>, etc.
            },
            timeout=15,
        )
    except Exception as exc:
        logging.exception("No se pudo enviar el mensaje de Telegram: %s", exc)
        return False

    if respuesta.ok:
        logging.info("Telegram enviado correctamente.")
        return True

    logging.error(
        "Telegram rechazó el envío (HTTP %s): %s",
        respuesta.status_code,
        respuesta.text[:300],
    )
    return False


def enviar_pendientes(estado: dict[str, Any]) -> None:
    pendientes = estado.get("alertas_pendientes", [])
    while pendientes:
        evento = pendientes[0]
        mensaje = (
            f"🔔 <b>ALERTA DE BOLSA:</b> {html.escape(str(evento['nombre']))} "
            f"(<code>{html.escape(str(evento['simbolo']))}</code>)\n"
            f"{html.escape(str(evento['direccion']))} el objetivo <b>{evento['objetivo']:,.4f}</b>\n"
            f"Precio observado: {evento['actual']:,.4f}\n"
            f"Precio anterior: {evento['anterior']:,.4f}\n"
            f"Hora del dato: {html.escape(str(evento['hora_dato']))}\n"
            f"Consulta: {ahora_local().strftime('%d/%m/%Y %H:%M')}"
        )
        if enviar_telegram(mensaje):
            pendientes.pop(0)
            guardar_estado(estado)
            # Pequeña pausa entre mensajes por cortesía con la API.
            if pendientes:
                time.sleep(1)
        else:
            # Conserva el evento para reintentarlo en la siguiente vuelta.
            break


def ciclo(estado: dict[str, Any]) -> None:
    df = cargar_alertas()
    activos = df[df["Activo"].isin(["SI", "SÍ", "YES", "TRUE", "1"])]

    if activos.empty:
        logging.info("No hay activos habilitados en el Excel (Activo = SI).")
        return

    for _, row in activos.iterrows():
        symbol = str(row["Símbolo"]).strip()
        if not symbol:
            continue
        nombre = str(row["Nombre"]).strip()
        if not nombre or nombre.lower() == "nan":
            nombre = symbol

        objetivos = []
        for col in ("Objetivo 1", "Objetivo 2"):
            value = row[col]
            if pd.notna(value):
                try:
                    objetivo = float(value)
                    if objetivo > 0 and objetivo not in objetivos:
                        objetivos.append(objetivo)
                except (TypeError, ValueError):
                    logging.warning("%s: objetivo no válido en %s.", symbol, col)

        if not objetivos:
            logging.info("%s: sin objetivos configurados; se omite.", symbol)
            continue

        precio, hora_dato = obtener_precio(symbol)
        if precio is None:
            continue

        key = symbol.upper()
        precio_anterior = estado["precios_previos"].get(key)

        # El primer precio se usa como referencia; no dispara alertas al arrancar.
        if precio_anterior is None:
            logging.info("%s (%s): precio inicial %.6f; se establece referencia.", nombre, symbol, precio)
        else:
            eventos = detectar_cruces(
                nombre, symbol, float(precio_anterior), precio, objetivos, hora_dato or "desconocida"
            )
            if eventos:
                estado["alertas_pendientes"].extend(eventos)
                logging.info("%s: detectados %s cruce(s).", symbol, len(eventos))

        estado["precios_previos"][key] = precio
        logging.info("%s (%s): %.6f | dato: %s", nombre, symbol, precio, hora_dato)

        # Guardar tras cada activo para conservar avances si el programa se interrumpe.
        guardar_estado(estado)
        time.sleep(1)

    enviar_pendientes(estado)


def main() -> None:
    una_vez = "--una-vez" in sys.argv[1:] or os.environ.get("EJECUCION_UNICA", "") == "1"

    crear_excel_si_no_existe()
    logging.info("=== Alertas de bolsa (vía Telegram) ===")
    logging.info("Excel: %s", EXCEL_PATH)
    logging.info("Zona horaria: %s", ZONA_HORARIA)

    if una_vez:
        # Modo GitHub Actions: un ciclo y fuera; el estado queda guardado
        # en estado_alertas.json y el workflow lo sube con un commit.
        logging.info("Modo: un solo ciclo (--una-vez).")
        estado = cargar_estado()
        if hora_dentro_de_ventana():
            ciclo(estado)
        else:
            logging.info("Fuera de la ventana horaria configurada; no se consultan precios.")
        return

    # Modo PC: bucle continuo como la versión anterior.
    logging.info("Frecuencia: cada %s minutos.", INTERVALO_MINUTOS)
    logging.info("Para detener, pulsa Ctrl+C en esta ventana.")

    if ENVIAR_AVISO_INICIO:
        enviar_telegram(
            "🚀 <b>Bot de alertas de bolsa iniciado</b>\n"
            f"Frecuencia de consulta: cada {INTERVALO_MINUTOS} min\n"
            f"Ventana horaria: {HORA_INICIO} - {HORA_FIN}\n"
            "Te avisaré cuando un precio cruce uno de sus objetivos."
        )

    estado = cargar_estado()

    while True:
        inicio = time.monotonic()
        try:
            if hora_dentro_de_ventana():
                ciclo(estado)
            else:
                logging.info("Fuera de la ventana horaria configurada; no se consultan precios.")
        except KeyboardInterrupt:
            logging.info("Programa detenido por el usuario.")
            break
        except Exception:
            logging.exception("Error inesperado en el ciclo.")

        transcurrido = time.monotonic() - inicio
        espera = max(1, INTERVALO_MINUTOS * 60 - int(transcurrido))
        proxima = ahora_local() + timedelta(seconds=espera)
        logging.info("Próxima consulta aproximadamente a las %s", proxima.strftime("%H:%M:%S"))
        try:
            time.sleep(espera)
        except KeyboardInterrupt:
            logging.info("Programa detenido por el usuario.")
            break


if __name__ == "__main__":
    main()
