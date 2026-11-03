from __future__ import annotations

import json
import re
import sys
import threading
from datetime import date, datetime, time, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

import pandas as pd
import oracledb

APP_DIR = Path.home() / "UPC_Disponibilidad_Aulas"
APP_DIR.mkdir(parents=True, exist_ok=True)
CONFIG_PATH = APP_DIR / "config.json"
DEFAULT_DSN = "SVRBNRBDPROD3:1521/UPCSTB"
KEYRING_SERVICE = "UPC_Disponibilidad_Aulas_Oracle"
DEFAULT_PERIODS = ["202701", "202705", "202710", "202715", "202720", "202725", "202711", "202790"]
SQL = """
SELECT PERIODO_NRC, NRC,
       TRIM(MATERIA)||TRIM(CURSO) AS COD_CURSO,
       TITULO_CURSO, NOMBRE_CURSO_LARGO, CAMPUS, DES_CAMPUS,
       METODO_EDUCATIVO_NRC, COD_TIPO_REUNION,
       FECHA_INICIO_SESION, FECHA_FIN_SESION,
       LUN, MAR, MIE, JUE, VIE, SAB, DOM,
       HORA_INICIO_SESION, HORA_FIN_SESION,
       COD_EDIFICIO, COD_SALON, COD_ATRIBUTO_SALON,
       DES_ATRIBUTO_SALON, TIPO_HORARIO, HORAS_SEMANA,
       INDICADOR_SESION
FROM UPC_REPORTES2.V_HORARIO_DOCENTE
WHERE PERIODO_NRC IN ({binds})
ORDER BY COD_EDIFICIO, COD_SALON, FECHA_INICIO_SESION, HORA_INICIO_SESION, NRC
"""
DAY_COLS = [("LUN", 0, "Lunes"), ("MAR", 1, "Martes"), ("MIE", 2, "Miércoles"),
            ("JUE", 3, "Jueves"), ("VIE", 4, "Viernes"), ("SAB", 5, "Sábado"),
            ("DOM", 6, "Domingo")]
PHYSICAL_TYPES = {"CLAS", "CLAH"}
START_MIN = 7 * 60
END_MIN = 24 * 60  # ventana hasta 23:59, representada internamente como 24:00


def resource_path(filename):
    """Obtiene recursos tanto al ejecutar el .py como desde un .exe de PyInstaller."""
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / filename


def apply_upc_icon(window):
    """Aplica el icono institucional UPC a una ventana Tkinter, si está disponible."""
    try:
        icon_path = resource_path("imagen_2026-10-27_232443008.ico")
        if icon_path.is_file():
            window.iconbitmap(str(icon_path))
    except (tk.TclError, OSError):
        # La aplicación debe poder abrir incluso si el recurso no está disponible.
        pass


def load_config():
    try:
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_config(config):
    CONFIG_PATH.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")


def get_saved_password(username):
    """Retrieve a password from the operating-system credential vault, if available."""
    if not username:
        return None
    try:
        import keyring
        return keyring.get_password(KEYRING_SERVICE, username)
    except Exception:
        return None


def save_password(username, password):
    """Save a password using the configured OS credential backend."""
    import keyring
    keyring.set_password(KEYRING_SERVICE, username, password)


def delete_saved_password(username):
    if not username:
        return
    try:
        import keyring
        keyring.delete_password(KEYRING_SERVICE, username)
    except Exception:
        pass


def norm(value):
    if pd.isna(value):
        return ""
    return str(value).strip().upper()


def parse_time(value):
    if isinstance(value, time):
        return value.hour * 60 + value.minute
    if isinstance(value, (datetime, pd.Timestamp)):
        return value.hour * 60 + value.minute
    if isinstance(value, (int, float)) and not pd.isna(value):
        # Excel puede almacenar las horas como fracciones de un día.
        if 0 <= float(value) < 1:
            return int(round(float(value) * 24 * 60))
    s = norm(value)
    if not s or s in {":", "NAT", "NAN", "NONE"}:
        return None
    m = re.fullmatch(r"(\d{1,2}):(\d{2})(?::\d{2})?", s)
    if not m:
        return None
    h, minute = int(m.group(1)), int(m.group(2))
    if h > 23 or minute > 59:
        return None
    return h * 60 + minute


def minute_label(m):
    m = max(0, min(24 * 60, int(m)))
    if m == 24 * 60:
        return "23:59"
    return f"{m // 60:02d}:{m % 60:02d}"


def interval_label(start, end_exclusive):
    # El límite final se muestra como el último minuto ocupado/libre.
    return f"{minute_label(start)}–{minute_label(end_exclusive - 1)}"


def merge_intervals(intervals):
    if not intervals:
        return []
    intervals = sorted((max(START_MIN, s), min(END_MIN, e)) for s, e in intervals if e > s)
    out = []
    for s, e in intervals:
        if e <= s:
            continue
        if out and s <= out[-1][1]:
            out[-1] = (out[-1][0], max(out[-1][1], e))
        else:
            out.append((s, e))
    return out


def subtract_intervals(occupied):
    occupied = merge_intervals(occupied)
    free = []
    cursor = START_MIN
    for s, e in occupied:
        if s > cursor:
            free.append((cursor, s))
        cursor = max(cursor, e)
    if cursor < END_MIN:
        free.append((cursor, END_MIN))
    return free


def intersect_intervals(left, right):
    """Devuelve los tramos comunes entre dos listas de intervalos."""
    left = merge_intervals(left)
    right = merge_intervals(right)
    out = []
    i = j = 0
    while i < len(left) and j < len(right):
        s = max(left[i][0], right[j][0])
        e = min(left[i][1], right[j][1])
        if s < e:
            out.append((s, e))
        if left[i][1] <= right[j][1]:
            i += 1
        else:
            j += 1
    return merge_intervals(out)


def common_availability(patterns):
    """Disponibilidad que se mantiene en todas las fechas evaluadas."""
    if not patterns:
        return []
    common = list(patterns[0])
    for pattern in patterns[1:]:
        common = intersect_intervals(common, pattern)
        if not common:
            break
    return common


def parse_date(v):
    if pd.isna(v) or v == "":
        return None
    try:
        return pd.to_datetime(v, errors="coerce", dayfirst=True).date()
    except Exception:
        return None


def normalize_weekday(value):
    """Convierte nombres de días en español a índices Monday=0 ... Sunday=6."""
    day = norm(value)
    day = (day.replace("Á", "A").replace("É", "E").replace("Í", "I")
              .replace("Ó", "O").replace("Ú", "U"))
    aliases = {
        "LUNES": 0, "LUN": 0,
        "MARTES": 1, "MAR": 1,
        "MIERCOLES": 2, "MIER": 2, "MIE": 2,
        "JUEVES": 3, "JUE": 3,
        "VIERNES": 4, "VIE": 4,
        "SABADO": 5, "SAB": 5,
        "DOMINGO": 6, "DOM": 6,
    }
    return aliases.get(day)


def read_unavailable(path):
    """Lee y valida el Excel de aulas no disponibles."""
    df = pd.read_excel(path, sheet_name=0, dtype=object)
    df.columns = [str(c).strip().upper() for c in df.columns]
    required = ["COD_AULA", "FECHA_INICIO", "FECHA_FIN", "NRO_DIA",
                "HORA_INICIO", "HORA_FIN"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError("Al Excel de aulas no disponibles le faltan columnas: " + ", ".join(missing))
    df = df[required + (["CAMPUS"] if "CAMPUS" in df.columns else []) +
            (["TIPO_AULA"] if "TIPO_AULA" in df.columns else []) +
            (["ESTADO"] if "ESTADO" in df.columns else [])].copy()
    df["COD_AULA"] = df["COD_AULA"].map(norm)
    df["FECHA_INICIO_PARSED"] = df["FECHA_INICIO"].map(parse_date)
    df["FECHA_FIN_PARSED"] = df["FECHA_FIN"].map(parse_date)
    df["DIA_IDX"] = df["NRO_DIA"].map(normalize_weekday)
    df["HORA_INICIO_MIN"] = df["HORA_INICIO"].map(parse_time)
    df["HORA_FIN_MIN"] = df["HORA_FIN"].map(parse_time)
    return df


def parse_ui_date(value):
    """Lee fechas escritas por el usuario en formato DD/MM/AAAA."""
    return datetime.strptime(value.strip(), "%d/%m/%Y").date()


def display_config_date(value, fallback):
    """Convierte fechas ISO guardadas en config.json al formato visible DD/MM/AAAA."""
    try:
        return datetime.strptime(str(value), "%Y-%m-%d").strftime("%d/%m/%Y")
    except (TypeError, ValueError):
        return fallback.strftime("%d/%m/%Y")


def read_master(path):
    df = pd.read_excel(path, sheet_name=0, dtype=object)
    df.columns = [str(c).strip().upper() for c in df.columns]
    required = ["DESCRIPCION LOCAL", "PABELLON", "CODIGO AMBIENTE",
                "ATRIBUTO SALON", "DESCRIPCION ATRIBUTO SALON", "CAPACIDAD"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError("Al Excel maestro le faltan columnas: " + ", ".join(missing))
    df = df[required].copy()
    df = df[df["CODIGO AMBIENTE"].notna() & df["PABELLON"].notna()].copy()
    for c in ["DESCRIPCION LOCAL", "PABELLON", "CODIGO AMBIENTE", "ATRIBUTO SALON",
              "DESCRIPCION ATRIBUTO SALON"]:
        df[c] = df[c].map(norm)
    # Normaliza capacidades enteras (40.0 -> 40) sin alterar valores
    # que realmente tengan una parte decimal.
    capacidad_num = pd.to_numeric(df["CAPACIDAD"], errors="coerce")
    df["CAPACIDAD"] = capacidad_num.map(
        lambda v: int(v) if pd.notna(v) and float(v).is_integer()
        else (float(v) if pd.notna(v) else None)
    ).astype(object)
    df["CLAVE_AULA"] = df["PABELLON"] + "|" + df["CODIGO AMBIENTE"]
    # No eliminar duplicados sin avisar: puede haber más de un atributo registrado por ambiente.
    return df


def query_oracle(user, password, dsn, periods):
    try:
        import oracledb
    except ImportError as e:
        raise RuntimeError("Falta instalar oracledb. Ejecuta: python -m pip install oracledb") from e
    binds = ", ".join(f":p{i}" for i in range(len(periods)))
    sql = SQL.format(binds=binds)
    conn = oracledb.connect(user=user, password=password, dsn=dsn)
    try:
        with conn.cursor() as cursor:
            cursor.execute(sql, {f"p{i}": p for i, p in enumerate(periods)})
            columns = [item[0].strip().upper() for item in cursor.description]
            rows = cursor.fetchall()
        df = pd.DataFrame(rows, columns=columns)
    finally:
        conn.close()
    return df


def normalize_sessions(df):
    df = df.copy()
    df.columns = [str(c).strip().upper() for c in df.columns]
    for c in ["PERIODO_NRC", "NRC", "COD_TIPO_REUNION", "LUN", "MAR", "MIE", "JUE", "VIE",
              "SAB", "DOM", "HORA_INICIO_SESION", "HORA_FIN_SESION", "COD_EDIFICIO",
              "COD_SALON", "COD_ATRIBUTO_SALON", "INDICADOR_SESION"]:
        if c in df:
            df[c] = df[c].map(norm)
    for c in ["FECHA_INICIO_SESION", "FECHA_FIN_SESION"]:
        df[c] = pd.to_datetime(df[c], errors="coerce").dt.date
    return df


def run_analysis(sessions, master, start_date, end_date, filter_mode, filter_value, selected_days=None, selected_locales=None, time_start=None, time_end=None, selected_capacities=None, unavailable=None):
    # Si no se seleccionan casillas, se mantiene el comportamiento actual: todos los días.
    if not selected_days:
        selected_days = {i for _, i, _ in DAY_COLS}
    else:
        selected_days = set(selected_days)
    sessions = normalize_sessions(sessions)
    alerts = []
    for _, row in master.iterrows():
        # duplicate room keys can create ambiguous attribute/capacity mapping
        pass

    master_keys = set(master["CLAVE_AULA"])
    room_meta = {}
    for _, r in master.iterrows():
        key = r["CLAVE_AULA"]
        room_meta.setdefault(key, r.to_dict())

    # Apply view filters to master rooms; room-only uses room code and optional multi-attribute list.
    selected_master = master.copy()
    if filter_mode == "Aula":
        wanted = {norm(x) for x in filter_value.split(",") if norm(x)}
        selected_master = selected_master[selected_master["CODIGO AMBIENTE"].isin(wanted)]
    elif filter_mode == "Atributo":
        wanted = {norm(x) for x in filter_value.split(",") if norm(x)}
        selected_master = selected_master[selected_master["ATRIBUTO SALON"].isin(wanted)]
    # Optional view filter: local values come from the Excel master, not a hardcoded list.
    if selected_locales:
        selected_locales_norm = {norm(x) for x in selected_locales if norm(x)}
        selected_master = selected_master[selected_master["DESCRIPCION LOCAL"].isin(selected_locales_norm)]
    if selected_capacities:
        selected_capacity_values = {float(x) for x in selected_capacities}
        selected_master = selected_master[selected_master["CAPACIDAD"].isin(selected_capacity_values)]
    selected_keys = set(selected_master["CLAVE_AULA"])
    if not selected_keys:
        raise ValueError("El filtro no encontró aulas en el Excel maestro.")

    # Flag data-quality issues on scheduled sessions. Only physical CLAS/CLAH blocks occupancy.
    occupancy = {}  # (room key, date) -> [(start, end_exclusive)]
    physical_with_room = set()
    for _, r in sessions.iterrows():
        typ = norm(r.get("COD_TIPO_REUNION", ""))
        building, room = norm(r.get("COD_EDIFICIO", "")), norm(r.get("COD_SALON", ""))
        has_room = bool(building and room)
        if typ == "VI" and has_room and f"{building}|{room}" in selected_keys:
            alerts.append({"Tipo de alerta": "VI con aula física asignada",
                           "Periodo": r.get("PERIODO_NRC", ""), "NRC": r.get("NRC", ""),
                           "Aula Oracle": f"{building} / {room}",
                           "Detalle": "Registro virtual con pabellón/aula informado; se excluyó de la ocupación física."})
        if typ not in PHYSICAL_TYPES:
            continue
        d1, d2 = r.get("FECHA_INICIO_SESION"), r.get("FECHA_FIN_SESION")
        if not isinstance(d1, date) or not isinstance(d2, date):
            alerts.append({"Tipo de alerta": "Fecha incompleta", "Periodo": r.get("PERIODO_NRC", ""),
                           "NRC": r.get("NRC", ""), "Aula Oracle": f"{building} / {room}",
                           "Detalle": "Sesión física sin fecha válida; no se asumió ocupación."})
            continue
        if d2 < d1:
            alerts.append({"Tipo de alerta": "Rango de fechas inválido", "Periodo": r.get("PERIODO_NRC", ""),
                           "NRC": r.get("NRC", ""), "Aula Oracle": f"{building} / {room}",
                           "Detalle": "La fecha fin es anterior a la fecha inicio."})
            continue
        days = [i for col, i, _ in DAY_COLS if norm(r.get(col, ""))]
        if not days:
            alerts.append({"Tipo de alerta": "Día sin marcar", "Periodo": r.get("PERIODO_NRC", ""),
                           "NRC": r.get("NRC", ""), "Aula Oracle": f"{building} / {room}",
                           "Detalle": "Sesión física sin día de semana marcado; no se asumió ocupación."})
            continue
        st, en = parse_time(r.get("HORA_INICIO_SESION")), parse_time(r.get("HORA_FIN_SESION"))
        if st is None or en is None or en < st:
            alerts.append({"Tipo de alerta": "Horario incompleto/inválido", "Periodo": r.get("PERIODO_NRC", ""),
                           "NRC": r.get("NRC", ""), "Aula Oracle": f"{building} / {room}",
                           "Detalle": f"Inicio={r.get('HORA_INICIO_SESION', '')}; fin={r.get('HORA_FIN_SESION', '')}. No se asumió ocupación."})
            continue
        if not has_room:
            alerts.append({"Tipo de alerta": "Sesión física sin aula", "Periodo": r.get("PERIODO_NRC", ""),
                           "NRC": r.get("NRC", ""), "Aula Oracle": "",
                           "Detalle": "Registro CLAS/CLAH con fecha, día y horario, pero sin pabellón/aula."})
            continue
        key = f"{building}|{room}"
        physical_with_room.add(key)
        if key not in master_keys:
            alerts.append({"Tipo de alerta": "Aula Oracle no encontrada en maestro",
                           "Periodo": r.get("PERIODO_NRC", ""), "NRC": r.get("NRC", ""),
                           "Aula Oracle": f"{building} / {room}",
                           "Detalle": "No existe combinación PABELLON + CODIGO AMBIENTE en el Excel maestro."})
        current = max(d1, start_date)
        last = min(d2, end_date)
        if current > last:
            continue
        while current <= last:
            if current.weekday() in days:
                occupancy.setdefault((key, current), []).append((st, en + 1))
            current += timedelta(days=1)

    # Add external room blocks from the optional Excel. These intervals are
    # combined with Oracle occupancy, so both sources reduce free availability.
    if unavailable is not None and not unavailable.empty:
        selected_room_codes = set(selected_master["CODIGO AMBIENTE"].map(norm))
        for idx, block_row in unavailable.iterrows():
            room_code = norm(block_row.get("COD_AULA", ""))
            if not room_code or room_code not in selected_room_codes:
                continue
            d1 = block_row.get("FECHA_INICIO_PARSED")
            d2 = block_row.get("FECHA_FIN_PARSED")
            weekday = block_row.get("DIA_IDX")
            st = block_row.get("HORA_INICIO_MIN")
            en = block_row.get("HORA_FIN_MIN")
            if not isinstance(d1, date) or not isinstance(d2, date) or d2 < d1:
                alerts.append({"Tipo de alerta": "Bloqueo Excel con fechas inválidas", "Periodo": "",
                               "NRC": "", "Aula Oracle": room_code,
                               "Detalle": f"Fila {idx + 2}: revise FECHA_INICIO y FECHA_FIN."})
                continue
            if weekday is None or pd.isna(weekday):
                alerts.append({"Tipo de alerta": "Bloqueo Excel con día inválido", "Periodo": "",
                               "NRC": "", "Aula Oracle": room_code,
                               "Detalle": f"Fila {idx + 2}: NRO_DIA no corresponde a un día reconocido."})
                continue
            if st is None or en is None or en < st:
                alerts.append({"Tipo de alerta": "Bloqueo Excel con horario inválido", "Periodo": "",
                               "NRC": "", "Aula Oracle": room_code,
                               "Detalle": f"Fila {idx + 2}: revise HORA_INICIO y HORA_FIN."})
                continue
            # COD_AULA coincide con CODIGO AMBIENTE. Si el código se repite
            # en más de un pabellón/local, el bloqueo se aplica a esas coincidencias.
            matching_keys = set(selected_master.loc[
                selected_master["CODIGO AMBIENTE"].eq(room_code), "CLAVE_AULA"].tolist())
            current = max(d1, start_date)
            last = min(d2, end_date)
            while current <= last:
                if current.weekday() == int(weekday) and current.weekday() in selected_days:
                    for room_key in matching_keys:
                        occupancy.setdefault((room_key, current), []).append((int(st), int(en) + 1))
                current += timedelta(days=1)

    # Availability for each date and each selected room.
    daily = []
    dates = []
    d = start_date
    while d <= end_date:
        if d.weekday() in selected_days:
            dates.append(d)
        d += timedelta(days=1)
    for _, roomrow in selected_master.iterrows():
        key = roomrow["CLAVE_AULA"]
        for d in dates:
            occ = occupancy.get((key, d), [])
            free = subtract_intervals(occ)
            free = [(s, e) for s, e in free if e > s]
            # Optional display filter: intersect free intervals with the requested time window.
            # The underlying occupancy calculation above remains unchanged.
            if time_start is not None and time_end is not None:
                window_start, window_end = time_start, time_end + 1
                clipped = []
                for free_start, free_end in free:
                    clipped_start = max(free_start, window_start)
                    clipped_end = min(free_end, window_end)
                    if clipped_start < clipped_end:
                        clipped.append((clipped_start, clipped_end))
                free = clipped
            intervals = " | ".join(interval_label(s, e) for s, e in free) if free else "Sin disponibilidad"
            daily.append({
                "CLAVE_AULA": key,
                "Local": roomrow["DESCRIPCION LOCAL"],
                "Pabellón": roomrow["PABELLON"],
                "Aula": roomrow["CODIGO AMBIENTE"],
                "Atributo": roomrow["ATRIBUTO SALON"],
                "Descripción atributo": roomrow["DESCRIPCION ATRIBUTO SALON"],
                "Capacidad": roomrow["CAPACIDAD"],
                "Día": DAY_COLS[d.weekday()][2],
                "Fecha": d,
                "Patrón": tuple(free),
                "Disponibilidad": intervals,
                "NRC ocupantes": ", ".join(sorted({str(sessions.iloc[i].get("NRC", "")) for i in []})),
            })
    daily_df = pd.DataFrame(daily)

    # Compress only consecutive dates with the same weekday and exact availability pattern.
    details = []
    # Agrupar primero por aula y día de semana; así se pueden unir lunes consecutivos
    # (cada 7 días) sin que martes/miércoles interrumpan la secuencia.
    for (key, day_name), grp in daily_df.groupby(["CLAVE_AULA", "Día"], sort=True):
        grp = grp.sort_values("Fecha")
        run = None
        for _, row in grp.iterrows():
            sig = row["Patrón"]
            if run and sig == run["sig"] and row["Fecha"] == run["hasta"] + timedelta(days=7):
                run["hasta"] = row["Fecha"]
            else:
                if run:
                    details.append(run)
                run = {"sig": sig, "desde": row["Fecha"], "hasta": row["Fecha"],
                       "Local": row["Local"], "Pabellón": row["Pabellón"], "Aula": row["Aula"],
                       "Atributo": row["Atributo"], "Descripción atributo": row["Descripción atributo"],
                       "Capacidad": row["Capacidad"], "Día": row["Día"],
                       "Disponibilidad": row["Disponibilidad"]}
        if run:
            details.append(run)
    detail_df = pd.DataFrame([{
        "Local": x["Local"], "Pabellón": x["Pabellón"], "Aula": x["Aula"],
        "Atributo": x["Atributo"], "Descripción atributo": x["Descripción atributo"],
        "Capacidad": x["Capacidad"], "Día": x["Día"], "Fecha desde": x["desde"],
        "Fecha hasta": x["hasta"], "Intervalos disponibles": x["Disponibilidad"]
    } for x in details])
    if not detail_df.empty:
        # Orden cronológico real de los días: lunes (0) a domingo (6).
        weekday_order = {name: idx for idx, (_, idx, name) in enumerate(DAY_COLS)}
        detail_df["_ORDEN_DIA"] = detail_df["Día"].map(weekday_order)
        detail_df.sort_values(
            ["Local", "Pabellón", "Aula", "_ORDEN_DIA", "Fecha desde"],
            kind="stable", inplace=True
        )
        detail_df.drop(columns=["_ORDEN_DIA"], inplace=True)
        detail_df.reset_index(drop=True, inplace=True)

    # Summary: one row per room, with a compact availability description by weekday.
    summary = []
    for key, grp in daily_df.groupby("CLAVE_AULA", sort=True):
        first = grp.iloc[0]
        patterns = []
        for col, weekday_idx, day_name in DAY_COLS:
            if weekday_idx not in selected_days:
                continue
            daygrp = grp[grp["Día"] == day_name].sort_values("Fecha")
            if daygrp.empty:
                patterns.append(f"{day_name}: Sin fechas")
                continue

            # Mostrar solo los horarios disponibles en TODAS las fechas de ese día
            # dentro del rango elegido. Así no se mezcla una semana previa al inicio
            # de clases (libre todo el día) con el patrón de las semanas lectivas.
            common = common_availability(daygrp["Patrón"].tolist())
            common_text = " | ".join(interval_label(s, e) for s, e in common) if common else "Sin disponibilidad común"
            patterns.append(f"{day_name}: {common_text}")
        summary.append({
            "Local": first["Local"], "Pabellón": first["Pabellón"], "Aula": first["Aula"],
            "Atributo": first["Atributo"], "Descripción atributo": first["Descripción atributo"],
            "Capacidad": first["Capacidad"],
            "Fechas analizadas": f"{start_date:%d/%m/%Y} al {end_date:%d/%m/%Y}",
            "Patrón de disponibilidad por día": " || ".join(patterns)
        })
    summary_df = pd.DataFrame(summary)

    # Filter detail outputs only by selected master keys already; preserve full raw rows for traceability.
    raw = sessions.copy()
    if not raw.empty:
        raw["CLAVE_AULA"] = raw["COD_EDIFICIO"].map(norm) + "|" + raw["COD_SALON"].map(norm)
        raw = raw[raw["CLAVE_AULA"].isin(selected_keys)]
        raw.drop(columns=["CLAVE_AULA"], inplace=True, errors="ignore")
    alert_df = pd.DataFrame(alerts, columns=["Tipo de alerta", "Periodo", "NRC", "Aula Oracle", "Detalle"])
    return summary_df, detail_df, alert_df, raw


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("UPC | Disponibilidad de ambientes académicos")
        # Icono institucional UPC para la ventana principal.
        apply_upc_icon(self)
        # Ventana más compacta para evitar espacio vacío debajo de los controles.
        self.geometry("1080x720")
        self.minsize(920, 700)
        self.configure(bg="#F2F5F9")
        self.cfg = load_config()
        self._configure_styles()
        self._build()
        if self.master_var.get().strip():
            self.refresh_local_options()

    def _configure_styles(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        self.option_add("*Font", ("Segoe UI", 10))
        style.configure("TFrame", background="#F2F5F9")
        style.configure("Card.TLabelframe", background="#FFFFFF", bordercolor="#D9E1EB", relief="solid", borderwidth=1)
        style.configure("Card.TLabelframe.Label", background="#FFFFFF", foreground="#183153", font=("Segoe UI", 10, "bold"))
        style.configure("TLabel", background="#F2F5F9", foreground="#26364A")
        style.configure("Card.TLabel", background="#FFFFFF", foreground="#26364A")
        style.configure("Muted.TLabel", background="#FFFFFF", foreground="#68778A", font=("Segoe UI", 9))
        style.configure("TEntry", fieldbackground="#FFFFFF", foreground="#1E293B", padding=(7, 5), bordercolor="#C9D4E2")
        style.configure("TCheckbutton", background="#FFFFFF", foreground="#26364A", padding=2)
        style.map("TCheckbutton", background=[("active", "#FFFFFF")])
        style.configure("TRadiobutton", background="#FFFFFF", foreground="#26364A", padding=3)
        style.map("TRadiobutton", background=[("active", "#FFFFFF")])
        style.configure("TButton", padding=(10, 6), background="#E8EEF5", foreground="#183153", bordercolor="#D0DAE6")
        style.map("TButton", background=[("active", "#DCE6F1"), ("pressed", "#D0DDEA")])
        style.configure("Accent.TButton", padding=(16, 9), background="#174A7E", foreground="#FFFFFF", bordercolor="#174A7E", font=("Segoe UI", 10, "bold"))
        style.map("Accent.TButton", background=[("active", "#123B65"), ("pressed", "#0F3153")], foreground=[("active", "#FFFFFF")])

    def _build(self):
        root = ttk.Frame(self, padding=(12, 8, 12, 7))
        root.pack(fill="both", expand=True)
        root.columnconfigure(0, weight=1)
        root.columnconfigure(1, weight=1)

        # Encabezado corporativo
        header = tk.Frame(root, bg="#17365D", height=66)
        header.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 7))
        header.grid_propagate(False)
        tk.Frame(header, bg="#3C8DCE", width=6).pack(side="left", fill="y")
        header_text = tk.Frame(header, bg="#17365D")
        header_text.pack(side="left", fill="both", expand=True, padx=16, pady=8)
        tk.Label(header_text, text="Disponibilidad de ambientes académicos", bg="#17365D", fg="#FFFFFF",
                 font=("Segoe UI", 16, "bold"), anchor="w").pack(anchor="w")
        tk.Label(header_text, text="PROGRAMACIÓN HORARIA  |  Consulta y generación de reportes", bg="#17365D",
                 fg="#D6E3F2", font=("Segoe UI", 9), anchor="w").pack(anchor="w", pady=(3, 0))

        # Conexión
        db = ttk.LabelFrame(root, text="  Conexión Oracle  ", style="Card.TLabelframe", padding=(9, 6))
        db.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 6))
        db.columnconfigure(1, weight=1)
        db.columnconfigure(3, weight=2)
        self.user_var = tk.StringVar(value=self.cfg.get("oracle_user", ""))
        self.dsn_var = tk.StringVar(value=self.cfg.get("oracle_dsn", DEFAULT_DSN))
        ttk.Label(db, text="Usuario Oracle", style="Card.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=3)
        ttk.Entry(db, textvariable=self.user_var, width=22).grid(row=0, column=1, sticky="ew", padx=(0, 18), pady=3)
        ttk.Label(db, text="Servidor / servicio", style="Card.TLabel").grid(row=0, column=2, sticky="w", padx=(0, 8), pady=3)
        ttk.Entry(db, textvariable=self.dsn_var).grid(row=0, column=3, sticky="ew", pady=3)
        ttk.Label(db, text="Contraseña", style="Card.TLabel").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=3)
        self.password_var = tk.StringVar(value="")
        self.password_entry = ttk.Entry(db, textvariable=self.password_var, show="*")
        self.password_entry.grid(row=1, column=1, sticky="ew", padx=(0, 18), pady=3)
        saved_password = get_saved_password(self.user_var.get().strip())
        self.remember_password_var = tk.BooleanVar(value=bool(saved_password))
        if saved_password:
            self.password_var.set(saved_password)
        ttk.Checkbutton(db, text="Recordar contraseña en este equipo", variable=self.remember_password_var).grid(
            row=1, column=2, columnspan=2, sticky="w", pady=3)

        # Archivo maestro y periodos, en dos columnas
        master = ttk.LabelFrame(root, text="  Archivos Excel  ", style="Card.TLabelframe", padding=(9, 6))
        master.grid(row=2, column=0, sticky="nsew", padx=(0, 5), pady=(0, 6))
        master.columnconfigure(0, weight=1)
        self.master_var = tk.StringVar(value=self.cfg.get("master_path", ""))
        self.unavailable_var = tk.StringVar(value=self.cfg.get("unavailable_path", ""))
        ttk.Label(master, text="Excel maestro de aulas", style="Card.TLabel").grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 2))
        ttk.Entry(master, textvariable=self.master_var).grid(row=1, column=0, sticky="ew", padx=(0, 7), pady=(0, 5))
        ttk.Button(master, text="Buscar archivo…", command=self.pick_master).grid(row=1, column=1, sticky="e", pady=(0, 5))
        ttk.Label(master, text="Aulas no disponibles (opcional)", style="Card.TLabel").grid(
            row=2, column=0, columnspan=2, sticky="w", pady=(1, 2))
        ttk.Entry(master, textvariable=self.unavailable_var).grid(row=3, column=0, sticky="ew", padx=(0, 7))
        ttk.Button(master, text="Buscar archivo…", command=self.pick_unavailable).grid(row=3, column=1, sticky="e")
        ttk.Label(master, text="Las rutas de ambos archivos se recuerdan localmente.", style="Muted.TLabel").grid(
            row=4, column=0, columnspan=2, sticky="w", pady=(4, 0))

        periods = ttk.LabelFrame(root, text="  Periodos académicos  ", style="Card.TLabelframe", padding=(9, 6))
        periods.grid(row=2, column=1, sticky="nsew", padx=(5, 0), pady=(0, 6))
        periods.columnconfigure(0, weight=1)
        self.period_var = tk.StringVar(value=", ".join(self.cfg.get("periods", DEFAULT_PERIODS)))
        ttk.Label(periods, text="Códigos separados por coma", style="Muted.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Entry(periods, textvariable=self.period_var).grid(row=1, column=0, sticky="ew", pady=(4, 5))
        self.remember_var = tk.BooleanVar(value=self.cfg.get("remember_periods", True))
        ttk.Checkbutton(periods, text="Recordar periodos para la próxima ejecución", variable=self.remember_var).grid(row=2, column=0, sticky="w")

        # Fechas y filtros
        dates = ttk.LabelFrame(root, text="  Fechas y filtros de disponibilidad  ", style="Card.TLabelframe", padding=(9, 6))
        dates.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(0, 6))
        for col, weight in [(1, 1), (3, 1), (5, 1), (7, 1)]:
            dates.columnconfigure(col, weight=weight)
        today = date.today()
        self.start_var = tk.StringVar(value=display_config_date(self.cfg.get("start_date"), today))
        self.end_var = tk.StringVar(value=display_config_date(self.cfg.get("end_date"), today + timedelta(days=30)))
        ttk.Label(dates, text="Fecha desde", style="Card.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 7), pady=4)
        ttk.Entry(dates, textvariable=self.start_var, width=13).grid(row=0, column=1, sticky="ew", padx=(0, 14), pady=4)
        ttk.Label(dates, text="Fecha hasta", style="Card.TLabel").grid(row=0, column=2, sticky="w", padx=(0, 7), pady=4)
        ttk.Entry(dates, textvariable=self.end_var, width=13).grid(row=0, column=3, sticky="ew", padx=(0, 14), pady=4)
        ttk.Label(dates, text="Sedes / locales", style="Card.TLabel").grid(row=0, column=4, sticky="w", padx=(0, 7), pady=4)
        self.local_options = []
        self.local_vars = {}
        self.local_popup = None
        self.local_button = ttk.Button(dates, text="Todas las sedes  ▾", command=self._toggle_local_popup, width=24)
        self.local_button.grid(row=0, column=5, sticky="ew", padx=(0, 14), pady=4)
        ttk.Label(dates, text="Horario desde", style="Card.TLabel").grid(row=0, column=6, sticky="w", padx=(0, 7), pady=4)
        self.time_start_var = tk.StringVar(value="")
        self.time_end_var = tk.StringVar(value="")
        ttk.Entry(dates, textvariable=self.time_start_var, width=8).grid(row=0, column=7, sticky="ew", padx=(0, 5), pady=4)
        ttk.Label(dates, text="hasta", style="Card.TLabel").grid(row=0, column=8, sticky="w", padx=(0, 5), pady=4)
        ttk.Entry(dates, textvariable=self.time_end_var, width=8).grid(row=0, column=9, sticky="ew", pady=4)

        ttk.Label(dates, text="Días de la semana", style="Card.TLabel").grid(row=1, column=0, sticky="w", padx=(0, 7), pady=(8, 2))
        self.day_vars = {}
        days_frame = ttk.Frame(dates)
        days_frame.grid(row=1, column=1, columnspan=9, sticky="w", pady=(8, 2))
        for col, weekday_idx, day_name in DAY_COLS:
            var = tk.BooleanVar(value=False)
            self.day_vars[weekday_idx] = var
            ttk.Checkbutton(days_frame, text=day_name, variable=var).pack(side="left", padx=(0, 13))
        ttk.Label(dates, text="Fechas en formato DD/MM/AAAA. Sin días ni sedes seleccionados se consideran todos. Horas opcionales HH:MM; horario base 07:00–23:59.",
                  style="Muted.TLabel").grid(row=2, column=0, columnspan=10, sticky="w", pady=(5, 0))

        filt = ttk.LabelFrame(root, text="  Alcance del reporte  ", style="Card.TLabelframe", padding=(9, 5))
        filt.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(0, 5))
        filt.columnconfigure(3, weight=1)
        self.filter_mode = tk.StringVar(value="Todo")
        self._last_filter_mode = "Todo"
        self.filter_value = tk.StringVar(value="")
        ttk.Radiobutton(filt, text="Todas las aulas", variable=self.filter_mode, value="Todo", command=self._filter_state).grid(row=0, column=0, sticky="w", padx=(0, 16))
        ttk.Radiobutton(filt, text="Por aula", variable=self.filter_mode, value="Aula", command=self._filter_state).grid(row=0, column=1, sticky="w", padx=(0, 16))
        ttk.Radiobutton(filt, text="Por atributo de salón", variable=self.filter_mode, value="Atributo", command=self._filter_state).grid(row=0, column=2, sticky="w", padx=(0, 16))
        self.filter_entry = ttk.Entry(filt, textvariable=self.filter_value, width=28, state="disabled")
        self.filter_entry.grid(row=0, column=3, sticky="ew", padx=(0, 10))
        ttk.Label(filt, text="Varios códigos: separados por coma (ej. U-01, U-A3).", style="Muted.TLabel").grid(row=1, column=0, columnspan=4, sticky="w", pady=(5, 0))
        ttk.Label(filt, text="Capacidad", style="Card.TLabel").grid(row=2, column=0, sticky="w", padx=(0, 7), pady=(8, 2))
        self.capacity_options = []
        self.capacity_vars = {}
        self.capacity_popup = None
        self.capacity_button = ttk.Button(filt, text="Todas las capacidades  ▾", command=self._toggle_capacity_popup, width=30)
        self.capacity_button.grid(row=2, column=1, columnspan=2, sticky="w", pady=(8, 2))
        self.filter_value.trace_add("write", lambda *_: self._on_filter_value_change())

        bottom = ttk.Frame(root)
        bottom.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(0, 5))
        ttk.Button(bottom, text="Generar reporte Excel", style="Accent.TButton", command=self.generate).pack(side="right", pady=(0, 1))
        ttk.Button(bottom, text="Carpeta de configuración", command=self.show_config_folder).pack(side="right", padx=(0, 8))
        self.status = tk.StringVar(value="Listo para generar reporte.")
        self.progress_text = tk.StringVar(value="")
        self.progress_window = None
        self.progress = None
        statusbar = tk.Frame(root, bg="#E5EBF2", height=27)
        statusbar.grid(row=7, column=0, columnspan=2, sticky="ew", pady=(4, 0))
        statusbar.grid_propagate(False)
        tk.Label(statusbar, textvariable=self.status, bg="#E5EBF2", fg="#45566B", anchor="w",
                 font=("Segoe UI", 9)).pack(fill="both", expand=True, padx=10, pady=5)

    def _filter_state(self):
        current_mode = self.filter_mode.get()
        previous_mode = getattr(self, "_last_filter_mode", "Todo")

        # La capacidad automática de un aula no debe quedarse seleccionada
        # al cambiar el alcance a atributo o a todas las aulas.
        if previous_mode == "Aula" and current_mode != "Aula":
            for var in self.capacity_vars.values():
                var.set(False)
            # La sede asignada automáticamente por el aula no debe convertirse
            # en un filtro manual al cambiar a Atributo o Todo.
            for var in self.local_vars.values():
                var.set(False)

        self._last_filter_mode = current_mode
        self.filter_entry.configure(state="normal" if current_mode != "Todo" else "disabled")
        self.refresh_local_options()

    def pick_master(self):
        p = filedialog.askopenfilename(title="Seleccionar Excel maestro de aulas",
                                       filetypes=[("Excel", "*.xlsx *.xlsm *.xls")])
        if p:
            self.master_var.set(p)
            self.refresh_local_options()

    def pick_unavailable(self):
        p = filedialog.askopenfilename(title="Seleccionar Excel de aulas no disponibles",
                                       filetypes=[("Excel", "*.xlsx *.xlsm *.xls")])
        if p:
            self.unavailable_var.set(p)

    def _update_local_button(self):
        selected = [name for name in self.local_options if self.local_vars.get(name) and self.local_vars[name].get()]
        if not selected and self.filter_mode.get() == "Aula" and self.filter_value.get().strip():
            label = "Sin sede encontrada"
        elif not selected:
            label = "Todas las sedes"
        elif len(selected) <= 2:
            label = ", ".join(selected)
        else:
            label = f"{len(selected)} sedes seleccionadas"
        if hasattr(self, "local_button"):
            self.local_button.configure(text=f"{label}  ▾")

    def _create_scrollable_checklist(self, popup, items, variable_getter, command, width=260):
        """Lista desplazable con rueda activa sobre el contenido o la barra.

        La zona de acciones queda fija en la parte inferior de la ventana.
        """
        outer = ttk.Frame(popup, padding=(10, 0, 10, 5))
        outer.pack(fill="both", expand=True)
        canvas = tk.Canvas(outer, highlightthickness=0, width=width, borderwidth=0)
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        frame = ttk.Frame(canvas)
        window_id = canvas.create_window((0, 0), window=frame, anchor="nw")

        def update_scrollregion(_event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def resize_inner(event):
            canvas.itemconfigure(window_id, width=event.width)
            update_scrollregion()

        frame.bind("<Configure>", update_scrollregion)
        canvas.bind("<Configure>", resize_inner)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def wheel(event):
            # Windows/macOS usan delta; Linux usa Button-4 y Button-5.
            if getattr(event, "num", None) == 4:
                step = -1
            elif getattr(event, "num", None) == 5:
                step = 1
            else:
                delta = getattr(event, "delta", 0)
                if not delta:
                    return "break"
                step = -1 if delta > 0 else 1
            canvas.yview_scroll(step * 3, "units")
            return "break"

        # La rueda funciona al pasar por el canvas, la barra o cualquier checkbox.
        for widget in (canvas, frame, scrollbar):
            widget.bind("<MouseWheel>", wheel)
            widget.bind("<Button-4>", wheel)
            widget.bind("<Button-5>", wheel)

        for item in items:
            checkbox = ttk.Checkbutton(frame, text=item, variable=variable_getter(item),
                                       command=command)
            checkbox.pack(anchor="w", fill="x", pady=2)
            checkbox.bind("<MouseWheel>", wheel)
            checkbox.bind("<Button-4>", wheel)
            checkbox.bind("<Button-5>", wheel)
        return outer

    def _toggle_local_popup(self):
        if self.local_popup is not None and self.local_popup.winfo_exists():
            self.local_popup.destroy()
            self.local_popup = None
            return
        popup = tk.Toplevel(self)
        apply_upc_icon(popup)
        self.local_popup = popup
        popup.title("Seleccionar sedes")
        popup.transient(self)
        popup.resizable(False, False)
        popup.protocol("WM_DELETE_WINDOW", self._close_local_popup)
        x = self.local_button.winfo_rootx()
        y = self.local_button.winfo_rooty() + self.local_button.winfo_height()
        popup.geometry(f"300x300+{x}+{y}")

        header = ttk.Frame(popup, padding=(10, 10, 10, 5))
        header.pack(fill="x")
        ttk.Label(header, text="Selecciona una o varias sedes", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        ttk.Label(header, text="Si no marcas ninguna, se incluyen todas.", foreground="#555555").pack(anchor="w", pady=(2, 0))

        actions = ttk.Frame(popup, padding=(10, 2, 10, 5))
        actions.pack(fill="x")
        ttk.Button(actions, text="Seleccionar todas", command=self._select_all_locales).pack(side="left")
        ttk.Button(actions, text="Limpiar", command=self._clear_locales).pack(side="left", padx=(6, 0))

        # Empaquetar el pie antes de la lista evita que el área expandible
        # oculte el botón cuando la ventana tiene poco espacio disponible.
        footer = ttk.Frame(popup, padding=(10, 0, 10, 10))
        footer.pack(side="bottom", fill="x")
        ttk.Button(footer, text="Listo", command=self._close_local_popup).pack(side="right")
        self._create_scrollable_checklist(
            popup, self.local_options, lambda name: self.local_vars[name],
            self._on_local_change, width=260)

    def _on_local_change(self):
        self._update_local_button()
        self.refresh_capacity_options()

    def _select_all_locales(self):
        for var in self.local_vars.values():
            var.set(True)
        self._update_local_button()
        self.refresh_capacity_options()

    def _clear_locales(self):
        for var in self.local_vars.values():
            var.set(False)
        self._update_local_button()
        self.refresh_capacity_options()

    def _close_local_popup(self):
        if self.local_popup is not None and self.local_popup.winfo_exists():
            self.local_popup.destroy()
        self.local_popup = None
        self._update_local_button()

    def _filtered_local_master(self, master):
        """Aplica el alcance Aula/Atributo antes de ofrecer sedes disponibles."""
        mode = self.filter_mode.get()
        value = self.filter_value.get().strip()
        if mode == "Aula" and value:
            wanted = {norm(x) for x in value.split(",") if norm(x)}
            master = master[master["CODIGO AMBIENTE"].isin(wanted)]
        elif mode == "Atributo" and value:
            wanted = {norm(x) for x in value.split(",") if norm(x)}
            master = master[master["ATRIBUTO SALON"].isin(wanted)]
        return master

    def _populate_local_options(self, master):
        """Muestra únicamente sedes compatibles con el alcance seleccionado.

        En modo Aula la sede se determina automáticamente y el selector se
        deshabilita. En modo Atributo se ofrecen solo sedes con ese atributo.
        En modo Todo se ofrecen todas las sedes del maestro.
        """
        selected_before = {name for name, var in self.local_vars.items() if var.get()}
        mode = self.filter_mode.get()
        value = self.filter_value.get().strip()
        filtered = self._filtered_local_master(master)
        values = sorted({norm(v) for v in filtered["DESCRIPCION LOCAL"].tolist() if norm(v)})
        if self.local_popup is not None and self.local_popup.winfo_exists():
            self._close_local_popup()

        # Para un aula, la sede no es una decisión manual: se obtiene del maestro.
        if mode == "Aula" and value:
            selected_now = set(values)
        else:
            selected_now = selected_before.intersection(values)

        self.local_options = values
        self.local_vars = {name: tk.BooleanVar(value=name in selected_now) for name in values}
        room_mode = mode == "Aula" and bool(value)
        if hasattr(self, "local_button"):
            self.local_button.configure(state="disabled" if room_mode else "normal")
        self._update_local_button()
        if hasattr(self, "capacity_button"):
            self.refresh_capacity_options()

    def _on_filter_value_change(self):
        # Actualiza primero las sedes válidas para el aula/atributo ingresado;
        # luego se recalculan las capacidades usando esas sedes.
        self.refresh_local_options()

    def refresh_local_options(self):
        path_text = self.master_var.get().strip()
        if not path_text or not Path(path_text).is_file():
            self.local_options = []
            self.local_vars = {}
            if hasattr(self, "local_button"):
                self.local_button.configure(state="disabled" if self.filter_mode.get() == "Aula" else "normal")
            self._update_local_button()
            return
        try:
            self._populate_local_options(read_master(Path(path_text)))
        except Exception:
            self.local_options = []
            self.local_vars = {}
            self._update_local_button()

    def _selected_locales_now(self):
        return [name for name in self.local_options
                if self.local_vars.get(name) and self.local_vars[name].get()]

    def _capacity_base_master(self):
        path_text = self.master_var.get().strip()
        if not path_text or not Path(path_text).is_file():
            return pd.DataFrame()
        master = read_master(Path(path_text))
        master = self._filtered_local_master(master)
        locales = self._selected_locales_now()
        if locales:
            master = master[master["DESCRIPCION LOCAL"].isin({norm(x) for x in locales})]
        return master

    def refresh_capacity_options(self):
        try:
            master = self._capacity_base_master()
            values = sorted({
                int(v) if float(v).is_integer() else float(v)
                for v in master["CAPACIDAD"].dropna().tolist()
            })
        except Exception:
            values = []

        selected_before = {v for v, var in self.capacity_vars.items() if var.get()}
        # Al seleccionar aula(s), la capacidad se obtiene automáticamente del maestro.
        # En los demás modos se conservan solo las capacidades que siguen disponibles.
        room_mode = self.filter_mode.get() == "Aula" and bool(self.filter_value.get().strip())
        if room_mode:
            selected_now = set(values)
        else:
            selected_now = selected_before.intersection(values)

        if self.capacity_popup is not None and self.capacity_popup.winfo_exists():
            self.capacity_popup.destroy()
            self.capacity_popup = None
        self.capacity_options = values
        self.capacity_vars = {v: tk.BooleanVar(value=v in selected_now) for v in values}
        if hasattr(self, "capacity_button"):
            self.capacity_button.configure(state="disabled" if room_mode else "normal")
        self._update_capacity_button()

    def _update_capacity_button(self):
        selected = [v for v in self.capacity_options
                    if self.capacity_vars.get(v) and self.capacity_vars[v].get()]
        if not selected:
            label = "Todas las capacidades"
        elif len(selected) <= 3:
            label = ", ".join(f"{v:g}" for v in selected)
        else:
            label = f"{len(selected)} capacidades seleccionadas"
        if hasattr(self, "capacity_button"):
            self.capacity_button.configure(text=f"{label}  ▾")

    def _toggle_capacity_popup(self):
        if self.capacity_popup is not None and self.capacity_popup.winfo_exists():
            self.capacity_popup.destroy()
            self.capacity_popup = None
            return
        popup = tk.Toplevel(self)
        apply_upc_icon(popup)
        self.capacity_popup = popup
        popup.title("Seleccionar capacidades")
        popup.transient(self)
        popup.resizable(False, False)
        popup.protocol("WM_DELETE_WINDOW", self._close_capacity_popup)
        x = self.capacity_button.winfo_rootx()
        y = self.capacity_button.winfo_rooty() + self.capacity_button.winfo_height()
        popup.geometry(f"280x300+{x}+{y}")

        header = ttk.Frame(popup, padding=(10, 10, 10, 5))
        header.pack(fill="x")
        ttk.Label(header, text="Selecciona una o varias capacidades",
                  font=("Segoe UI", 10, "bold")).pack(anchor="w")
        ttk.Label(header, text="Si no marcas ninguna, se incluyen todas.",
                  foreground="#555555").pack(anchor="w", pady=(2, 0))

        actions = ttk.Frame(popup, padding=(10, 2, 10, 5))
        actions.pack(fill="x")
        ttk.Button(actions, text="Seleccionar todas",
                   command=self._select_all_capacities).pack(side="left")
        ttk.Button(actions, text="Limpiar",
                   command=self._clear_capacities).pack(side="left", padx=(6, 0))

        footer = ttk.Frame(popup, padding=(10, 0, 10, 10))
        footer.pack(side="bottom", fill="x")
        ttk.Button(footer, text="Listo", command=self._close_capacity_popup).pack(side="right")
        self._create_scrollable_checklist(
            popup, self.capacity_options, lambda value: self.capacity_vars[value],
            self._update_capacity_button, width=240)

    def _select_all_capacities(self):
        for var in self.capacity_vars.values():
            var.set(True)
        self._update_capacity_button()

    def _clear_capacities(self):
        for var in self.capacity_vars.values():
            var.set(False)
        self._update_capacity_button()

    def _close_capacity_popup(self):
        if self.capacity_popup is not None and self.capacity_popup.winfo_exists():
            self.capacity_popup.destroy()
        self.capacity_popup = None
        self._update_capacity_button()

    def show_config_folder(self):
        messagebox.showinfo("Configuración local", f"Los ajustes se guardan aquí:\n{APP_DIR}\n\nLa contraseña, si eliges recordarla, se guarda en el almacén de credenciales de Windows.", parent=self)

    def _set_progress(self, message, active=True):
        """Muestra el avance en una ventana independiente, siempre visible."""
        self.status.set(message)
        if active:
            if self.progress_window is None or not self.progress_window.winfo_exists():
                win = tk.Toplevel(self)
                apply_upc_icon(win)
                self.progress_window = win
                win.title("Generando reporte de disponibilidad")
                win.transient(self)
                win.resizable(False, False)
                win.protocol("WM_DELETE_WINDOW", lambda: None)
                frame = ttk.Frame(win, padding=18)
                frame.pack(fill="both", expand=True)
                ttk.Label(frame, text="Generación del reporte", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(0, 10))
                self.progress_label = ttk.Label(frame, text=message, font=("Segoe UI", 9, "bold"), wraplength=440)
                self.progress_label.pack(anchor="w", fill="x", pady=(0, 10))
                self.progress = ttk.Progressbar(frame, mode="indeterminate", length=440)
                self.progress.pack(fill="x")
                win.update_idletasks()
                x = self.winfo_rootx() + max(0, (self.winfo_width() - win.winfo_width()) // 2)
                y = self.winfo_rooty() + max(0, (self.winfo_height() - win.winfo_height()) // 2)
                win.geometry(f"+{x}+{y}")
                self.progress.start(12)
            else:
                self.progress_label.configure(text=message)
                if not self.progress.instate(["!disabled"]):
                    self.progress.state(["!disabled"])
            self.progress_window.update_idletasks()
            self.update_idletasks()
        else:
            if self.progress_window is not None and self.progress_window.winfo_exists():
                if self.progress is not None:
                    self.progress.stop()
                self.progress_window.destroy()
            self.progress_window = None
            self.progress = None
            self.status.set(message)
        self.update_idletasks()

    def generate(self):
        try:
            user = self.user_var.get().strip()
            if not user:
                raise ValueError("Ingresa tu usuario Oracle.")
            dsn = self.dsn_var.get().strip()
            if not dsn:
                raise ValueError("Ingresa el servidor/servicio Oracle.")
            master_path = Path(self.master_var.get().strip())
            if not master_path.is_file():
                raise ValueError("Selecciona un Excel maestro de aulas válido.")
            master = read_master(master_path)
            unavailable_path_text = self.unavailable_var.get().strip()
            unavailable_path = Path(unavailable_path_text) if unavailable_path_text else None
            if unavailable_path is not None and not unavailable_path.is_file():
                raise ValueError("El Excel de aulas no disponibles no existe. Selecciona un archivo válido o deja el campo vacío.")
            unavailable = read_unavailable(unavailable_path) if unavailable_path is not None else None
            self._populate_local_options(master)
            selected_locales = [name for name in self.local_options if self.local_vars.get(name) and self.local_vars[name].get()]
            if not selected_locales:
                selected_locales = None  # Sin selección equivale a todas las sedes.
            selected_capacities = [v for v in self.capacity_options
                                   if self.capacity_vars.get(v) and self.capacity_vars[v].get()]
            if not selected_capacities:
                selected_capacities = None
            time_start_text = self.time_start_var.get().strip()
            time_end_text = self.time_end_var.get().strip()
            if bool(time_start_text) != bool(time_end_text):
                raise ValueError("Para filtrar por horario, completa Hora desde y Hora hasta, o deja ambos campos vacíos.")
            time_start = time_end = None
            if time_start_text and time_end_text:
                time_start = parse_time(time_start_text)
                time_end = parse_time(time_end_text)
                if time_start is None or time_end is None:
                    raise ValueError("Ingresa las horas en formato HH:MM, por ejemplo 08:00 y 12:00.")
                if time_end < time_start:
                    raise ValueError("Hora hasta no puede ser anterior a Hora desde.")
            periods = list(dict.fromkeys(norm(x) for x in self.period_var.get().split(",") if norm(x)))
            if not periods:
                raise ValueError("Ingresa al menos un periodo.")
            if any(not re.fullmatch(r"\d{6}", p) for p in periods):
                raise ValueError("Cada periodo debe tener 6 dígitos, separados por coma.")
            try:
                start_date = parse_ui_date(self.start_var.get())
                end_date = parse_ui_date(self.end_var.get())
            except ValueError as date_error:
                raise ValueError("Ingresa las fechas en formato DD/MM/AAAA, por ejemplo 30/05/2026.") from date_error
            if end_date < start_date:
                raise ValueError("La fecha final no puede ser anterior a la inicial.")
            if (end_date - start_date).days > 730:
                raise ValueError("El rango máximo es de 731 días para evitar reportes demasiado grandes.")
            mode = self.filter_mode.get()
            value = self.filter_value.get().strip()
            selected_days = {idx for idx, var in self.day_vars.items() if var.get()}
            if not selected_days:
                selected_days = None  # Ninguna casilla marcada equivale a todos los días.
            if mode != "Todo" and not value:
                raise ValueError("Ingresa el código de aula o atributo para aplicar el filtro.")
            self.cfg.update({"oracle_user": user, "oracle_dsn": dsn, "master_path": str(master_path),
                             "unavailable_path": str(unavailable_path) if unavailable_path else "",
                             "start_date": start_date.isoformat(), "end_date": end_date.isoformat(),
                             "remember_periods": bool(self.remember_var.get())})
            if self.remember_var.get():
                self.cfg["periods"] = periods
            else:
                self.cfg.pop("periods", None)
            save_config(self.cfg)
            password = self.password_var.get()
            if not password:
                messagebox.showwarning("Falta la contraseña", "Ingresa tu contraseña Oracle. Si deseas, marca Recordar contraseña en este equipo.", parent=self)
                self.password_entry.focus_set()
                return
            output = filedialog.asksaveasfilename(title="Guardar reporte de disponibilidad",
                                                  defaultextension=".xlsx",
                                                  initialdir=str(self.cfg.get("output_dir", Path.home() / "Documents")),
                                                  initialfile=f"Reporte_Disponibilidad_Aulas_{start_date:%Y%m%d}_{end_date:%Y%m%d}.xlsx",
                                                  filetypes=[("Excel", "*.xlsx")])
            if not output:
                self.status.set("Generación cancelada; no se seleccionó dónde guardar el Excel.")
                return
            self._set_progress("Preparando consulta a Oracle…")
            self._set_progress("Consultando Oracle… esto puede tardar unos minutos.")
            # La consulta se ejecuta en segundo plano para que la barra siga animándose
            # y la ventana no parezca bloqueada mientras Oracle responde.
            query_result = {"sessions": None, "error": None}
            query_done = tk.BooleanVar(value=False)

            def _query_worker():
                try:
                    query_result["sessions"] = query_oracle(user, password, dsn, periods)
                except Exception as query_error:
                    query_result["error"] = query_error
                finally:
                    self.after(0, lambda: query_done.set(True))

            threading.Thread(target=_query_worker, daemon=True).start()
            self.wait_variable(query_done)
            if query_result["error"] is not None:
                raise query_result["error"]
            sessions = query_result["sessions"]
            # Persist only after Oracle accepts the credentials. Never write the password to config.json.
            if self.remember_password_var.get():
                try:
                    save_password(user, password)
                except Exception as cred_error:
                    messagebox.showwarning("No se pudo recordar la contraseña", f"Oracle aceptó la conexión, pero Windows no permitió guardar la contraseña. Puedes continuar sin recordarla.\n\nDetalle: {cred_error}", parent=self)
            else:
                delete_saved_password(user)
            if sessions.empty:
                messagebox.showwarning("Sin registros", "La consulta no devolvió registros para los periodos indicados. No se generó disponibilidad, porque esto no significa que las aulas estén libres.", parent=self)
                self.status.set("La consulta terminó sin registros.")
                return
            self._set_progress("Analizando sesiones y disponibilidad…")
            summary, detail, alerts, raw = run_analysis(
                sessions, master, start_date, end_date, mode, value,
                selected_days=selected_days, selected_locales=selected_locales,
                time_start=time_start, time_end=time_end,
                selected_capacities=selected_capacities, unavailable=unavailable
            )
            self._set_progress("Preparando las hojas del reporte…")
            self._set_progress("Generando y dando formato al Excel…")
            with pd.ExcelWriter(output, engine="openpyxl", date_format="DD/MM/YYYY", datetime_format="DD/MM/YYYY") as writer:
                summary.to_excel(writer, sheet_name="Resumen ejecutivo", index=False)
                detail.to_excel(writer, sheet_name="Disponibilidad", index=False)
                alerts.to_excel(writer, sheet_name="Alertas", index=False)
                raw.to_excel(writer, sheet_name="Sesiones Oracle", index=False)

                # Presentación institucional del reporte. No altera datos ni cálculos.
                from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
                from openpyxl.utils import get_column_letter
                from openpyxl.formatting.rule import ColorScaleRule

                header_fill = PatternFill("solid", fgColor="C8102E")  # Rojo institucional
                header_font = Font(name="Aptos", size=10, bold=True, color="FFFFFF")
                body_font = Font(name="Aptos", size=10, color="262626")
                alternate_fill = PatternFill("solid", fgColor="F3F5F7")
                white_fill = PatternFill("solid", fgColor="FFFFFF")
                bottom_line = Side(style="hair", color="D9DEE5")
                header_line = Side(style="thin", color="9E1026")
                tab_colors = {
                    "Resumen ejecutivo": "C8102E",
                    "Disponibilidad": "548235",
                    "Alertas": "ED7D31",
                    "Sesiones Oracle": "7F8C9A",
                }
                preferred_widths = {
                    "PERIODO_NRC": 13, "NRC": 11, "COD_CURSO": 16,
                    "TITULO_CURSO": 34, "NOMBRE_CURSO_LARGO": 42,
                    "DESCRIPCION LOCAL": 24, "DES_CAMPUS": 20,
                    "CAMPUS": 13, "PABELLON": 15, "COD_EDIFICIO": 15,
                    "COD_SALON": 14, "CODIGO AMBIENTE": 17,
                    "COD_ATRIBUTO_SALON": 19, "DES_ATRIBUTO_SALON": 25,
                    "DESCRIPCION ATRIBUTO SALON": 30, "CAPACIDAD": 12,
                    "FECHA": 13, "FECHA_INICIO_SESION": 16,
                    "FECHA_FIN_SESION": 16, "HORA_INICIO": 13,
                    "HORA_FIN": 13, "HORA_INICIO_SESION": 18,
                    "HORA_FIN_SESION": 18, "DIA": 14, "DIAS": 24,
                    "OBSERVACION": 46, "ALERTA": 46, "DETALLE": 46,
                }
                # Columnas descriptivas: alineadas a la izquierda y con ajuste de línea.
                descriptive_tokens = ("TITULO", "NOMBRE", "DESCRIPCION", "OBSERVACION", "ALERTA", "DETALLE", "MOTIVO")
                # Columnas de identificación, fechas, horas y cifras: centradas para lectura rápida.
                centered_tokens = ("NRC", "PERIODO", "COD_", "CODIGO", "CAPACIDAD", "FECHA", "HORA", "DIA", "DIAS", "CANTIDAD", "TOTAL", "PORCENTAJE", "%")

                for ws in writer.book.worksheets:
                    ws.sheet_view.showGridLines = False
                    ws.freeze_panes = "A2"
                    ws.auto_filter.ref = ws.dimensions
                    ws.sheet_properties.tabColor = tab_colors.get(ws.title, "C8102E")
                    ws.row_dimensions[1].height = 34
                    ws.sheet_view.zoomScale = 90
                    ws.page_setup.orientation = "landscape"
                    ws.page_setup.fitToWidth = 1
                    ws.page_setup.fitToHeight = 0
                    ws.sheet_properties.pageSetUpPr.fitToPage = True
                    ws.page_margins.left = 0.25
                    ws.page_margins.right = 0.25
                    ws.page_margins.top = 0.5
                    ws.page_margins.bottom = 0.5
                    ws.print_title_rows = "1:1"

                    # Encabezados institucionales, centrados y con texto multilínea.
                    for cell in ws[1]:
                        cell.fill = header_fill
                        cell.font = header_font
                        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                        cell.border = Border(bottom=header_line)

                    # Ancho automático real: mide encabezado y todos los datos de cada columna.
                    # Las columnas descriptivas pueden crecer más; si el texto es extenso,
                    # se ajusta en varias líneas y también aumenta la altura de la fila.
                    column_widths = {}
                    for col_idx in range(1, ws.max_column + 1):
                        letter = get_column_letter(col_idx)
                        header = str(ws.cell(1, col_idx).value or "").strip()
                        header_upper = header.upper()
                        is_description = any(token in header_upper for token in descriptive_tokens)
                        max_len = len(header)
                        for row_idx in range(2, ws.max_row + 1):
                            value = ws.cell(row_idx, col_idx).value
                            if value is not None:
                                # Fechas/horas se dimensionan con una representación compacta.
                                if isinstance(value, (datetime, date)):
                                    text_value = value.strftime("dd/mm/yyyy")
                                elif isinstance(value, time):
                                    text_value = value.strftime("HH:MM")
                                else:
                                    text_value = str(value)
                                max_len = max(max_len, max((len(part) for part in text_value.splitlines()), default=0))
                        suggested = preferred_widths.get(header, max_len + 3)
                        if is_description:
                            width = min(max(max_len + 3, suggested, 22), 58)
                        else:
                            width = min(max(max_len + 3, suggested, len(header) + 3, 12), 42)
                        column_widths[col_idx] = width
                        ws.column_dimensions[letter].width = width

                    # Filas limpias, alternadas y altura calculada para que el contenido
                    # no quede pegado ni cortado al abrir el archivo.
                    for row_idx in range(2, ws.max_row + 1):
                        fill = alternate_fill if row_idx % 2 == 0 else white_fill
                        max_lines = 1
                        for cell in ws[row_idx]:
                            cell.fill = fill
                            cell.font = body_font
                            cell.border = Border(bottom=bottom_line)
                            header = str(ws.cell(1, cell.column).value or "").strip().upper()
                            is_description = any(token in header for token in descriptive_tokens)
                            is_centered = any(token in header for token in centered_tokens)
                            width = max(1, int(column_widths.get(cell.column, 12)) - 2)
                            if cell.value is not None:
                                value_text = str(cell.value)
                                lines = 0
                                for part in value_text.splitlines() or [""]:
                                    lines += max(1, (len(part) + width - 1) // width)
                                max_lines = max(max_lines, lines)
                            if is_description:
                                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                            elif is_centered or isinstance(cell.value, (int, float)):
                                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                            else:
                                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                        ws.row_dimensions[row_idx].height = min(max(22, max_lines * 15), 90)

                    # Formatos legibles para fechas y horas, sin modificar sus valores.
                    for row in ws.iter_rows(min_row=2):
                        for cell in row:
                            h = str(ws.cell(1, cell.column).value or "").upper()
                            if "FECHA" in h and cell.value is not None:
                                cell.number_format = "dd/mm/yyyy"
                            elif "HORA" in h and cell.value is not None:
                                cell.number_format = "hh:mm"

                    # Resalta visualmente los estados/porcentajes si esas columnas existen.
                    headers = {str(c.value).strip().upper(): c.column for c in ws[1] if c.value is not None}
                    for key, col_idx in headers.items():
                        if "PORCENTAJE" in key or key in {"% USO", "% UTILIZACION", "% UTILIZACIÓN"}:
                            if ws.max_row >= 3:
                                ws.conditional_formatting.add(
                                    f"{get_column_letter(col_idx)}2:{get_column_letter(col_idx)}{ws.max_row}",
                                    ColorScaleRule(start_type="min", start_color="F4CCCC",
                                                   mid_type="percentile", mid_value=50, mid_color="FFF2CC",
                                                   end_type="max", end_color="D9EAD3")
                                )
            self._set_progress("Guardando archivo y finalizando…")
            self.cfg["output_dir"] = str(Path(output).parent)
            save_config(self.cfg)
            self._set_progress(f"Reporte generado: {output}", active=False)
            messagebox.showinfo("Reporte generado",
                                f"Archivo creado correctamente:\n{output}\n\nAulas en resumen: {len(summary):,}\nFilas de disponibilidad: {len(detail):,}\nAlertas: {len(alerts):,}\nSesiones Oracle consideradas: {len(raw):,}")
        except Exception as e:
            self._set_progress("No se pudo generar el reporte.", active=False)
            messagebox.showerror("No se pudo generar el reporte", f"{e}\n\nSi el error persiste, copia el mensaje y lo revisamos.", parent=self)


if __name__ == "__main__":
    App().mainloop()
