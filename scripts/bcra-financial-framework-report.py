from __future__ import annotations

import argparse
import subprocess
import sys
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter


@dataclass(frozen=True)
class VariableProfile:
    name: str
    block: str
    unit: str
    metric_kind: str
    preferred_frequency: str
    reading_hint: str


RATE_COLUMNS = {
    "Inflación mensual",
    "Inflación interanual",
    "REM",
    "TNA de depósitos a plazo fijo en pesos, 30-44 días",
    "TNA de depósitos a plazo fijo en pesos, 30-44 días, hasta $100.000",
    "TNA de depósitos a plazo fijo en pesos, 30-44 días, de más de $1.000.000",
    "TAMAR",
    "BADLAR",
}


PROFILES: dict[str, VariableProfile] = {
    "TC Minorista": VariableProfile("TC Minorista", "Tipo de cambio", "ARS/USD", "level", "diaria", "Referencia de precio minorista del dólar."),
    "TC Mayorista": VariableProfile("TC Mayorista", "Tipo de cambio", "ARS/USD", "level", "diaria", "Referencia mayorista y ancla para precios financieros."),
    "Base monetaria": VariableProfile("Base monetaria", "Agregados monetarios", "millones de pesos", "level", "semanal/mensual", "Pulso de expansión o absorción monetaria."),
    "Circulación monetaria": VariableProfile("Circulación monetaria", "Agregados monetarios", "millones de pesos", "level", "semanal/mensual", "Efectivo total en circulación."),
    "Billetes en público": VariableProfile("Billetes en público", "Agregados monetarios", "millones de pesos", "level", "semanal/mensual", "Demanda transaccional de efectivo del público."),
    "Efectivo en entidades financieras": VariableProfile("Efectivo en entidades financieras", "Agregados monetarios", "millones de pesos", "level", "semanal/mensual", "Liquidez física dentro del sistema financiero."),
    "Depósitos de bancos en cta. cte.": VariableProfile("Depósitos de bancos en cta. cte.", "Agregados monetarios", "millones de pesos", "level", "semanal/mensual", "Liquidez bancaria en cuenta corriente del BCRA."),
    "Depósitos en efectivo en entidades financieras": VariableProfile("Depósitos en efectivo en entidades financieras", "Depósitos", "millones de pesos", "level", "mensual", "Tamaño del fondeo bancario en efectivo."),
    "En cuentas corrientes": VariableProfile("En cuentas corrientes", "Depósitos", "millones de pesos", "level", "mensual", "Liquidez transaccional privada y pública."),
    "En Caja de ahorros": VariableProfile("En Caja de ahorros", "Depósitos", "millones de pesos", "level", "mensual", "Ahorro líquido de corto plazo."),
    "A plazo": VariableProfile("A plazo", "Depósitos", "millones de pesos", "level", "mensual", "Fondeo a plazo y preferencia por tasa."),
    "M2 privado": VariableProfile("M2 privado", "Agregados monetarios", "porcentaje", "rate", "diaria/semanal", "Lectura de liquidez privada amplia informada por BCRA."),
    "Préstamos de entidades financieras al sector privado": VariableProfile("Préstamos de entidades financieras al sector privado", "Crédito", "millones de pesos", "level", "mensual", "Pulso de crédito privado nominal."),
    "Inflación mensual": VariableProfile("Inflación mensual", "Inflación", "% mensual", "rate", "mensual", "Ritmo mensual observado de inflación."),
    "Inflación interanual": VariableProfile("Inflación interanual", "Inflación", "% interanual", "rate", "mensual", "Régimen inflacionario de doce meses."),
    "REM": VariableProfile("REM", "Expectativas", "%", "rate", "mensual", "Expectativa de inflación relevada por BCRA."),
    "TNA de depósitos a plazo fijo en pesos, 30-44 días": VariableProfile("TNA de depósitos a plazo fijo en pesos, 30-44 días", "Tasas", "% TNA", "rate", "diaria/semanal", "Tasa pasiva promedio de plazo fijo."),
    "TNA de depósitos a plazo fijo en pesos, 30-44 días, hasta $100.000": VariableProfile("TNA de depósitos a plazo fijo en pesos, 30-44 días, hasta $100.000", "Tasas", "% TNA", "rate", "diaria/semanal", "Tasa minorista aproximada de plazo fijo."),
    "TNA de depósitos a plazo fijo en pesos, 30-44 días, de más de $1.000.000": VariableProfile("TNA de depósitos a plazo fijo en pesos, 30-44 días, de más de $1.000.000", "Tasas", "% TNA", "rate", "diaria/semanal", "Tasa de mayor monto / tramo institucional."),
    "TAMAR": VariableProfile("TAMAR", "Tasas", "% TNA", "rate", "diaria/semanal", "Referencia mayorista de depósitos de mayor monto."),
    "BADLAR": VariableProfile("BADLAR", "Tasas", "% TNA", "rate", "diaria/semanal", "Referencia de depósitos mayoristas de bancos privados."),
    "CER": VariableProfile("CER", "Índices de inflación", "índice", "level", "diaria", "Coeficiente de estabilización de referencia para instrumentos CER."),
    "UVA": VariableProfile("UVA", "Índices de inflación", "índice", "level", "diaria", "Unidad ajustada por CER para crédito/ahorro."),
    "UVI": VariableProfile("UVI", "Índices de inflación", "índice", "level", "diaria", "Unidad indexada por costo de construcción/inflación según serie BCRA."),
}


def main() -> None:
    parser = argparse.ArgumentParser(description="Genera un marco financiero BCRA con evolución diaria/semanal/mensual/anual.")
    parser.add_argument("--today", default=date.today().isoformat(), help="Fecha de corrida YYYY-MM-DD.")
    parser.add_argument("--snapshot-dir", default=None, help="Directorio con bcra_public_series.xlsx/csv. Default: data/snapshots/YYYY-MM-DD.")
    parser.add_argument("--output-dir", default=None, help="Directorio de salida. Default: data/reports/YYYY-MM-DD.")
    parser.add_argument("--regenerate-snapshot", action="store_true", help="Fuerza regenerar bcra_public_series antes del reporte.")
    args = parser.parse_args()

    root = Path(__file__).resolve().parents[1]
    run_date = pd.to_datetime(args.today).date()
    snapshot_dir = Path(args.snapshot_dir) if args.snapshot_dir else root / "data" / "snapshots" / run_date.isoformat()
    output_dir = Path(args.output_dir) if args.output_dir else root / "data" / "reports" / run_date.isoformat()
    output_dir.mkdir(parents=True, exist_ok=True)

    snapshot_path = ensure_snapshot(root, snapshot_dir, run_date, args.regenerate_snapshot)
    series = load_series(snapshot_path)
    metadata = load_metadata(snapshot_path)

    daily = build_periodic_frame(series, "D")
    weekly = build_periodic_frame(series, "W-FRI")
    monthly = build_periodic_frame(series, "ME")
    annual = build_periodic_frame(series, "YE")
    summary = build_summary(series)
    changes = build_changes(series)
    framework = build_framework(summary, changes, run_date, snapshot_path)

    output_xlsx = output_dir / f"marco_financiero_bcra_{run_date:%Y%m%d}.xlsx"
    output_md = output_dir / f"marco_financiero_bcra_{run_date:%Y%m%d}.md"
    write_workbook(
        output_xlsx,
        {
            "Marco_actual": framework,
            "Resumen_variables": summary,
            "Evolucion_diaria": daily,
            "Evolucion_semanal": weekly,
            "Evolucion_mensual": monthly,
            "Evolucion_anual": annual,
            "Cambios_por_variable": changes,
            "Metadata": metadata,
        },
    )
    write_markdown(output_md, framework, summary, changes, run_date, snapshot_path)

    print(f"OK Excel: {output_xlsx}")
    print(f"OK Markdown: {output_md}")


def ensure_snapshot(root: Path, snapshot_dir: Path, run_date: date, regenerate: bool) -> Path:
    xlsx_path = snapshot_dir / "bcra_public_series.xlsx"
    csv_path = snapshot_dir / "bcra_public_series.csv"
    if regenerate or not xlsx_path.exists():
        subprocess.run(
            [
                sys.executable,
                str(root / "scripts" / "bcra-public-series-export.py"),
                "--today",
                run_date.isoformat(),
                "--output-dir",
                str(snapshot_dir),
            ],
            check=True,
        )
    if xlsx_path.exists():
        return xlsx_path
    if csv_path.exists():
        return csv_path
    raise FileNotFoundError(f"No se encontró snapshot BCRA en {snapshot_dir}")


def load_series(path: Path) -> pd.DataFrame:
    if path.suffix.lower() == ".xlsx":
        frame = pd.read_excel(path, sheet_name="BCRA_public_series")
    else:
        frame = pd.read_csv(path)
    frame["Fecha"] = pd.to_datetime(frame["Fecha"], dayfirst=True, errors="coerce")
    frame = frame.dropna(subset=["Fecha"]).sort_values("Fecha").reset_index(drop=True)
    for column in frame.columns:
        if column != "Fecha":
            frame[column] = pd.to_numeric(frame[column], errors="coerce")
    return frame


def load_metadata(path: Path) -> pd.DataFrame:
    if path.suffix.lower() != ".xlsx":
        return pd.DataFrame()
    try:
        return pd.read_excel(path, sheet_name="Metadata")
    except ValueError:
        return pd.DataFrame()


def build_periodic_frame(series: pd.DataFrame, rule: str) -> pd.DataFrame:
    values = series.set_index("Fecha").sort_index()
    if rule == "D":
        out = values.tail(260).copy()
    else:
        out = values.resample(rule).last().dropna(how="all")
        out = out.tail({"W-FRI": 104, "ME": 60, "YE": 20}.get(rule, 100))
    return out.reset_index()


def build_summary(series: pd.DataFrame) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    data = series.set_index("Fecha").sort_index()
    for column in [c for c in series.columns if c != "Fecha"]:
        values = data[column].dropna()
        if values.empty:
            continue
        profile = PROFILES.get(column, default_profile(column))
        latest_date = values.index[-1]
        latest_value = float(values.iloc[-1])
        row = {
            "bloque": profile.block,
            "variable": column,
            "unidad": profile.unit,
            "tipo_metrica": "tasa" if profile.metric_kind == "rate" else "nivel/indice",
            "frecuencia_sugerida": profile.preferred_frequency,
            "ultima_fecha": latest_date.date().isoformat(),
            "ultimo_valor": latest_value,
            "cambio_dato_previo": diff_from_previous(values),
            "var_7d": period_change(values, latest_date, 7, profile.metric_kind),
            "var_30d": period_change(values, latest_date, 30, profile.metric_kind),
            "var_90d": period_change(values, latest_date, 90, profile.metric_kind),
            "var_ytd": ytd_change(values, latest_date, profile.metric_kind),
            "var_1y": period_change(values, latest_date, 365, profile.metric_kind),
            "lectura": profile.reading_hint,
        }
        rows.append(row)
    return pd.DataFrame(rows).sort_values(["bloque", "variable"]).reset_index(drop=True)


def build_changes(series: pd.DataFrame) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    data = series.set_index("Fecha").sort_index()
    windows = [("7d", 7), ("30d", 30), ("90d", 90), ("1y", 365)]
    for column in [c for c in series.columns if c != "Fecha"]:
        values = data[column].dropna()
        if values.empty:
            continue
        profile = PROFILES.get(column, default_profile(column))
        latest_date = values.index[-1]
        latest_value = float(values.iloc[-1])
        for label, days in windows:
            previous_date = latest_date - pd.Timedelta(days=days)
            previous_value = value_on_or_before(values, previous_date)
            if previous_value is None:
                continue
            rows.append(
                {
                    "bloque": profile.block,
                    "variable": column,
                    "ventana": label,
                    "fecha_actual": latest_date.date().isoformat(),
                    "valor_actual": latest_value,
                    "fecha_base": values[values.index <= previous_date].index[-1].date().isoformat(),
                    "valor_base": float(previous_value),
                    "cambio_abs": latest_value - float(previous_value),
                    "cambio_pct_o_pp": period_change(values, latest_date, days, profile.metric_kind),
                    "unidad_cambio": "p.p." if profile.metric_kind == "rate" else "%",
                }
            )
    return pd.DataFrame(rows)


def build_framework(summary: pd.DataFrame, changes: pd.DataFrame, run_date: date, snapshot_path: Path) -> pd.DataFrame:
    rows = [
        {
            "seccion": "Corte",
            "metrica": "Fecha de generación",
            "valor": run_date.isoformat(),
            "lectura": f"Reporte construido con {snapshot_path.name}; marco de trabajo para Finanzas con series públicas BCRA disponibles.",
        },
        {
            "seccion": "Cobertura",
            "metrica": "Variables disponibles",
            "valor": len(summary),
            "lectura": "Incluye tipo de cambio, agregados monetarios, depósitos, crédito privado, inflación, expectativas, tasas e índices CER/UVA/UVI.",
        },
    ]
    rows.extend(block_readings(summary))
    rows.extend(relative_rate_readings(summary))
    rows.extend(momentum_readings(summary))
    return pd.DataFrame(rows)


def block_readings(summary: pd.DataFrame) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for block, group in summary.groupby("bloque"):
        latest_dates = sorted(group["ultima_fecha"].dropna().astype(str).unique())
        rows.append(
            {
                "seccion": "Bloques disponibles",
                "metrica": block,
                "valor": len(group),
                "lectura": f"{len(group)} variables; última fecha observada {latest_dates[-1] if latest_dates else 'n/d'}.",
            }
        )
    return rows


def relative_rate_readings(summary: pd.DataFrame) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    latest = summary.set_index("variable")
    for left, right in [
        ("TAMAR", "BADLAR"),
        ("TNA de depósitos a plazo fijo en pesos, 30-44 días", "BADLAR"),
        ("Inflación mensual", "REM"),
    ]:
        if left in latest.index and right in latest.index:
            spread = float(latest.loc[left, "ultimo_valor"]) - float(latest.loc[right, "ultimo_valor"])
            rows.append(
                {
                    "seccion": "Relaciones clave",
                    "metrica": f"{left} vs {right}",
                    "valor": round(spread, 4),
                    "lectura": f"Diferencial actual de {spread:.2f} puntos; útil para monitorear premio relativo y expectativas.",
                }
            )
    if "UVI" in latest.index and "UVA" in latest.index:
        ratio = float(latest.loc["UVI", "ultimo_valor"]) / float(latest.loc["UVA", "ultimo_valor"])
        rows.append(
            {
                "seccion": "Relaciones clave",
                "metrica": "UVI / UVA",
                "valor": round(ratio, 4),
                "lectura": "Relación entre unidades indexadas para seguir divergencias de indexación.",
            }
        )
    return rows


def momentum_readings(summary: pd.DataFrame) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    candidates = summary.dropna(subset=["var_30d"]).copy()
    if candidates.empty:
        return rows
    strongest = candidates.reindex(candidates["var_30d"].abs().sort_values(ascending=False).index).head(5)
    for _, row in strongest.iterrows():
        unit = "p.p." if row["tipo_metrica"] == "tasa" else "%"
        rows.append(
            {
                "seccion": "Movimientos 30d destacados",
                "metrica": row["variable"],
                "valor": round(float(row["var_30d"]), 4),
                "lectura": f"Movimiento de 30 días: {float(row['var_30d']):.2f} {unit}.",
            }
        )
    return rows


def diff_from_previous(values: pd.Series) -> float | None:
    if len(values) < 2:
        return None
    return float(values.iloc[-1] - values.iloc[-2])


def period_change(values: pd.Series, latest_date: pd.Timestamp, days: int, metric_kind: str) -> float | None:
    previous = value_on_or_before(values, latest_date - pd.Timedelta(days=days))
    if previous is None:
        return None
    current = float(values.iloc[-1])
    previous = float(previous)
    if metric_kind == "rate":
        return current - previous
    if previous == 0:
        return None
    return (current / previous - 1) * 100


def ytd_change(values: pd.Series, latest_date: pd.Timestamp, metric_kind: str) -> float | None:
    start = pd.Timestamp(year=latest_date.year, month=1, day=1)
    previous = value_on_or_before(values, start)
    if previous is None:
        return None
    current = float(values.iloc[-1])
    previous = float(previous)
    if metric_kind == "rate":
        return current - previous
    if previous == 0:
        return None
    return (current / previous - 1) * 100


def value_on_or_before(values: pd.Series, target: pd.Timestamp) -> float | None:
    prior = values[values.index <= target].dropna()
    if prior.empty:
        return None
    return float(prior.iloc[-1])


def default_profile(column: str) -> VariableProfile:
    kind = "rate" if column in RATE_COLUMNS else "level"
    return VariableProfile(column, "Otros", "", kind, "según disponibilidad", "Serie disponible para seguimiento financiero.")


def write_workbook(path: Path, sheets: dict[str, pd.DataFrame]) -> None:
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=name[:31], index=False)

    wb = load_workbook(path)
    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
        for column_cells in ws.columns:
            width = min(max(len(str(c.value)) if c.value is not None else 0 for c in column_cells) + 2, 60)
            ws.column_dimensions[get_column_letter(column_cells[0].column)].width = width
        ws.auto_filter.ref = ws.dimensions

    add_line_chart(wb, "Evolucion_mensual", "BCRA mensual - variables seleccionadas", ["TC Mayorista", "Base monetaria", "BADLAR", "TAMAR", "CER", "UVA", "UVI"], "J2")
    add_line_chart(wb, "Evolucion_anual", "BCRA anual - cierre de año", ["TC Mayorista", "Base monetaria", "BADLAR", "TAMAR", "CER", "UVA", "UVI"], "J2")
    wb.save(path)


def add_line_chart(wb, sheet_name: str, title: str, desired_columns: list[str], anchor: str) -> None:
    if sheet_name not in wb.sheetnames:
        return
    ws = wb[sheet_name]
    if ws.max_row < 3:
        return
    header = [cell.value for cell in ws[1]]
    selected = [header.index(col) + 1 for col in desired_columns if col in header]
    if not selected:
        return
    chart = LineChart()
    chart.title = title
    chart.y_axis.title = "Nivel / tasa"
    chart.x_axis.title = "Fecha"
    cats = Reference(ws, min_col=1, min_row=2, max_row=ws.max_row)
    for col_idx in selected[:7]:
        data = Reference(ws, min_col=col_idx, max_col=col_idx, min_row=1, max_row=ws.max_row)
        chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.height = 9
    chart.width = 20
    ws.add_chart(chart, anchor)


def write_markdown(path: Path, framework: pd.DataFrame, summary: pd.DataFrame, changes: pd.DataFrame, run_date: date, snapshot_path: Path) -> None:
    lines = [
        "# Marco financiero BCRA",
        "",
        f"Fecha de generación: {run_date.isoformat()}",
        f"Snapshot fuente: `{snapshot_path}`",
        "",
        "## Lectura ejecutiva",
        "",
    ]
    for _, row in framework.iterrows():
        lines.append(f"- **{row['seccion']} / {row['metrica']}**: {row['valor']}. {row['lectura']}")

    lines.extend(["", "## Último dato por variable", ""])
    display = summary.copy()
    numeric_cols = ["ultimo_valor", "cambio_dato_previo", "var_7d", "var_30d", "var_90d", "var_ytd", "var_1y"]
    for col in numeric_cols:
        if col in display:
            display[col] = display[col].map(format_number)
    lines.extend(markdown_table(display))

    lines.extend(["", "## Cambios por ventana", ""])
    if not changes.empty:
        changes_display = changes.copy()
        for col in ["valor_actual", "valor_base", "cambio_abs", "cambio_pct_o_pp"]:
            changes_display[col] = changes_display[col].map(format_number)
        lines.extend(markdown_table(changes_display.head(120)))
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def markdown_table(frame: pd.DataFrame) -> list[str]:
    columns = list(frame.columns)
    lines = [
        "| " + " | ".join(columns) + " |",
        "| " + " | ".join("---" for _ in columns) + " |",
    ]
    for _, row in frame.iterrows():
        values = ["" if pd.isna(row[col]) else str(row[col]).replace("|", "\\|") for col in columns]
        lines.append("| " + " | ".join(values) + " |")
    return lines


def format_number(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    return f"{float(value):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


if __name__ == "__main__":
    main()
