#!/usr/bin/env python3
"""Create a public, aggregate-only survey summary from an XLSX workbook."""

from __future__ import annotations

from collections import Counter, defaultdict
from datetime import date, datetime, timezone
import json
from pathlib import Path
import re
import sys
import unicodedata

from openpyxl import load_workbook
from openpyxl.utils.datetime import from_excel


MIN_GROUP_SIZE = 10
ALIASES = {
    "service": ["Servicio evaluado"],
    "dependency": ["Unidad Administrativa"],
    "start": ["Hora de inicio"],
    "solution": ["¿Se dio solución o respuesta a sus requerimientos?"],
    "timeliness": ["¿Los servicios fueron prestados de manera oportuna?"],
    "attitude": ["¿Cómo califica la actitud y disposición del personal que atendió sus requerimientos?"],
    "clarity": ["¿Las indicaciones del personal que atendió sus requerimientos fueron claras?"],
}
CRITERIA = ("solution", "timeliness", "attitude", "clarity")


def clean(value: object) -> str:
    return " ".join(str(value).split()) if value is not None else ""


def normalize(value: object) -> str:
    text = unicodedata.normalize("NFD", clean(value).lower())
    return "".join(char for char in text if unicodedata.category(char) != "Mn")


def find_column(headers: list[str], aliases: list[str]) -> int | None:
    normalized_aliases = [normalize(alias) for alias in aliases]
    for index, header in enumerate(headers):
        if normalize(header) in normalized_aliases:
            return index
    for index, header in enumerate(headers):
        if any(alias in normalize(header) for alias in normalized_aliases):
            return index
    return None


def parse_date(value: object) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        try:
            converted = from_excel(value)
            return converted.date() if isinstance(converted, datetime) else converted
        except (OverflowError, TypeError, ValueError):
            return None
    text = clean(value)
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(text[:10], fmt).date()
        except ValueError:
            pass
    return None


def parse_score(value: object) -> float | None:
    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, (int, float)):
        score = float(value)
    else:
        text = re.sub(r"[^\d,.-]", "", clean(value))
        if "," in text and "." not in text:
            text = text.replace(",", ".")
        elif "," in text and "." in text:
            text = text.replace(",", "")
        try:
            score = float(text)
        except ValueError:
            return None
    return score if 1 <= score <= 5 else None


def read_responses(workbook_path: Path) -> list[dict]:
    workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    sheet = workbook[workbook.sheetnames[0]]
    rows = sheet.iter_rows(values_only=True)
    raw_headers = next(rows, ())
    headers = [clean(value) for value in raw_headers]
    columns = {
        name: find_column(headers, aliases)
        for name, aliases in ALIASES.items()
    }
    for required in ("service", "dependency", "start"):
        if columns[required] is None:
            raise ValueError(f"No se encontró la columna requerida: {ALIASES[required][0]}")

    responses = []
    for cells in rows:
        if not any(clean(value) for value in cells):
            continue
        response_date = parse_date(cells[columns["start"]])
        service = clean(cells[columns["service"]])
        dependency = clean(cells[columns["dependency"]])
        if not response_date or not service or not dependency:
            continue

        ratings = {
            criterion: parse_score(cells[index]) if index is not None else None
            for criterion in CRITERIA
            for index in (columns[criterion],)
        }
        available_ratings = [score for score in ratings.values() if score is not None]
        responses.append({
            "date": response_date,
            "service": service,
            "dependency": dependency,
            "average": sum(available_ratings) / len(available_ratings) if available_ratings else 0,
            "ratings": ratings,
        })
    workbook.close()
    return responses


def period_label(response: dict, granularity: str) -> str:
    response_date = response["date"]
    if granularity == "month":
        return response_date.strftime("%Y-%m")
    if granularity == "quarter":
        quarter = (response_date.month - 1) // 3 + 1
        return f"{response_date.year}-T{quarter}"
    return str(response_date.year)


def group_periods(responses: list[dict], granularity: str) -> dict[str, list[dict]]:
    grouped = defaultdict(list)
    for response in responses:
        grouped[period_label(response, granularity)].append(response)
    return dict(sorted(grouped.items()))


def safe_distribution(responses: list[dict], key: str) -> list[dict] | None:
    counts = Counter(response[key] for response in responses)
    if not counts or any(count < MIN_GROUP_SIZE for count in counts.values()):
        return None
    return [
        {"name": name, "count": count}
        for name, count in sorted(counts.items(), key=lambda item: (-item[1], item[0].casefold()))
    ]


def summarize_period(label: str, responses: list[dict]) -> dict:
    rating_means = {}
    for criterion in CRITERIA:
        scores = [
            response["ratings"][criterion]
            for response in responses
            if response["ratings"][criterion] is not None
        ]
        rating_means[criterion] = round(sum(scores) / len(scores), 2) if len(scores) >= MIN_GROUP_SIZE else None

    low_count = sum(response["average"] < 5 for response in responses)
    return {
        "label": label,
        "count": len(responses),
        "average": round(sum(response["average"] for response in responses) / len(responses), 2),
        "below_five": low_count if low_count == 0 or low_count >= MIN_GROUP_SIZE else None,
        "criteria": rating_means,
        "services": safe_distribution(responses, "service"),
        "dependencies": safe_distribution(responses, "dependency"),
    }


def build_summary(workbook_path: Path) -> dict:
    responses = read_responses(workbook_path)
    base = {
        "schema_version": 1,
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "privacy_threshold": MIN_GROUP_SIZE,
    }
    if len(responses) < MIN_GROUP_SIZE:
        return {
            **base,
            "available": False,
            "granularity": "none",
            "message": "No hay suficientes respuestas para publicar resultados agregados.",
            "periods": [],
        }

    for granularity in ("month", "quarter", "year"):
        grouped = group_periods(responses, granularity)
        if all(len(group) >= MIN_GROUP_SIZE for group in grouped.values()):
            break
    else:
        granularity = "general"
        grouped = {"Todos los periodos": responses}

    if granularity == "general":
        summarized = [summarize_period(label, group) for label, group in grouped.items()]
        summarized[0]["label"] = "Todos los periodos"
    else:
        summarized = [summarize_period(label, group) for label, group in grouped.items()]

    return {
        **base,
        "available": True,
        "granularity": granularity,
        "message": "Solo se publican resultados agregados; no se incluyen nombres, identificadores ni comentarios.",
        "periods": summarized,
    }


def main() -> None:
    if len(sys.argv) != 3:
        raise SystemExit("Uso: build_public_summary.py ENTRADA.xlsx SALIDA.json")
    source_path = Path(sys.argv[1])
    output_path = Path(sys.argv[2])
    summary = build_summary(source_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Resumen público agregado generado: disponible={summary['available']}, periodos={len(summary['periods'])}.")


if __name__ == "__main__":
    main()
