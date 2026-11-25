import json
from datetime import date
from pathlib import Path
import sys
import tempfile
import unittest

from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from scripts.build_public_summary import build_summary


HEADERS = [
    "ID",
    "Nombre",
    "Hora de inicio",
    "Servicio evaluado",
    "Unidad Administrativa",
    "Sugerencias y/o recomendaciones",
    "¿Se dio solución o respuesta a sus requerimientos?",
    "¿Los servicios fueron prestados de manera oportuna?",
    "¿Cómo califica la actitud y disposición del personal que atendió sus requerimientos?",
    "¿Las indicaciones del personal que atendió sus requerimientos fueron claras?",
]


def write_workbook(path: Path, dates: list[date]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(HEADERS)
    for index, response_date in enumerate(dates):
        sheet.append([
            f"private-id-{index}",
            f"Private name {index}",
            response_date,
            "Soporte TIC",
            "Mesa de Servicios",
            f"Private comment {index}",
            4,
            5,
            4,
            5,
        ])
    workbook.save(path)


class PublicSummaryTests(unittest.TestCase):
    def test_summary_excludes_personal_and_free_text_fields(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            workbook_path = Path(temp_dir) / "input.xlsx"
            write_workbook(workbook_path, [date(2026, 1, 15)] * 12)

            summary = build_summary(workbook_path)
            serialized = json.dumps(summary, ensure_ascii=False)

            self.assertTrue(summary["available"])
            self.assertEqual(summary["granularity"], "month")
            self.assertEqual(summary["periods"][0]["count"], 12)
            self.assertEqual(summary["periods"][0]["services"][0]["count"], 12)
            self.assertNotIn("Private name", serialized)
            self.assertNotIn("private-id", serialized)
            self.assertNotIn("Private comment", serialized)

    def test_rolls_up_small_months_to_quarter(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            workbook_path = Path(temp_dir) / "input.xlsx"
            dates = [date(2026, 1, 15)] * 5 + [date(2026, 2, 15)] * 6
            write_workbook(workbook_path, dates)

            summary = build_summary(workbook_path)

            self.assertEqual(summary["granularity"], "quarter")
            self.assertEqual(summary["periods"][0]["label"], "2026-T1")
            self.assertEqual(summary["periods"][0]["count"], 11)

    def test_suppresses_results_below_minimum_group_size(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            workbook_path = Path(temp_dir) / "input.xlsx"
            write_workbook(workbook_path, [date(2026, 1, 15)] * 9)

            summary = build_summary(workbook_path)

            self.assertFalse(summary["available"])
            self.assertEqual(summary["periods"], [])


if __name__ == "__main__":
    unittest.main()
