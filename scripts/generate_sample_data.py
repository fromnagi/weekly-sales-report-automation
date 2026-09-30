"""Generate deterministic, fictional weekly sales files for the CLI demo."""

import csv
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook


ROOT = Path(__file__).resolve().parents[1]
DESTINATION = ROOT / "sample_data" / "input"
HEADERS = (
    "sale_date", "branch_code", "order_id", "line_no", "product_code",
    "category", "quantity", "unit_price", "status",
)
BRANCHES = ("SEOUL-01", "BUSAN-02", "INCHEON-03")
PRODUCTS = (("COFFEE-250", "Grocery", "7.50"),
            ("TOWEL-002", "Home", "12.00"),
            ("NOTEBOOK-A5", "Stationery", "4.25"))


def sale(index: int) -> list:
    product, category, price = PRODUCTS[index % len(PRODUCTS)]
    return [(date(2026, 9, 14) + timedelta(days=index % 7)).isoformat(),
            BRANCHES[index % len(BRANCHES)], f"ORD-{1000 + index}", 1,
            product, category, index % 4 + 1, price, "COMPLETED"]


def main() -> None:
    DESTINATION.mkdir(parents=True, exist_ok=True)
    first = [sale(index) for index in range(30)]
    first[3][0] = "2026-09-99"
    first[7][4] = ""
    first[11][6] = "two"
    first[15][0] = "unknown"
    first[15][6] = 0
    with (DESTINATION / "01_north.csv").open("w", encoding="utf-8", newline="") as stream:
        writer = csv.writer(stream)
        writer.writerow(HEADERS)
        writer.writerows(first)

    second = [sale(index) for index in range(30, 60)]
    second[0] = first[0].copy()  # A resubmitted sale in the second export.
    second[1][3] = 0
    second[2][8] = "CANCELLED"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Sales"
    sheet.append(HEADERS)
    for row in second:
        sheet.append(row)
    workbook.save(DESTINATION / "02_south.xlsx")
    workbook.close()

    with (DESTINATION / "03_legacy.csv").open("w", encoding="utf-8", newline="") as stream:
        writer = csv.writer(stream)
        writer.writerow([name for name in HEADERS if name != "quantity"])
        for index in range(60, 75):
            writer.writerow([value for name, value in zip(HEADERS, sale(index)) if name != "quantity"])
    print(f"Generated 3 files and 75 source rows in {DESTINATION}")


if __name__ == "__main__":
    main()
