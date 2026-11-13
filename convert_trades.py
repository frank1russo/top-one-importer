import csv
import re
from datetime import datetime, timezone
from decimal import Decimal, InvalidOperation
from pathlib import Path
from zoneinfo import ZoneInfo

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill


ROOT = Path(__file__).resolve().parent
IMPORTS_DIR = ROOT / "imports"
CONVERTED_DIR = ROOT / "converted"
LOCAL_TIMEZONE = ZoneInfo("America/New_York")

HEADERS = [
    "Ticket #",
    "Symbol",
    "Direction",
    "Contracts",
    "Open Time",
    "Open Price",
    "Close Time",
    "Close Price",
    "Duration",
    "PnL",
    "Commission",
    "Total PnL",
]

CURRENCY_FORMAT = '$#,##0.00;($#,##0.00);$0.00'
DATETIME_FORMAT = "yyyy/mm/dd hh:mm:ss"


def normalized_key(value: str) -> str:
    return re.sub(r"[^a-z0-9]", "", value.lower())


def find_value(row: dict[str, str], *names: str, default: str | None = None) -> str | None:
    normalized_row = {normalized_key(key): value for key, value in row.items() if key}
    for name in names:
        value = normalized_row.get(normalized_key(name))
        if value is not None and value.strip():
            return value.strip()
    return default


def parse_timestamp(value: str) -> datetime:
    text = value.strip()
    try:
        parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
    except ValueError:
        try:
            parsed = datetime.strptime(text, "%m/%d/%Y %H:%M:%S")
        except ValueError:
            parsed = datetime.strptime(text, "%d-%m-%Y, %H:%M:%S")

    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=timezone.utc)
    return parsed.astimezone(timezone.utc)


def excel_local_time(value: datetime) -> datetime:
    return value.astimezone(LOCAL_TIMEZONE).replace(tzinfo=None)


def parse_amount(value: str | None, *, default: Decimal | None = None) -> Decimal:
    if value is None or not value.strip():
        if default is not None:
            return default
        raise ValueError("A required monetary value is missing")

    text = value.strip().replace("$", "").replace(",", "")
    is_parenthesized = text.startswith("(") and text.endswith(")")
    if is_parenthesized:
        text = text[1:-1].strip()
    try:
        amount = Decimal(text)
    except InvalidOperation as error:
        raise ValueError(f"Invalid monetary value: {value!r}") from error
    return -amount if is_parenthesized else amount


def required(row: dict[str, str], *names: str) -> str:
    value = find_value(row, *names)
    if value is None:
        raise ValueError(f"Missing required field: {names[0]}")
    return value


def trade_values(row: dict[str, str], row_number: int) -> list[object]:
    buy_timestamp = find_value(row, "boughtTimestamp", "Buy Time", "Buy Timestamp")
    sell_timestamp = find_value(row, "soldTimestamp", "Sell Time", "Sell Timestamp")

    open_time_value = find_value(row, "Open Time")
    close_time_value = find_value(row, "Close Time")
    if open_time_value is not None and close_time_value is not None:
        open_utc = parse_timestamp(open_time_value)
        close_utc = parse_timestamp(close_time_value)
    elif buy_timestamp is not None and sell_timestamp is not None:
        buy_utc = parse_timestamp(buy_timestamp)
        sell_utc = parse_timestamp(sell_timestamp)
        open_utc, close_utc = sorted((buy_utc, sell_utc))
    elif open_time_value is not None and close_time_value is None:
        raise ValueError("Missing Close Time")
    else:
        raise ValueError("Missing Open/Close Time or boughtTimestamp/soldTimestamp fields")

    side = find_value(row, "Side", "Direction")
    if side is not None:
        normalized_side = side.strip().upper()
        if normalized_side == "BUY":
            direction = "Long"
        elif normalized_side == "SELL":
            direction = "Short"
        else:
            raise ValueError(f"Unsupported Side value on row {row_number}: {side!r}")
    elif buy_timestamp is not None and sell_timestamp is not None:
        direction = "Long" if parse_timestamp(buy_timestamp) < parse_timestamp(sell_timestamp) else "Short"
    else:
        direction = "Long" if open_utc <= close_utc else "Short"

    open_price = find_value(row, "Open Price")
    close_price = find_value(row, "Close Price")
    if open_price is None or close_price is None:
        buy_price = required(row, "buyPrice", "Buy Price")
        sell_price = required(row, "sellPrice", "Sell Price")
        if direction == "Long":
            open_price = open_price or buy_price
            close_price = close_price or sell_price
        else:
            open_price = open_price or sell_price
            close_price = close_price or buy_price

    commission = parse_amount(find_value(row, "Commission", "Commissions"), default=Decimal("0"))
    pnl = parse_amount(find_value(row, "PnL", "pnl"))
    ticket = find_value(row, "Ticket #", "Ticket", "buyFillId", "Buy Fill ID", default=str(row_number))

    return [
        ticket,
        required(row, "Symbol", "symbol"),
        direction,
        int(Decimal(required(row, "Contracts", "Lots", "qty", "Quantity"))),
        excel_local_time(open_utc),
        float(parse_amount(open_price)),
        excel_local_time(close_utc),
        float(parse_amount(close_price)),
        abs((close_utc - open_utc).total_seconds()),
        float(pnl),
        float(commission),
    ]


def autosize_columns(worksheet) -> None:
    for cells in worksheet.columns:
        column_letter = cells[0].column_letter
        max_length = max(len(str(cell.value or "")) for cell in cells)
        worksheet.column_dimensions[column_letter].width = min(max(max_length + 2, 12), 40)


def apply_trade_row_formatting(worksheet, excel_row: int) -> None:
    worksheet.cell(excel_row, 5).number_format = DATETIME_FORMAT
    worksheet.cell(excel_row, 7).number_format = DATETIME_FORMAT
    for column in (6, 8, 10, 11, 12):
        worksheet.cell(excel_row, column).number_format = CURRENCY_FORMAT
    if worksheet.cell(excel_row, 9).value < 10:
        worksheet.cell(excel_row, 9).fill = PatternFill(fill_type="solid", fgColor="FFFF00")


def trade_key(values: list[object]) -> tuple[object, ...]:
    ticket = str(values[0]).strip()
    return ("ticket", ticket) if ticket else ("trade", *values[1:])


def update_trade_worksheet(workbook: Workbook, csv_path: Path) -> tuple[object, int, int]:
    sheet_name = csv_path.stem.removeprefix("Trade-History-")[:31]
    legacy_sheet_name = csv_path.name[:31]
    if sheet_name in workbook.sheetnames:
        worksheet = workbook[sheet_name]
        existing_headers = [worksheet.cell(1, column).value for column in range(1, len(HEADERS) + 1)]
        if existing_headers != HEADERS:
            raise ValueError(f"Unexpected headers in worksheet {sheet_name!r}")
    elif legacy_sheet_name in workbook.sheetnames:
        worksheet = workbook[legacy_sheet_name]
        worksheet.title = sheet_name
        existing_headers = [worksheet.cell(1, column).value for column in range(1, len(HEADERS) + 1)]
        if existing_headers != HEADERS:
            raise ValueError(f"Unexpected headers in worksheet {sheet_name!r}")
    else:
        worksheet = workbook.create_sheet(sheet_name)
        worksheet.append(HEADERS)

    worksheet.freeze_panes = "A2"
    header_font = Font(bold=True)
    for cell in worksheet[1]:
        cell.font = header_font

    total_row = next(
        (row for row in range(2, worksheet.max_row + 1) if worksheet.cell(row, 1).value == "Total:"),
        worksheet.max_row + 1,
    )
    existing_keys = {
        trade_key([worksheet.cell(row, column).value for column in range(1, 12)])
        for row in range(2, total_row)
    }

    new_trades: list[list[object]] = []
    with csv_path.open("r", newline="", encoding="utf-8-sig") as source:
        reader = csv.DictReader(source)
        for row_number, row in enumerate(reader, start=1):
            values = trade_values(row, row_number)
            key = trade_key(values)
            if key in existing_keys:
                continue
            existing_keys.add(key)
            new_trades.append(values)

    if total_row <= worksheet.max_row:
        worksheet.delete_rows(total_row)

    for values in new_trades:
        worksheet.append(values + [None])
        excel_row = worksheet.max_row
        worksheet.cell(excel_row, 12, f"=J{excel_row}-K{excel_row}")
        apply_trade_row_formatting(worksheet, excel_row)

    total_row = worksheet.max_row + 1
    worksheet.cell(total_row, 1, "Total:")
    worksheet.cell(total_row, 12, f"=SUM(L2:L{total_row - 1})")
    worksheet.cell(total_row, 1).font = header_font
    worksheet.cell(total_row, 12).font = header_font
    worksheet.cell(total_row, 12).number_format = CURRENCY_FORMAT

    autosize_columns(worksheet)
    return worksheet, total_row, len(new_trades)


def main() -> None:
    csv_files = sorted(IMPORTS_DIR.glob("*.csv"))
    if not csv_files:
        raise SystemExit(f"No CSV files found in {IMPORTS_DIR}")

    output_path = CONVERTED_DIR / "main_trade_history.xlsx"
    if output_path.exists():
        workbook = load_workbook(output_path)
    else:
        workbook = Workbook()

    if "Summary" in workbook.sheetnames:
        summary = workbook["Summary"]
    else:
        summary = workbook.create_sheet("Summary", 0)
    if workbook.worksheets[0] is not summary:
        workbook._sheets.remove(summary)
        workbook._sheets.insert(0, summary)

    if summary.max_row:
        summary.delete_rows(1, summary.max_row)
    summary.append(["Worksheet", "Total PnL"])
    summary.freeze_panes = "A2"
    header_font = Font(bold=True)
    for cell in summary[1]:
        cell.font = header_font

    for csv_path in csv_files:
        worksheet, total_row, new_count = update_trade_worksheet(workbook, csv_path)
        summary_row = summary.max_row + 1
        escaped_name = worksheet.title.replace("'", "''")
        summary.append([worksheet.title, f"='{escaped_name}'!L{total_row}"])
        summary.cell(summary_row, 2).number_format = CURRENCY_FORMAT
        print(f"{worksheet.title}: added {new_count} new trades")

    summary_total_row = summary.max_row + 1
    summary.cell(summary_total_row, 1, "Total:")
    summary.cell(summary_total_row, 2, f"=SUM(B2:B{summary_total_row - 1})")
    summary.cell(summary_total_row, 1).font = header_font
    summary.cell(summary_total_row, 2).font = header_font
    summary.cell(summary_total_row, 2).number_format = CURRENCY_FORMAT
    autosize_columns(summary)

    CONVERTED_DIR.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
    print(f"Updated {output_path.relative_to(ROOT)}")


if __name__ == "__main__":
    main()