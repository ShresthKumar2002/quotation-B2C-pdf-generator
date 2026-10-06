from __future__ import annotations

from datetime import date
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Optional
import copy
import csv
import io
import json
import os
import re
import shutil
import subprocess
import tempfile
import urllib.parse
import urllib.request
import uuid
from concurrent.futures import ThreadPoolExecutor
import time

from fastapi import FastAPI, Form, Request
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from docx import Document

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT


def center_cell(cell):
    # Vertical center
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    # Horizontal center
    for paragraph in cell.paragraphs:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        
            

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_DIR = BASE_DIR / "templates"
STATIC_DIR = BASE_DIR / "static"
GENERATED_DIR = BASE_DIR / "generated"
GENERATED_DIR.mkdir(parents=True, exist_ok=True)

app = FastAPI(title="QuotationPro", version="2.0.0")
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")
templates = Jinja2Templates(directory=STATIC_DIR)

GST_RATE = Decimal("0.18")
FLAGS = set(range(14, 24))
SHEET_NAMES = ["Main", "Size", "Vehicle selection", "Vehicle data", "Add ons", "Packaging charge"]

# Add the Google Sheets URL later, for example:
# GOOGLE_SHEET_URL=https://docs.google.com/spreadsheets/d/<SHEET_ID>/edit
GOOGLE_SHEET_URL = os.getenv("GOOGLE_SHEET_URL", "").strip()
GOOGLE_SHEET_CACHE_SECONDS = int(os.getenv("GOOGLE_SHEET_CACHE_SECONDS", "300"))
_sheet_cache: dict[str, tuple[float, list[list[str]]]] = {}


def money(value: Decimal | str | float | int) -> Decimal:
    try:
        amount = Decimal(str(value or "0"))
    except (InvalidOperation, ValueError, TypeError):
        amount = Decimal("0")
    return amount.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def decimal_value(value) -> Decimal:
    try:
        text = str(value if value is not None else "0").strip().replace(",", "")
        return Decimal(text or "0")
    except (InvalidOperation, ValueError, TypeError):
        return Decimal("0")


def safe_filename(value: str) -> str:
    value = re.sub(r"[^A-Za-z0-9._-]+", "_", value.strip())
    return value[:80] or "quotation"


def last_day_of_month(value: date) -> str:
    if value.month == 12:
        next_month = date(value.year + 1, 1, 1)
    else:
        next_month = date(value.year, value.month + 1, 1)
    return str(next_month.fromordinal(next_month.toordinal() - 1))


def replace_text_in_paragraph(paragraph, replacements: dict[str, str]) -> None:
    original = "".join(run.text or "" for run in paragraph.runs)
    if not original:
        return
    updated = original
    for key, value in replacements.items():
        updated = updated.replace(key, str(value))
    if updated == original:
        return
    if paragraph.runs:
        paragraph.runs[0].text = updated
        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.add_run(updated)


def replace_in_cell(cell, replacements: dict[str, str]) -> None:
    for paragraph in cell.paragraphs:
        replace_text_in_paragraph(paragraph, replacements)
    for table in cell.tables:
        for row in table.rows:
            for nested_cell in row.cells:
                replace_in_cell(nested_cell, replacements)


def replace_in_document(document: Document, replacements: dict[str, str]) -> None:
    for paragraph in document.paragraphs:
        replace_text_in_paragraph(paragraph, replacements)
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_in_cell(cell, replacements)
    for section in document.sections:
        for header_footer in (section.header, section.footer):
            for paragraph in header_footer.paragraphs:
                replace_text_in_paragraph(paragraph, replacements)
            for table in header_footer.tables:
                for row in table.rows:
                    for cell in row.cells:
                        replace_in_cell(cell, replacements)


def google_sheet_id(sheet_url: str) -> str:
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", sheet_url)
    if not match:
        raise RuntimeError("GOOGLE_SHEET_URL is not a valid Google Sheets URL.")
    return match.group(1)


def fetch_google_sheet(sheet_name: str) -> list[list[str]]:
    if not GOOGLE_SHEET_URL:
        raise RuntimeError("Google Sheets URL is not configured. Set GOOGLE_SHEET_URL first.")

    cache_key = sheet_name
    cached = _sheet_cache.get(cache_key)
    now = time.time()
    if cached and now - cached[0] < GOOGLE_SHEET_CACHE_SECONDS:
        return cached[1]

    sheet_id = google_sheet_id(GOOGLE_SHEET_URL)
    params = urllib.parse.urlencode({"tqx": "out:csv", "sheet": sheet_name})
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?{params}"

    try:
        with urllib.request.urlopen(url, timeout=20) as response:
            raw = response.read().decode("utf-8-sig")
    except Exception as exc:
        raise RuntimeError(
            f"Could not fetch Google Sheet tab '{sheet_name}'. "
            "Make sure the spreadsheet is accessible to the application. "
            f"Details: {exc}"
        ) from exc

    rows = [list(row) for row in csv.reader(io.StringIO(raw))]
    _sheet_cache[cache_key] = (now, rows)
    return rows


def warm_sheet_cache(sheet_names: list[str]) -> None:
    """Fetch uncached Google Sheet tabs concurrently to reduce first-preview latency."""
    names = [name for name in sheet_names if name not in _sheet_cache]
    if not names:
        return
    with ThreadPoolExecutor(max_workers=min(6, len(names))) as executor:
        list(executor.map(fetch_google_sheet, names))


def ensure_quote_sheets_cached() -> None:
    # The quotation preview needs these four tabs. Fetching them concurrently
    # avoids waiting for several Google requests one after another.
    warm_sheet_cache(["Main", "Vehicle selection", "Vehicle data", "Add ons", "Packaging charge"])


def clean(value) -> str:
    return str(value if value is not None else "").strip()


def parse_flags(value) -> set[int]:
    if value is None or clean(value) == "":
        return set()
    found = set()
    for token in re.findall(r"\d+", clean(value)):
        number = int(token)
        if number in FLAGS:
            found.add(number)
    return found


def main_items() -> list[dict]:
    rows = fetch_google_sheet("Main")
    result = []
    # Main: B=description, C=Vol, D=Rate, E=flags/features; data starts at row 2.
    for row_number, row in enumerate(rows[1:], start=2):
        if len(row) < 2:
            continue
        description = clean(row[1])
        if not description:
            continue
        result.append({
            "row": row_number,
            "description": description,
            "vol": clean(row[2]) if len(row) > 2 else "0",
            "rate": clean(row[3]) if len(row) > 3 else "0",
            "flags": sorted(parse_flags(row[4] if len(row) > 4 else "")),
        })
    return result


def vehicle_selection() -> list[dict]:
    rows = fetch_google_sheet("Vehicle selection")
    result = []
    for row in rows[1:]:
        if len(row) < 4:
            continue
        vehicle_id = clean(row[0])
        if not vehicle_id:
            continue
        result.append({
            "vehicle_id": vehicle_id,
            "min_capacity": decimal_value(row[2]),
            "max_capacity": decimal_value(row[3]),
        })
    return result


def vehicle_data() -> list[dict]:
    rows = fetch_google_sheet("Vehicle data")
    result = []
    for row in rows[1:]:
        if len(row) < 6:
            continue
        vehicle_id = clean(row[0])
        city = clean(row[1])
        if not vehicle_id or not city:
            continue
        result.append({
            "vehicle_id": vehicle_id,
            "city": city,
            "fixed_cost": decimal_value(row[3]),
            "base_manpower": decimal_value(row[4]),
            "manpower_cost": decimal_value(row[5]),
        })
    return result


def addon_master() -> dict[int, dict]:
    rows = fetch_google_sheet("Add ons")
    result = {}
    for row in rows[1:]:
        if len(row) < 3:
            continue
        try:
            flag = int(decimal_value(row[0]))
        except Exception:
            continue
        if flag not in range(16, 24):
            continue
        result[flag] = {
            "flag": flag,
            "description": clean(row[1]),
            "cost": decimal_value(row[2]),
        }
    return result


def packaging_rates() -> tuple[Decimal, Decimal]:
    rows = fetch_google_sheet("Packaging charge")
    standard = decimal_value(rows[1][1]) if len(rows) > 1 and len(rows[1]) > 1 else Decimal("0")
    fragile = decimal_value(rows[2][1]) if len(rows) > 2 and len(rows[2]) > 1 else Decimal("0")
    return standard, fragile


def master_data_for_frontend() -> dict:
    warm_sheet_cache(["Main", "Vehicle data"])
    items = main_items()
    cities = sorted({row["city"] for row in vehicle_data() if row["city"]}, key=str.casefold)
    return {"items": [{"description": x["description"]} for x in items], "cities": cities}


def find_vehicle(total_volume: Decimal, city: str) -> dict:
    selected_vehicle = None
    for vehicle in vehicle_selection():
        if vehicle["min_capacity"] <= total_volume <= vehicle["max_capacity"]:
            selected_vehicle = vehicle
            break
    if not selected_vehicle:
        raise RuntimeError(f"No vehicle capacity range covers total item volume {total_volume}.")

    for data in vehicle_data():
        if data["vehicle_id"] == selected_vehicle["vehicle_id"] and data["city"].casefold() == city.casefold():
            result = dict(data)
            result["vehicle_id"] = selected_vehicle["vehicle_id"]
            return result
    raise RuntimeError(
        f"No Vehicle data row was found for vehicle '{selected_vehicle['vehicle_id']}' and city '{city}'."
    )


def calculate_quote(data: dict) -> dict:
    requested_items = data.get("items", [])
    ensure_quote_sheets_cached()
    if not requested_items:
        raise RuntimeError("Please select at least one item.")

    master_by_description = {item["description"]: item for item in main_items()}
    addon_lookup = addon_master()
    standard_packaging, fragile_packaging = packaging_rates()

    # Logistics is accepted by default unless explicitly set to No.
    raw_logistics = data.get("logistics_selected", "yes")
    if isinstance(raw_logistics, bool):
        logistics_selected = raw_logistics
    else:
        logistics_selected = str(raw_logistics).strip().lower() not in {"no", "false", "0", "off", "self drop", "self_drop"}

    storage_discount = max(
    Decimal("0"),
    min(Decimal("100"), decimal_value(data.get("storage_discount", "0"))))

    logistics_discount = max(
    Decimal("0"),
    min(Decimal("100"), decimal_value(data.get("logistics_discount", "0"))))

    # Add-on decisions are keyed by "description::flag". Missing decisions default to Yes.
    raw_addon_choices = data.get("addon_choices", {}) or {}
    if isinstance(raw_addon_choices, str):
        try:
            raw_addon_choices = json.loads(raw_addon_choices or "{}")
        except json.JSONDecodeError:
            raw_addon_choices = {}
    addon_choices = {
        str(k): str(v).strip().lower() != "no"
        for k, v in raw_addon_choices.items()
    }

    item_rows = []
    subtotal = Decimal("0")
    net_gst = Decimal("0")
    vol_of_items = Decimal("0")
    triggered_flags: set[int] = set()

    for submitted in requested_items:
        description = clean(submitted.get("description"))
        if description not in master_by_description:
            raise RuntimeError(f"Item '{description}' is not present in the Google Sheet Main tab.")

        master = master_by_description[description]
        qty = decimal_value(submitted.get("qty"))
        if qty <= 0:
            raise RuntimeError(f"Quantity for '{description}' must be greater than zero.")

        vol = decimal_value(master["vol"])
        rate = money(master["rate"])
        item_amount = money(rate * qty)
        item_volume = vol * qty
        item_gst = money(item_amount * GST_RATE)
        item_total = money(item_amount + item_gst)
        flags = set(master["flags"])
        triggered_flags.update(flags)
        vol_of_items += item_volume
        subtotal += item_amount
        net_gst += item_gst

        item_rows.append({
            "description": description,
            "qty": qty,
            "rate": rate,
            "vol": vol,
            "amount": item_amount,
            "volume": item_volume,
            "gst": item_gst,
            "total": item_total,
            "flags": flags,
        })

    subtotal = money(subtotal)
    if subtotal<500: subtotal = money(500)

# ==========================================
# STORAGE DISCOUNT
# ==========================================
    storage_discount_amount = money(
    subtotal * storage_discount / Decimal("100"))

    discounted_storage_subtotal = money(
    subtotal - storage_discount_amount)

# ==========================================
# GST AFTER STORAGE DISCOUNT
# ==========================================
    net_gst = money(
    discounted_storage_subtotal * GST_RATE)

# ==========================================
# FINAL MONTHLY STORAGE
# ==========================================
    net_monthly_storage = money(
    discounted_storage_subtotal + net_gst)

    vol_of_items = vol_of_items.quantize(
    Decimal("0.01"),
    rounding=ROUND_HALF_UP)

    vehicle = find_vehicle(vol_of_items, data["city"])
    base_manpower = vehicle["base_manpower"]
    if 15 in triggered_flags and base_manpower < 4:
        base_manpower = Decimal("4")

    packing_charge = Decimal("0")
    for item in item_rows:
        if 14 in item["flags"]:
            packing_charge += item["volume"] * fragile_packaging
        else:
            packing_charge += item["volume"] * standard_packaging
    packing_charge = money(packing_charge)

    # Base logistics is vehicle + manpower + packing. The quoted logistics
    # charge includes 18% GST, hence the required 1.18 multiplier.
    logistics_base_charge = money(
        vehicle["fixed_cost"] + base_manpower * vehicle["manpower_cost"] + packing_charge
    )
    logistics_charge = money(logistics_base_charge)
    logistics_discount_amount = money(
    logistics_charge * logistics_discount / Decimal("100"))

    discounted_logistics_charge = money(
    logistics_charge - logistics_discount_amount)

    applied_logistics_charge = (
    discounted_logistics_charge
    if logistics_selected
    else Decimal("0"))

    addon_rows = []
    addon_subtotal = Decimal("0")
    addon_flags = []
    for item in item_rows:
        for flag in sorted(item["flags"]):
            if flag not in range(16, 24):
                continue
            addon = addon_lookup.get(flag)
            if not addon:
                raise RuntimeError(f"Flag {flag} is triggered, but no matching row exists in Add ons.")

            addon_key = f"{item['description']}::{flag}"
            selected = addon_choices.get(addon_key, True)
            qty = item["qty"]
            cost = money(addon["cost"] * qty)
            row = {
                "id": addon_key,
                "description": f"{item['description']} - {addon['description']}",
                "qty": qty,
                "unit_cost": money(addon["cost"]),
                "amount": cost,
                "flag": flag,
                "selected": selected,
            }
            addon_rows.append(row)
            addon_flags.append(flag)
            if selected:
                addon_subtotal += cost

    addon_subtotal = money(addon_subtotal)
    addon_gst = money(addon_subtotal * GST_RATE)
    addon_total = money(addon_subtotal + addon_gst)

    security_deposit = discounted_storage_subtotal
    final_amount = money(net_monthly_storage + security_deposit + applied_logistics_charge + addon_total)
    token_amount = money(data.get("token_amount", "0"))

    return {
        "items": item_rows,
        "subtotal": subtotal,
        "storage_discount": storage_discount,
        "storage_discount_amount": storage_discount_amount,
        "discounted_storage_subtotal": discounted_storage_subtotal,

        "logistics_discount": logistics_discount,
        "logistics_discount_amount": logistics_discount_amount,
        "discounted_logistics_charge": discounted_logistics_charge,
        "net_gst": net_gst,
        "net_monthly_storage": net_monthly_storage,
        "security_deposit": security_deposit,
        "vol_of_items": vol_of_items,
        "triggered_flags": sorted(triggered_flags),
        "vehicle": vehicle,
        "base_manpower": base_manpower,
        "packing_charge": packing_charge,
        "logistics_base_charge": logistics_base_charge,
        "logistics_charge": logistics_charge,
        "applied_logistics_charge": applied_logistics_charge,
        "logistics_selected": logistics_selected,
        "addons": addon_rows,
        "addon_subtotal": addon_subtotal,
        "addon_gst": addon_gst,
        "addon_total": addon_total,
        "final_amount": final_amount,
        "token_amount": token_amount,
    }

def build_replacements(data: dict, calc: dict) -> dict[str, str]:
    quotation_date = date.fromisoformat(data["quotation_date"])
    storage_date = date.fromisoformat(data["storage_date"])
    items = calc["items"]
    first_item = items[0]

    item_placeholders = {}
    for index, row in enumerate(items, start=1):
        item_placeholders[f"{{{{ items.item{index}.description }}}}"] = row["description"]
        item_placeholders[f"{{{{ items.item{index}.qty }}}}"] = str(row["qty"])
        item_placeholders[f"{{{{ items.item{index}.amount }}}}"] = f"₹ {row['amount']:,.2f}"

    replacements = {
        "{{date}}": quotation_date.strftime("%d-%m-%Y"),
        "{{city}}": data["city"],
        "{{customer_name}}": data["customer_name"],
        "{{customer_address}}": data["customer_address"],
        "{{storage.date}}": storage_date.strftime("%d-%m-%Y"),
        "{{customer.cno}}": data["customer_cno"],
        "{{Produt.cat}}": data["product_category"],
        "{{tenure}}": data["tenure"],
        "{{Titem.amount}}": f"₹ {calc['subtotal']:,.2f}",
        "{{Titem.amount}}": f"₹ {calc['subtotal']:,.2f}",
        "{{s_discount}}":(f"{calc['storage_discount']:,.2f}%" if calc["storage_discount"] > 0 else "-"),
        "{{Titemd.amount}}": f"₹ {calc['discounted_storage_subtotal']:,.2f}",
        "{{ Titem.amout }}": f"₹ {calc['subtotal']:,.2f}",
        "{{Titem.amout}}": f"₹ {calc['subtotal']:,.2f}",
        "{{ Titem.gst }}": f"₹ {calc['net_gst']:,.2f}",
        "{{ Titem.total_amt }}": f"₹ {calc['net_monthly_storage']:,.2f}",
        "{{IFRSD}}": f"₹ {calc['security_deposit']:,.2f}",
        "{{IFRSD (B)}}": f"₹ {calc['security_deposit']:,.2f}",
        "{{ item.amout }}": f"₹ {calc['security_deposit']:,.2f}",
        "{{ item.amount }}": f"₹ {calc['security_deposit']:,.2f}",
        "{{packing_charge}}": (f"₹ {calc['logistics_charge']:,.2f}" if calc["logistics_selected"] else "Self drop"),
        "{{l_discount}}":(f"{calc['logistics_discount']:,.2f}%" if calc["logistics_selected"] else "-"),
        "{{packingd_charge}}": (f"₹ {calc['applied_logistics_charge']:,.2f}" if calc["logistics_selected"] else "Self drop"),
        "{{special.gst}}": f"₹ {calc['addon_gst']:,.2f}",
        "{{special.total}}": f"₹ {calc['addon_total']:,.2f}",
        "{{special.amount}}": f"₹ {calc['addon_subtotal']:,.2f}",
        "{{special.qty}}": "",
        "{{special.service}}": "",
        "{{net}}": f"₹ {calc['final_amount']:,.2f}",
        "{{tamount}}": f"₹ {calc['token_amount']:,.2f}",
        "{{last_day_of_ date _month}}": last_day_of_month(quotation_date),
        "{{last_day_of_date_month}}": last_day_of_month(quotation_date),
        # Legacy placeholders, if present.
        "{{item.description}}": first_item["description"],
        "{{ item.description }}": first_item["description"],
        "{{item.qty}}": str(first_item["qty"]),
        "{{ item.qty }}": str(first_item["qty"]),
        "{{item.amout}}": f"₹ {first_item['amount']:,.2f}",
        "{{ item.amout }}": f"₹ {first_item['amount']:,.2f}",
        "{{item.gst}}": f"₹ {first_item['gst']:,.2f}",
        "{{ item.gst }}": f"₹ {first_item['gst']:,.2f}",
        "{{item.total_amt}}": f"₹ {calc['net_monthly_storage']:,.2f}",
        "{{ item.total_amt }}": f"₹ {calc['net_monthly_storage']:,.2f}",
        "{{ item.total_a mt }}": f"₹ {calc['net_monthly_storage']:,.2f}",
    }
    replacements.update(item_placeholders)
    return replacements


def _set_unique_row_cells(row, values):
    # Directly set the requested cells. Repeated references caused by merged
    # cells are harmless here and avoiding id()-based de-duplication prevents
    # Python proxy id reuse from leaving placeholders behind.
    for index, value in values.items():
        if index >= len(row.cells):
            continue
        row.cells[index].text = "" if value is None else str(value)
        for cell in row.cells:
            center_cell(cell)


def populate_shared_storage_table(document: Document, items: list[dict]) -> None:
    """Expand the Shared Storage table to exactly match the selected items.

    The Word template contains one or more sample item rows.  Those sample
    rows are only prototypes; they must never limit the number of items in the
    generated quotation.
    """
    if not items:
        return

    target = None
    for table in document.tables:
        table_text = " ".join(
            cell.text.replace("\n", " ").strip()
            for row in table.rows for cell in row.cells
        ).lower()
        if "shared storage service" in table_text and "item & description" in table_text:
            target = table
            break
    if target is None or len(target.rows) < 3:
        return

    # Find the real column-header row.  The adjusted template can contain a
    # duplicated header row because of its Word table structure, so keep the
    # first header and treat everything below it as sample data.
    header_row = None
    for i, row in enumerate(target.rows):
        txt = " ".join(c.text.replace("\n", " ").strip().lower() for c in row.cells)
        if "item & description" in txt and "qty" in txt and "denom" in txt:
            header_row = i
            break
    if header_row is None:
        return

    following_rows = list(target.rows)[header_row + 1:]
    if not following_rows:
        return

    # The adjusted Word template may contain a duplicated header row. Find the
    # first actual sample-data row and use it as the formatting prototype.
    prototype_row = None
    for row in following_rows:
        txt = " ".join(c.text.replace("\n", " ").strip().lower() for c in row.cells)
        if not ("item & description" in txt and "qty" in txt and "denom" in txt):
            prototype_row = row
            break
    if prototype_row is None:
        return

    prototype_xml = copy.deepcopy(prototype_row._tr)

    # Remove every row after the header: duplicate headers and all old sample
    # rows are only prototypes.
    for row in reversed(following_rows):
        target._tbl.remove(row._tr)

    # Insert rows in order by advancing the XML insertion anchor each time.
    anchor = target.rows[header_row]._tr
    for index, item in enumerate(items, start=1):
        row_xml = copy.deepcopy(prototype_xml)
        anchor.addnext(row_xml)
        anchor = row_xml
        inserted_index = next(i for i, r in enumerate(target.rows) if r._tr is row_xml)
        row = target.rows[inserted_index]

        # The template's visible columns are S.No, Description, Denom, Qty.
        values = {
            0: str(index),
            1: item["description"],
            2: "Nos",
            3: str(item["qty"]),
        }
        _set_unique_row_cells(row, values)
        for cell_index in range(4, len(row.cells)):
            row.cells[cell_index].text = ""


def populate_addon_table(document: Document, addons: list[dict]) -> None:
    """Populate the dynamic Add on Services table from the selected-item flags.

    The current template contains:
      row 0: section title
      row 1: column headers
      row 2+: add-on data rows
      final 3 rows: Total, GST(18%), One Time Logistics Charges

    Row 2 is used as the prototype, and additional rows are cloned from it.
    The three summary rows are always retained and populated with the aggregate
    add-on subtotal, GST, and total.
    """
    for table in document.tables:
        text = " ".join(
            cell.text.replace("\n", " ").strip()
            for row in table.rows
            for cell in row.cells
        ).lower()
        if "add on services" not in text or "s.special1.service" not in text:
            continue

        # The adjusted template has 2 header rows + 3 summary rows and at
        # least one prototype data row.
        if len(table.rows) < 6 or len(table.columns) < 5:
            return

        prototype_xml = copy.deepcopy(table.rows[2]._tr)

        # Keep the first prototype only while removing all existing add-on
        # data rows. The final three rows are the summary rows.
        for row in list(table.rows[2:-3])[::-1]:
            table._tbl.remove(row._tr)

        # Re-acquire rows after XML mutation.
        total_row = table.rows[-3]
        gst_row = table.rows[-2]
        one_time_row = table.rows[-1]

        active_addons = [a for a in addons if a.get("selected", True)]

        if active_addons:
            for index, addon in enumerate(active_addons, start=1):
                row_xml = copy.deepcopy(prototype_xml)
                total_row._tr.addprevious(row_xml)
                inserted_index = next(
                    i for i, r in enumerate(table.rows) if r._tr is row_xml
                )
                row = table.rows[inserted_index]
                _set_unique_row_cells(row, {
                    0: index,
                    1: addon["description"],
                    2: "Nos",
                    3: str(addon["qty"]),
                    4: f"₹ {addon['amount']:,.2f}",
                })
        else:
            # Keep one clean data row when no add-on flag was triggered.
            row_xml = copy.deepcopy(prototype_xml)
            total_row._tr.addprevious(row_xml)
            inserted_index = next(
                i for i, r in enumerate(table.rows) if r._tr is row_xml
            )
            row = table.rows[inserted_index]
            _set_unique_row_cells(row, {
                0: "",
                1: "No add-on services",
                2: "",
                3: "",
                4: "₹ 0.00",
            })

        # The calculation object already contains the authoritative aggregate
        # values. The adjusted template names them s.tamount, s.gst and s.total.
        # Set the final column directly so malformed/legacy placeholder spacing
        # in the DOCX cannot prevent replacement.
        calc = getattr(document, "_quotation_calc", None)
        if calc is not None:
            _set_unique_row_cells(total_row, {4: f"₹ {calc['addon_subtotal']:,.2f}"})
            _set_unique_row_cells(gst_row, {4: f"₹ {calc['addon_gst']:,.2f}"})
            _set_unique_row_cells(one_time_row, {4: f"₹ {calc['addon_total']:,.2f}"})
        else:
            # Fallback for direct callers of this helper.
            subtotal = sum(Decimal(str(a.get("amount", 0))) for a in active_addons)
            subtotal = money(subtotal)
            gst = money(subtotal * GST_RATE)
            total = money(subtotal + gst)
            _set_unique_row_cells(total_row, {4: f"₹ {subtotal:,.2f}"})
            _set_unique_row_cells(gst_row, {4: f"₹ {gst:,.2f}"})
            _set_unique_row_cells(one_time_row, {4: f"₹ {total:,.2f}"})
        return


def generate_docx(data: dict, output_path: Path) -> dict:
    template_path = TEMPLATE_DIR / "B2CquotationTemplate.docx"
    if not template_path.exists():
        raise FileNotFoundError("B2CquotationTemplate.docx was not found in app/templates")
    calc = calculate_quote(data)
    document = Document(str(template_path))
    replace_in_document(document, build_replacements(data, calc))
    populate_shared_storage_table(document, calc["items"])
    document._quotation_calc = calc
    populate_addon_table(document, calc["addons"])
    try:
        del document._quotation_calc
    except AttributeError:
        pass
    document.save(str(output_path))
    return calc


def find_libreoffice():
    env_path = os.getenv("LIBREOFFICE_PATH")
    candidates = [
        env_path, shutil.which("libreoffice"), shutil.which("soffice"),
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/usr/bin/libreoffice", "/usr/bin/soffice",
        "/usr/lib/libreoffice/program/soffice",
        "/usr/local/bin/libreoffice", "/usr/local/bin/soffice", "/snap/bin/libreoffice",
    ]
    for candidate in candidates:
        if candidate and Path(candidate).is_file():
            return str(candidate)
    return None


def convert_to_pdf(docx_path, pdf_path):
    soffice = find_libreoffice()
    if not soffice:
        raise RuntimeError("LibreOffice was not found. Install LibreOffice or set LIBREOFFICE_PATH to soffice.exe.")
    output_dir = Path(pdf_path).parent
    result = subprocess.run([
        soffice, "--headless", "--convert-to", "pdf", "--outdir", str(output_dir), str(docx_path)
    ], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=120)
    generated = output_dir / (Path(docx_path).stem + ".pdf")
    requested = Path(pdf_path)
    if result.returncode != 0 or not generated.exists():
        raise RuntimeError(f"LibreOffice PDF conversion failed. stdout: {result.stdout.strip()} stderr: {result.stderr.strip()}")
    if generated.resolve() != requested.resolve():
        if requested.exists():
            requested.unlink()
        shutil.move(str(generated), str(requested))


@app.get("/", response_class=HTMLResponse)
def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request, "today": date.today().isoformat()})


@app.get("/api/master-data")
def api_master_data():
    try:
        return master_data_for_frontend()
    except Exception as exc:
        return {"items": [], "cities": [], "error": str(exc)}


@app.post("/api/preview")
def api_preview(
    items_json: str = Form("[]"),
    city: str = Form(""),
    logistics_selected: str = Form("yes"),
    addon_choices_json: str = Form("{}"),
    storage_discount: str = Form("0"),
    logistics_discount: str = Form("0"),
):
    try:
        items = json.loads(items_json or "[]")
        addon_choices = json.loads(addon_choices_json or "{}")
        calc = calculate_quote({
            "items": items,
            "city": city,
            "logistics_selected": logistics_selected,
            "addon_choices": addon_choices,
            "storage_discount": storage_discount,
            "logistics_discount": logistics_discount,
        })
        return {
            "subtotal": str(calc["subtotal"]),
            "storage_discount": str(calc["storage_discount"]),
            "storage_discount_amount": str(calc["storage_discount_amount"]),
            "discounted_storage_subtotal": str(calc["discounted_storage_subtotal"]),

            "logistics_discount": str(calc["logistics_discount"]),
            "logistics_discount_amount": str(calc["logistics_discount_amount"]),
            "discounted_logistics_charge": str(calc["discounted_logistics_charge"]),
            "net_gst": str(calc["net_gst"]),
            "net_monthly_storage": str(calc["net_monthly_storage"]),
            "security_deposit": str(calc["security_deposit"]),
            "vol_of_items": str(calc["vol_of_items"]),
            "logistics_charge": str(calc["applied_logistics_charge"]),
            "logistics_base_charge": str(calc["logistics_base_charge"]),
            "logistics_selected": calc["logistics_selected"],
            "addons": [
                {
                    "id": a["id"],
                    "description": a["description"],
                    "qty": str(a["qty"]),
                    "amount": str(a["amount"]),
                    "selected": a["selected"],
                }
                for a in calc["addons"]
            ],
            "addon_subtotal": str(calc["addon_subtotal"]),
            "addon_gst": str(calc["addon_gst"]),
            "addon_total": str(calc["addon_total"]),
            "final_amount": str(calc["final_amount"]),
        }
    except Exception as exc:
        return {"error": str(exc)}


@app.post("/generate")
def generate(
    quotation_type: str = Form("storage"), quotation_date: str = Form(...), city: str = Form(...),
    customer_name: str = Form(...), customer_address: str = Form(...), customer_cno: str = Form(...),
    storage_date: str = Form(...), product_category: str = Form(...), tenure: str = Form(...),
    token_amount: str = Form("0"), items_json: str = Form("[]"), logistics_selected: str = Form("yes"), addon_choices_json: str = Form("{}"),storage_discount: str = Form("0"),logistics_discount: str = Form("0"), output_format: str = Form("pdf"),
):
    try:
        submitted_items = json.loads(items_json or "[]")
        if not isinstance(submitted_items, list):
            submitted_items = []
    except json.JSONDecodeError:
        submitted_items = []

    try:
        submitted_addon_choices = json.loads(addon_choices_json or "{}")
        if not isinstance(submitted_addon_choices, dict):
            submitted_addon_choices = {}
    except json.JSONDecodeError:
        submitted_addon_choices = {}

    data = {
        "items": submitted_items, "quotation_type": quotation_type, "quotation_date": quotation_date,
        "city": city, "customer_name": customer_name, "customer_address": customer_address,
        "customer_cno": customer_cno, "storage_date": storage_date, "product_category": product_category,
        "tenure": tenure, "token_amount": token_amount,
        "logistics_selected": logistics_selected,
        "addon_choices": submitted_addon_choices,"storage_discount": storage_discount,
        "logistics_discount": logistics_discount,
    }

    identifier = safe_filename(quotation_date)
    customer = safe_filename(customer_name)
    city=safe_filename(city)
    # The Word template is used only as an internal intermediate file for
    # LibreOffice conversion. The user receives PDF only.
    docx_path = GENERATED_DIR / f".quotation_{customer}_{identifier}.docx"
    pdf_path = GENERATED_DIR / f"quotation_{customer}_{city}_{identifier}.pdf"

    try:
        generate_docx(data, docx_path)
        convert_to_pdf(docx_path, pdf_path)
        # Never expose or retain the generated DOCX.
        if docx_path.exists():
            docx_path.unlink()
        return FileResponse(pdf_path, media_type="application/pdf", filename=pdf_path.name)
    except RuntimeError as exc:
        if docx_path.exists():
            docx_path.unlink()
        return HTMLResponse(f"<h2>Quotation generation error</h2><p>{exc}</p>", status_code=400)
    except Exception as exc:
        if docx_path.exists():
            docx_path.unlink()
        return HTMLResponse(f"<h2>Quotation generation error</h2><p>{exc}</p>", status_code=400)
