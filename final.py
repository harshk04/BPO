import os
import re
from decimal import Decimal, InvalidOperation
from pathlib import Path
from types import SimpleNamespace
from typing import List, Tuple

from dotenv import load_dotenv
from google import genai
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

from excel_loan_calculator import LoanOutputs, calculate
from maingem import compute_insurance_rate, get_image_paths, process_single_image
from roboflow_inference import (
    DEFAULT_CONFIDENCE,
    DEFAULT_OVERLAP,
    ensure_dependencies as ensure_roboflow_dependencies,
    load_settings as load_roboflow_settings,
    run_inference,
)
from run_logger import RunLogger

INPUT_IMAGES_FOLDER = "images"
ROBOFLOW_OUTPUT_ROOT = "outputs"
ROBOFLOW_CONFIDENCE = DEFAULT_CONFIDENCE
ROBOFLOW_OVERLAP = DEFAULT_OVERLAP
OUTPUT_EXCEL_PATH = "final_loan_outputs.xlsx"
LOGS_DIR = "logs"
SHEET_NAME = "Final_Output"

EXCEL_HEADERS = [
    "sample no . , record no,",
    "Customer Reference Number",
    "Customer name",
    "City,State",
    "Purchase Value AND Down payment",
    "Loan Period AND Annual Interest",
    "Guarantor name",
    "Guarantor reference number",
    "Loan Amount AND Principal",
    "Total Interest for Loan Period AND Property Insurance Per Month",
]


def extract_decimal(value: str, field_name: str) -> Decimal:
    text = str(value or "").replace(",", "")
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    if not match:
        raise ValueError(f"Could not parse numeric value for {field_name}: {value!r}")
    try:
        return Decimal(match.group(0))
    except InvalidOperation as exc:
        raise ValueError(f"Invalid numeric value for {field_name}: {value!r}") from exc


def decimal_to_plain(value: Decimal) -> str:
    text = format(value.normalize(), "f")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text


def format_currency(value: Decimal, spaces_after_dollar: int = 2) -> str:
    grouped = f"{value:,.2f}".replace(",", " , ")
    return f"$" + (" " * spaces_after_dollar) + grouped


def format_currency_trimmed(value: Decimal, spaces_after_dollar: int = 2) -> str:
    text = f"{value:,.2f}"
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    grouped = text.replace(",", " , ")
    return f"$" + (" " * spaces_after_dollar) + grouped


def format_percent(raw_value: str) -> str:
    value = extract_decimal(raw_value, "percent_value")
    return f"{decimal_to_plain(value)} %"


def parse_percent_for_calculator(raw_value: str, field_name: str) -> Decimal:
    value = extract_decimal(raw_value, field_name)
    if "%" in str(raw_value):
        return value / Decimal("100")
    return value


def format_upper_words(raw_value: str, spaces_between_words: int) -> str:
    tokens = [token for token in str(raw_value or "").split() if token]
    separator = " " * spaces_between_words
    return separator.join(token.upper() for token in tokens)


def format_city_state(city: str, state: str) -> str:
    city_clean = " ".join(str(city or "").split()).upper()
    state_clean = " ".join(str(state or "").split()).upper()
    return f"{city_clean} , {state_clean}"


def format_loan_period_and_interest(record: dict) -> str:
    years = extract_decimal(record["loan_period_in_year"], "loan_period_in_year")
    years_text = f"{int(years)} YEARS"
    annual_interest_text = format_percent(record["annual_interest"])
    return f"{years_text} AND {annual_interest_text}"


def autosize_columns(ws) -> None:
    for column_cells in ws.columns:
        max_length = 0
        column_letter = column_cells[0].column_letter
        for cell in column_cells:
            cell_value = "" if cell.value is None else str(cell.value)
            max_length = max(max_length, len(cell_value))
        ws.column_dimensions[column_letter].width = min(max(max_length + 2, 18), 80)


def ensure_workbook(path: str):
    excel_path = Path(path)
    if excel_path.exists():
        wb = load_workbook(excel_path)
        ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.create_sheet(SHEET_NAME)
        existing_headers = [ws.cell(row=1, column=i + 1).value for i in range(len(EXCEL_HEADERS))]
        if existing_headers != EXCEL_HEADERS:
            ws.delete_rows(1, ws.max_row)
            ws.append(EXCEL_HEADERS)
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = SHEET_NAME
        ws.append(EXCEL_HEADERS)

    header_fill = PatternFill(fill_type="solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    return wb, ws


def append_row_to_excel(path: str, row: List[str]) -> None:
    wb, ws = ensure_workbook(path)
    ws.append(row)

    for cell in ws[ws.max_row]:
        cell.alignment = Alignment(vertical="top", wrap_text=True)

    autosize_columns(ws)
    wb.save(path)


def calculate_outputs(record: dict) -> Tuple[LoanOutputs, bool]:
    insurance_rate = compute_insurance_rate(record["discount_percent"], record["loan_period_in_year"])
    insurance_na = insurance_rate == "NA"

    insurance_rate_percent = (
        Decimal("0")
        if insurance_na
        else parse_percent_for_calculator(insurance_rate, "insurance_rate")
    )

    # To match your requested output format, map as follows:
    # reduction_percent <- reduction_percent_value_down_payment
    # down_payment_percent <- discount_percent
    result = calculate(
        purchase_value=extract_decimal(record["purchase_value"], "purchase_value"),
        reduction_percent=parse_percent_for_calculator(
            record["reduction_percent_value_down_payment"],
            "reduction_percent_value_down_payment",
        ),
        down_payment_percent=parse_percent_for_calculator(
            record["discount_percent"],
            "discount_percent",
        ),
        loan_period_years=extract_decimal(record["loan_period_in_year"], "loan_period_in_year"),
        monthly_principal_reduction_percent=parse_percent_for_calculator(
            record["monthly_principal_reduction_percent"],
            "monthly_principal_reduction_percent",
        ),
        interest_rate_percent=parse_percent_for_calculator(
            record["int_rate_percent"],
            "int_rate_percent",
        ),
        total_interest_reduction_percent=parse_percent_for_calculator(
            record["total_interest_reduction"],
            "total_interest_reduction",
        ),
        insurance_rate_percent=insurance_rate_percent,
    )

    return result, insurance_na


def build_excel_row(record: dict, result: LoanOutputs, insurance_na: bool) -> List[str]:
    total_interest_and_insurance = (
        f"{format_currency_trimmed(result.total_interest_after_reduction, spaces_after_dollar=3)} AND NA"
        if insurance_na
        else (
            f"{format_currency_trimmed(result.total_interest_after_reduction, spaces_after_dollar=3)} AND "
            f"{format_currency_trimmed(result.insurance_amount_per_month, spaces_after_dollar=2)}"
        )
    )

    return [
        f"SAMPLE NO. {record['sample_no']} ,RECORD.{record['record_no']}",
        format_upper_words(record["customer_reference_number"], spaces_between_words=3),
        format_upper_words(record["customer_name"], spaces_between_words=2),
        format_city_state(record["city"], record["state"]),
        (
            f"{format_currency_trimmed(result.purchase_value_reduction_amount, spaces_after_dollar=2)} AND "
            f"{format_percent(record['discount_percent'])}"
        ),
        format_loan_period_and_interest(record),
        format_upper_words(record["guarantor_name"], spaces_between_words=2),
        format_upper_words(record["guarantor_reference_number"], spaces_between_words=3),
        (
            f"{format_currency_trimmed(result.loan_amount, spaces_after_dollar=2)} AND "
            f"{format_currency_trimmed(result.principal_after_monthly_reduction, spaces_after_dollar=2)}"
        ),
        total_interest_and_insurance,
    ]


def build_roboflow_settings():
    args = SimpleNamespace(confidence=ROBOFLOW_CONFIDENCE, overlap=ROBOFLOW_OVERLAP)
    return load_roboflow_settings(args)


def extract_crop_paths(inference_result: dict) -> List[Path]:
    crop_paths: List[Path] = []
    for detection in inference_result.get("detections", []):
        crop_path = detection.get("crop_path")
        if not crop_path:
            continue
        path = Path(crop_path)
        if path.exists() and path.is_file():
            crop_paths.append(path)
    return crop_paths


def main() -> None:
    load_dotenv()

    api_key = os.getenv("GEMINI_API_KEY")
    model = os.getenv("GEMINI_MODEL", "gemini-2.5-flash")

    if not api_key:
        raise RuntimeError("GEMINI_API_KEY is missing in .env")

    run_logger = RunLogger.create(logs_dir=LOGS_DIR, prefix="final")
    run_logger.log(f"input_images_folder={INPUT_IMAGES_FOLDER}")
    run_logger.log(f"roboflow_output_root={ROBOFLOW_OUTPUT_ROOT}")
    run_logger.log(f"output_excel_path={OUTPUT_EXCEL_PATH}")
    run_logger.log(f"model={model}")
    print(f"Logs: {run_logger.file_path}")

    client = genai.Client(api_key=api_key)
    ensure_roboflow_dependencies()
    roboflow_settings = build_roboflow_settings()
    run_logger.log(f"roboflow_confidence={roboflow_settings.confidence}")
    run_logger.log(f"roboflow_overlap={roboflow_settings.overlap}")

    source_image_paths = get_image_paths(INPUT_IMAGES_FOLDER)
    output_dir = Path(ROBOFLOW_OUTPUT_ROOT).expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    processed_source_images = 0
    success_count = 0

    try:
        for source_image_path in source_image_paths:
            try:
                run_logger.section(f"SOURCE IMAGE START: {source_image_path.name}")
                inference_result = run_inference(
                    image_path=source_image_path.resolve(),
                    output_dir=output_dir,
                    settings=roboflow_settings,
                )
                processed_source_images += 1
            except Exception as exc:
                run_logger.section(f"ROBOFLOW FAILED: {source_image_path.name}")
                run_logger.log(f"error={exc}")
                print(f"Failed Roboflow for {source_image_path.name}: {exc}")
                continue

            result_dir = inference_result["result_dir"]
            crop_paths = extract_crop_paths(inference_result)

            run_logger.log(f"roboflow_output_dir={result_dir}")
            run_logger.log(f"roboflow_crop_count={len(crop_paths)}")
            run_logger.log_json(
                f"ROBOFLOW DETECTIONS: {source_image_path.name}",
                inference_result.get("detections", []),
            )

            if not crop_paths:
                message = (
                    f"No cropped records for {source_image_path.name}. "
                    f"Output folder: {result_dir}"
                )
                run_logger.log(message)
                print(message)
                continue

            print(
                f"Roboflow output for {source_image_path.name}: {result_dir} "
                f"({len(crop_paths)} crop image(s))"
            )

            for crop_path in crop_paths:
                try:
                    record = process_single_image(client, model, crop_path, logger=run_logger)
                    result, insurance_na = calculate_outputs(record)
                    row = build_excel_row(record, result, insurance_na)
                    append_row_to_excel(OUTPUT_EXCEL_PATH, row)
                    success_count += 1
                    run_logger.log_json(f"FINAL RECORD: {crop_path.name}", record)
                    print(f"Processed {crop_path.name}")
                except Exception as exc:
                    run_logger.section(f"CROP FAILED: {crop_path.name}")
                    run_logger.log(f"error={exc}")
                    print(f"Failed {crop_path.name}: {exc}")
    finally:
        run_logger.log(
            f"roboflow_processed_source_images={processed_source_images}/{len(source_image_paths)}"
        )
        run_logger.log(f"processed_cropped_record_images={success_count}")
        run_logger.log(f"output_excel={OUTPUT_EXCEL_PATH}")
        run_logger.close()

    print(
        f"Roboflow processed {processed_source_images}/{len(source_image_paths)} "
        f"source image(s)."
    )
    print(f"Processed {success_count} cropped record image(s).")
    print(f"Excel updated at: {OUTPUT_EXCEL_PATH}")


if __name__ == "__main__":
    main()
