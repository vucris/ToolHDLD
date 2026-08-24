# -*- coding: utf-8 -*-
"""
TL CONTRACT TOOL
Tao hop dong lao dong hang loat tu Excel + Word template va in PDF tong tren Windows.

Yeu cau Windows 10/11 + Microsoft Word Desktop.
Python packages: PySide6, openpyxl, python-docx, pypdf, pywin32
"""

from __future__ import annotations

import csv
import hashlib
import json
import logging
import os
import re
import shutil
import subprocess
import sys
import traceback
import unicodedata
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from openpyxl import load_workbook
from docx import Document
from pypdf import PdfReader, PdfWriter

try:
    from PySide6.QtCore import QThread, Signal, Qt, QUrl
    from PySide6.QtGui import QDesktopServices
    from PySide6.QtWidgets import (
        QApplication, QMainWindow, QWidget, QFileDialog, QMessageBox,
        QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QLineEdit,
        QPushButton, QTableWidget, QTableWidgetItem, QHeaderView,
        QProgressBar, QSpinBox, QComboBox, QPlainTextEdit, QGroupBox,
        QAbstractItemView, QSizePolicy
    )
except ImportError as exc:
    print("Thieu PySide6. Chay: pip install -r requirements.txt")
    raise

APP_NAME = "TL Contract Tool V9"
BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "tool_config.json"
LEGACY_EXCEL_NAME = "DANH_SACH_HDLD_MAU.xlsx"
LEGACY_TEMPLATE_NAME = "MAU_HDLD_TEMPLATE.docx"
UPDATED_TEMPLATE_NAME = "MAU_HDLD_TEMPLATE_CAP_NHAT.docx"
DEFAULT_EXCEL = BASE_DIR / "DANH_SACH_HDLD_MAU_CAP_NHAT.xlsx"
DEFAULT_TEMPLATE = BASE_DIR / LEGACY_TEMPLATE_NAME
DEFAULT_OUTPUT = BASE_DIR / "OUTPUT"

DEFAULT_BATCH_SIZE = 25
MAX_BATCH_SIZE = 25
MAX_PRINT_JOB_PAGES = 100
PRINT_GUARD_FILENAME = "print_guard.json"
TABLE_VISIBLE_ROWS = 10
TABLE_ROW_HEIGHT = 28

SHEET_NAME = "NHAN_SU_HDLD"

EXPECTED_COLUMNS = [
    "CHON", "STT", "MA_NV", "HO_TEN", "NGAY_SINH", "GIOI_TINH", "DAN_TOC",
    "DIA_CHI_THUONG_TRU", "CCCD", "NGAY_CAP_CCCD", "NOI_CAP_CCCD", "MA_SO_THUE",
    "DIEN_THOAI", "CHUC_DANH", "LOAI_HOP_DONG", "THOI_HAN_THANG", "TU_NGAY",
    "DEN_NGAY", "MUC_LUONG_CHINH", "SO_HOP_DONG"
]

REQUIRED_FIELDS = [
    "STT", "MA_NV", "HO_TEN", "NGAY_SINH", "GIOI_TINH",
    "DIA_CHI_THUONG_TRU", "CCCD", "NGAY_CAP_CCCD", "NOI_CAP_CCCD",
    "CHUC_DANH", "LOAI_HOP_DONG", "THOI_HAN_THANG", "TU_NGAY", "DEN_NGAY",
    "MUC_LUONG_CHINH", "SO_HOP_DONG"
]

# Placeholder ngan de template khong bi thay doi page-break/bocuc.
PLACEHOLDER_MAP = {
    "{{TEN}}": "HO_TEN",
    "{{NS}}": "NGAY_SINH",
    "{{GT}}": "GIOI_TINH",
    "{{DT}}": "DAN_TOC",
    "{{DCTT}}": "DIA_CHI_THUONG_TRU",
    "{{CCCD}}": "CCCD",
    "{{NCC}}": "NGAY_CAP_CCCD",
    "{{NCAP}}": "NOI_CAP_CCCD",
    "{{MST}}": "MA_SO_THUE",
    "{{SDT}}": "DIEN_THOAI",
    "{{CDANH}}": "CHUC_DANH",
    "{{LOAIHD}}": "LOAI_HOP_DONG",
    "{{THANG}}": "THOI_HAN_THANG",
    "{{TUNGAY}}": "TU_NGAY",
    "{{DENNGAY}}": "DEN_NGAY",
    "{{LUONG}}": "MUC_LUONG_CHINH",
}

# Hien thi theo dung thu tu du lieu trong Excel. Cot CHON duoc thay bang
# checkbox "CHON IN" tren giao dien; TRANG_THAI la cot bo sung cua Tool.
TABLE_COLUMNS = [
    ("STT", "STT"),
    ("MA_NV", "Mã NV"),
    ("HO_TEN", "Họ tên"),
    ("NGAY_SINH", "Ngày sinh"),
    ("GIOI_TINH", "Giới tính"),
    ("DAN_TOC", "Dân tộc"),
    ("DIA_CHI_THUONG_TRU", "Địa chỉ thường trú"),
    ("CCCD", "CCCD"),
    ("NGAY_CAP_CCCD", "Ngày cấp CCCD"),
    ("NOI_CAP_CCCD", "Nơi cấp CCCD"),
    ("MA_SO_THUE", "Mã số thuế"),
    ("DIEN_THOAI", "Điện thoại"),
    ("CHUC_DANH", "Chức danh"),
    ("LOAI_HOP_DONG", "Loại HĐ"),
    ("THOI_HAN_THANG", "Thời hạn (tháng)"),
    ("TU_NGAY", "Từ ngày"),
    ("DEN_NGAY", "Đến ngày"),
    ("MUC_LUONG_CHINH", "Mức lương chính"),
    ("SO_HOP_DONG", "Số hợp đồng"),
]

DATE_FIELDS = {"NGAY_SINH", "NGAY_CAP_CCCD", "TU_NGAY", "DEN_NGAY"}


@dataclass
class Employee:
    row_number: int
    data: Dict[str, Any]
    selected: bool = False
    valid: bool = True
    error: str = ""

    @property
    def stt(self) -> int:
        try:
            return int(self.data.get("STT") or 0)
        except Exception:
            return 0

    @property
    def ma_nv(self) -> str:
        return str(self.data.get("MA_NV") or "").strip()

    @property
    def ho_ten(self) -> str:
        return str(self.data.get("HO_TEN") or "").strip()


def configure_logging(output_dir: Path) -> None:
    log_dir = output_dir / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    root = logging.getLogger()
    root.setLevel(logging.INFO)
    for handler in list(root.handlers):
        root.removeHandler(handler)
    fmt = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
    fh = logging.FileHandler(log_dir / "app.log", encoding="utf-8")
    fh.setFormatter(fmt)
    root.addHandler(fh)


def is_selected(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    if value is None:
        return False
    return str(value).strip().upper() in {"X", "TRUE", "1", "YES", "Y", "CO", "CÓ"}


def text_value(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def format_date(value: Any) -> str:
    if value in (None, ""):
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d/%m/%Y")
    if isinstance(value, date):
        return value.strftime("%d/%m/%Y")
    if isinstance(value, (int, float)):
        # openpyxl normally converts actual Excel dates to datetime when format is date.
        return text_value(value)
    s = str(value).strip()
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt).strftime("%d/%m/%Y")
        except ValueError:
            pass
    return s


def format_money(value: Any) -> str:
    if value in (None, ""):
        return ""
    try:
        if isinstance(value, str):
            clean = re.sub(r"[^0-9.-]", "", value)
            number = float(clean)
        else:
            number = float(value)
        if number.is_integer():
            return f"{int(number):,}".replace(",", ".")
        return f"{number:,.0f}".replace(",", ".")
    except Exception:
        return str(value).strip()


def sanitize_filename(text: str) -> str:
    text = unicodedata.normalize("NFD", text)
    text = "".join(c for c in text if unicodedata.category(c) != "Mn")
    text = text.replace("đ", "d").replace("Đ", "D")
    text = text.upper().strip()
    text = re.sub(r"[^A-Z0-9]+", "_", text)
    return text.strip("_") or "NHAN_VIEN"


def employee_filename(emp: Employee, ext: str) -> str:
    return f"{emp.stt:03d}_{sanitize_filename(emp.ma_nv)}_{sanitize_filename(emp.ho_ten)}.{ext}"


def validate_employee_batch(employees: List[Employee], max_size: int = MAX_BATCH_SIZE) -> None:
    """Validate one generation/print batch before Word or the printer is used."""
    problems: List[str] = []
    if len(employees) > max_size:
        problems.append(
            f"Đã tick {len(employees)} nhân viên, vượt giới hạn an toàn {max_size} nhân viên/lô."
        )

    checks = (
        ("STT", lambda emp: str(emp.stt), "STT"),
        ("MA_NV", lambda emp: emp.ma_nv.casefold(), "Mã nhân viên"),
    )
    for _, getter, label in checks:
        seen: Dict[str, Employee] = {}
        duplicates: List[str] = []
        for emp in employees:
            value = getter(emp).strip()
            if not value:
                continue
            previous = seen.get(value)
            if previous is not None:
                duplicates.append(
                    f"- {label} {text_value(emp.data.get('STT') if label == 'STT' else emp.ma_nv)}: "
                    f"dòng Excel {previous.row_number} và {emp.row_number}"
                )
            else:
                seen[value] = emp
        if duplicates:
            problems.append(f"{label} bị trùng:\n" + "\n".join(duplicates[:20]))

    if problems:
        raise ValueError(
            "Không thể sinh/in lô này vì dữ liệu chưa an toàn:\n\n"
            + "\n\n".join(problems)
        )


def validate_output_targets(employees: List[Employee], word_dir: Path, pdf_dir: Path) -> None:
    """Stop before generation if a selected employee would reuse an output file."""
    duplicate_rows: List[str] = []
    existing_files: List[str] = []
    seen_pdf_names: Dict[str, Employee] = {}

    for emp in employees:
        pdf_name = employee_filename(emp, "pdf")
        key = pdf_name.casefold()
        previous = seen_pdf_names.get(key)
        if previous is not None:
            duplicate_rows.append(
                f"- {pdf_name}: dòng Excel {previous.row_number} và {emp.row_number}"
            )
        else:
            seen_pdf_names[key] = emp

        for path in (
            word_dir / employee_filename(emp, "docx"),
            pdf_dir / pdf_name,
        ):
            if path.exists():
                existing_files.append(f"- {path.name} ({path.parent})")

    problems: List[str] = []
    if duplicate_rows:
        problems.append(
            "Trùng tên file giữa các nhân viên được tick:\n"
            + "\n".join(duplicate_rows[:20])
        )
    if existing_files:
        problems.append(
            "File nhân viên đã tồn tại, Tool sẽ không ghi đè:\n"
            + "\n".join(existing_files[:20])
        )
    if problems:
        raise FileExistsError(
            "\n\n".join(problems)
            + "\n\nHãy kiểm tra dữ liệu hoặc bấm 'XÓA FILE ĐÃ TẠO' rồi thực hiện lại."
        )


def read_employees(excel_path: Path) -> List[Employee]:
    if not excel_path.exists():
        raise FileNotFoundError(f"Khong tim thay file Excel: {excel_path}")
    if excel_path.suffix.lower() != ".xlsx":
        raise ValueError("Tool V1 chi ho tro file .xlsx")

    wb = load_workbook(excel_path, data_only=True, read_only=True)
    try:
        if SHEET_NAME not in wb.sheetnames:
            raise ValueError(f"Khong tim thay sheet '{SHEET_NAME}'")
        ws = wb[SHEET_NAME]
        headers = [text_value(c.value).upper() for c in ws[1]]
        missing_cols = [c for c in EXPECTED_COLUMNS if c not in headers]
        if missing_cols:
            raise ValueError("Excel thieu cot: " + ", ".join(missing_cols))
        idx = {name: headers.index(name) for name in EXPECTED_COLUMNS}

        employees: List[Employee] = []
        for row_no, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if not any(v not in (None, "") for v in row):
                continue
            data = {col: row[idx[col]] if idx[col] < len(row) else None for col in EXPECTED_COLUMNS}
            emp = Employee(
                row_number=row_no,
                data=data,
                selected=is_selected(data.get("CHON")),
            )
            missing = [field for field in REQUIRED_FIELDS if data.get(field) in (None, "")]
            if missing:
                emp.valid = False
                emp.error = "Thieu: " + ", ".join(missing)
            elif emp.stt <= 0:
                emp.valid = False
                emp.error = "STT khong hop le"
            employees.append(emp)

        duplicate_checks = (
            ("STT", lambda emp: str(emp.stt) if emp.stt > 0 else ""),
            ("MA_NV", lambda emp: emp.ma_nv.casefold()),
        )
        for label, getter in duplicate_checks:
            grouped: Dict[str, List[Employee]] = {}
            for emp in employees:
                value = getter(emp).strip()
                if value:
                    grouped.setdefault(value, []).append(emp)
            for group in grouped.values():
                if len(group) < 2:
                    continue
                display_value = group[0].stt if label == "STT" else group[0].ma_nv
                for emp in group:
                    duplicate_error = f"Trùng {label}: {display_value}"
                    emp.valid = False
                    emp.error = (
                        emp.error + "; " + duplicate_error
                        if emp.error else duplicate_error
                    )
        employees.sort(key=lambda e: (e.stt, e.ma_nv))
        return employees
    finally:
        wb.close()


def replacement_values(emp: Employee) -> Dict[str, str]:
    d = emp.data
    start_date = d.get("TU_NGAY")
    parsed_start_date: Optional[date] = None
    if isinstance(start_date, datetime):
        parsed_start_date = start_date.date()
    elif isinstance(start_date, date):
        parsed_start_date = start_date
    elif start_date not in (None, ""):
        raw_date = str(start_date).strip()
        for pattern in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
            try:
                parsed_start_date = datetime.strptime(raw_date, pattern).date()
                break
            except ValueError:
                continue
    return {
        "{{TEN}}": text_value(d.get("HO_TEN")),
        "{{NS}}": format_date(d.get("NGAY_SINH")),
        "{{GT}}": text_value(d.get("GIOI_TINH")),
        "{{DT}}": text_value(d.get("DAN_TOC")),
        "{{DCTT}}": text_value(d.get("DIA_CHI_THUONG_TRU")),
        "{{CCCD}}": text_value(d.get("CCCD")),
        "{{NCC}}": format_date(d.get("NGAY_CAP_CCCD")),
        "{{NCAP}}": text_value(d.get("NOI_CAP_CCCD")),
        "{{MST}}": text_value(d.get("MA_SO_THUE")),
        "{{SDT}}": text_value(d.get("DIEN_THOAI")),
        "{{CDANH}}": text_value(d.get("CHUC_DANH")),
        "{{LOAIHD}}": text_value(d.get("LOAI_HOP_DONG")),
        "{{THANG}}": text_value(d.get("THOI_HAN_THANG")),
        "{{TUNGAY}}": format_date(d.get("TU_NGAY")),
        "{{DENNGAY}}": format_date(d.get("DEN_NGAY")),
        "{{LUONG}}": format_money(d.get("MUC_LUONG_CHINH")),
        "{{SOHD}}": text_value(d.get("SO_HOP_DONG")),
        "{{NGAYKY}}": parsed_start_date.strftime("%d") if parsed_start_date else "",
        "{{THANGKY}}": parsed_start_date.strftime("%m") if parsed_start_date else "",
        "{{NAMKY}}": parsed_start_date.strftime("%Y") if parsed_start_date else "",
    }


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def build_print_job_key(employees: List[Employee], template_path: Path,
                        copies: int, duplex: bool) -> str:
    payload = {
        "template_sha256": sha256_file(template_path),
        "copies": int(copies),
        "duplex": bool(duplex),
        "employees": [
            {
                "stt": emp.stt,
                "ma_nv": emp.ma_nv,
                "values": replacement_values(emp),
            }
            for emp in employees
        ],
    }
    raw = json.dumps(payload, ensure_ascii=False, sort_keys=True).encode("utf-8")
    return hashlib.sha256(raw).hexdigest()


def _print_guard_path(output_dir: Path) -> Path:
    return output_dir / "logs" / PRINT_GUARD_FILENAME


def load_print_guard(output_dir: Path) -> Dict[str, Any]:
    path = _print_guard_path(output_dir)
    if not path.exists():
        return {"version": 1, "jobs": {}}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(data, dict) or not isinstance(data.get("jobs"), dict):
            raise ValueError("Cau truc print_guard.json khong hop le")
        return data
    except Exception as exc:
        raise RuntimeError(
            f"Không đọc được khóa chống in trùng: {path}\n{exc}\n"
            "Tool dừng in để tránh gửi lệnh trùng."
        ) from exc


def get_print_guard_record(output_dir: Path, job_key: str) -> Optional[Dict[str, Any]]:
    record = load_print_guard(output_dir).get("jobs", {}).get(job_key)
    return record if isinstance(record, dict) else None


def save_print_guard_record(output_dir: Path, job_key: str,
                            record: Dict[str, Any]) -> None:
    path = _print_guard_path(output_dir)
    path.parent.mkdir(parents=True, exist_ok=True)
    data = load_print_guard(output_dir)
    jobs = data.setdefault("jobs", {})
    previous = jobs.get(job_key) if isinstance(jobs.get(job_key), dict) else {}
    attempt = int(previous.get("attempt", 0)) + (1 if record.get("status") == "SENDING" else 0)
    merged = dict(previous)
    merged.update(record)
    merged["attempt"] = max(attempt, 1)
    merged["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    jobs[job_key] = merged

    temp_path = path.with_name(f".{path.name}.{os.getpid()}.tmp")
    try:
        temp_path.write_text(
            json.dumps(data, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        os.replace(temp_path, path)
    finally:
        if temp_path.exists():
            try:
                temp_path.unlink()
            except Exception:
                pass


def iter_all_paragraphs(doc: Document):
    for p in doc.paragraphs:
        yield p
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p
    for section in doc.sections:
        for p in section.header.paragraphs:
            yield p
        for p in section.footer.paragraphs:
            yield p


def replace_token_in_paragraph(paragraph, token: str, value: str) -> bool:
    """Replace token even if Word splits it across runs, preserving surrounding formatting."""
    full = "".join(run.text for run in paragraph.runs)
    if token not in full:
        return False

    changed = False
    while token in full:
        start = full.index(token)
        end = start + len(token)
        positions: List[Tuple[int, int, int]] = []
        cursor = 0
        for i, run in enumerate(paragraph.runs):
            positions.append((i, cursor, cursor + len(run.text)))
            cursor += len(run.text)

        start_run = end_run = None
        for i, a, b in positions:
            if start_run is None and a <= start < b:
                start_run = i
            if a < end <= b:
                end_run = i
                break
        if start_run is None:
            return changed
        if end_run is None:
            end_run = start_run

        runs = paragraph.runs
        start_a = positions[start_run][1]
        end_a = positions[end_run][1]
        before = runs[start_run].text[: start - start_a]
        after = runs[end_run].text[end - end_a :]
        runs[start_run].text = before + value + after
        runs[start_run].font.highlight_color = None
        for i in range(start_run + 1, end_run + 1):
            runs[i].text = ""
            runs[i].font.highlight_color = None
        changed = True
        full = "".join(run.text for run in paragraph.runs)
    return changed


def create_contract_docx(template_path: Path, emp: Employee, output_docx: Path) -> None:
    doc = Document(template_path)
    values = replacement_values(emp)
    found = {token: False for token in PLACEHOLDER_MAP}
    literal_replacements = {
        "Số: ......./202..../HĐLĐ-TLNT": (
            f"Số: {values['{{SOHD}}']}", "Số hợp đồng"
        ),
        "Hôm nay, ngày ..... tháng ..... năm 202...": (
            "Hôm nay, ngày "
            f"{values['{{NGAYKY}}']} tháng {values['{{THANGKY}}']} "
            f"năm {values['{{NAMKY}}']}",
            "Ngày ký hợp đồng",
        ),
    }
    literal_found = {source: False for source in literal_replacements}
    for paragraph in iter_all_paragraphs(doc):
        # Mẫu Word gốc được giữ nguyên. Hai vùng dấu chấm được thay trực tiếp
        # trong bản DOCX sinh ra, không cần sửa/chèn placeholder vào template.
        for source, (replacement, _label) in literal_replacements.items():
            if replace_token_in_paragraph(paragraph, source, replacement):
                literal_found[source] = True
        for token in PLACEHOLDER_MAP:
            value = values[token]
            if replace_token_in_paragraph(paragraph, token, value):
                found[token] = True
        # Contract output must not contain yellow highlight from input fields.
        for run in paragraph.runs:
            if run.font.highlight_color is not None:
                run.font.highlight_color = None

    missing_tokens = [token for token, ok in found.items() if not ok]
    missing_literals = [
        label
        for source, (_replacement, label) in literal_replacements.items()
        if not literal_found[source]
    ]
    if missing_tokens or missing_literals:
        details = []
        if missing_tokens:
            details.append("placeholder: " + ", ".join(missing_tokens))
        if missing_literals:
            details.append("vùng dữ liệu: " + ", ".join(missing_literals))
        raise ValueError("Template thiếu " + "; ".join(details))

    output_docx.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_docx)


def _word_story_ranges(doc):
    """Yield all Word story ranges (main body, headers, footers, text frames, etc.)."""
    # Word WdStoryType values are 1..17. Some story types may not exist in a document.
    for story_type in range(1, 18):
        try:
            rng = doc.StoryRanges.Item(story_type)
        except Exception:
            continue
        while rng is not None:
            yield rng
            try:
                rng = rng.NextStoryRange
            except Exception:
                break


def _word_replace_token(doc, token: str, value: str) -> bool:
    """Replace a token everywhere in Word with wildcard search explicitly disabled."""
    found_any = False
    for rng in _word_story_ranges(doc):
        try:
            # Work on a duplicate so Find does not disturb the story range iterator.
            work = rng.Duplicate
            find = work.Find
            find.ClearFormatting()
            find.Replacement.ClearFormatting()
            find.Text = token
            find.Replacement.Text = value
            find.Forward = True
            find.Wrap = 1                 # wdFindContinue
            find.Format = False
            find.MatchCase = False
            find.MatchWholeWord = False
            find.MatchWildcards = False  # IMPORTANT: {{TOKEN}} contains braces
            find.MatchSoundsLike = False
            find.MatchAllWordForms = False
            if find.Execute(Replace=2):   # wdReplaceAll
                found_any = True
        except Exception:
            # Continue with the remaining Word stories. Final validation below
            # will stop the export if any placeholder is still present.
            pass
    return found_any


def _word_remaining_tokens(doc, tokens: List[str]) -> List[str]:
    """Return placeholders that are still present anywhere in the Word document."""
    remaining = set()
    for rng in _word_story_ranges(doc):
        try:
            txt = str(rng.Text or '')
        except Exception:
            continue
        for token in tokens:
            if token in txt:
                remaining.add(token)
    return sorted(remaining)


def _word_clear_highlight(doc) -> None:
    for rng in _word_story_ranges(doc):
        try:
            rng.HighlightColorIndex = 0  # wdNoHighlight
        except Exception:
            pass



def _word_find_main_range(doc, text: str, match_case: bool = False):
    """Find text in the main document story and return the found Word Range."""
    rng = doc.Content.Duplicate
    find = rng.Find
    find.ClearFormatting()
    find.Text = text
    find.Forward = True
    find.Wrap = 0  # wdFindStop
    find.Format = False
    find.MatchCase = match_case
    find.MatchWholeWord = False
    find.MatchWildcards = False
    if find.Execute():
        return rng
    return None


def _word_fix_signature_layout(doc, employee_name: str) -> None:
    """
    V7 - Co dinh khoi chu ky Giam doc tren cung trang voi dong "GIÁM ĐỐC".

    Thay vi co gang xoa cac dong trong o trang sau, V7 dua ten
    "PHẠM QUÝ TRỌNG" vao CUNG MỘT PARAGRAPH voi "GIÁM ĐỐC" bang manual line-break.
    Cach nay on dinh hon voi template co bang/column va khong bi Word day ten sang trang moi.
    """
    DIRECTOR_NAME = "PHẠM QUÝ TRỌNG"
    DIRECTOR_HEADING = "GIÁM ĐỐC"
    EMPLOYEE_INSTRUCTION = "(Ký, ghi rõ họ tên)"
    employee_name = text_value(employee_name)

    try:
        doc.Repaginate()
        heading = _word_find_main_range(doc, DIRECTOR_HEADING, match_case=True)
        if heading is None:
            logging.warning("Khong tim thay dong GIÁM ĐỐC de can chu ky")
            return

        # Chuan hoa paragraph chua GIÁM ĐỐC.
        try:
            p = heading.Paragraphs.Item(1)
            p.Format.PageBreakBefore = False
            p.Format.KeepWithNext = False
            p.Format.KeepTogether = False
            p.Format.SpaceBefore = 0
            p.Format.SpaceAfter = 0
            p.Format.Alignment = 1  # wdAlignParagraphCenter
        except Exception:
            pass

        # Tim CHI ten ky o PHIA SAU dong GIÁM ĐỐC.
        # Khong cham vao ten giam doc o phan thong tin Ben A tren trang dau.
        def find_signature_name():
            try:
                rng = doc.Range(Start=heading.End, End=doc.Content.End)
                f = rng.Find
                f.ClearFormatting()
                f.Text = DIRECTOR_NAME
                f.Forward = True
                f.Wrap = 0
                f.Format = False
                f.MatchCase = True
                f.MatchWholeWord = False
                f.MatchWildcards = False
                if f.Execute():
                    return rng
            except Exception:
                return None
            return None

        old_name = find_signature_name()
        if old_name is not None:
            # Xoa paragraph ten cu neu no dung mot dong rieng.
            # Dieu nay thuong la nguyen nhan tao trang 5 trong template cu.
            try:
                para = old_name.Paragraphs.Item(1)
                raw = str(para.Range.Text or "")
                clean = raw.replace("\r", "").replace("\x07", "").replace("\x0b", "").strip()
                if clean == DIRECTOR_NAME:
                    para.Range.Delete()
                else:
                    old_name.Text = ""
            except Exception:
                try:
                    old_name.Text = ""
                except Exception:
                    pass

        # Neu sau khi xoa ten cu van con cac paragraph/page-break rong o CUOI tai lieu,
        # thu don dep toi da 12 paragraph rong. Khong xoa paragraph co noi dung.
        for _ in range(12):
            try:
                doc.Repaginate()
                last_para = doc.Paragraphs.Item(doc.Paragraphs.Count)
                raw = str(last_para.Range.Text or "")
                clean = raw.replace("\r", "").replace("\x07", "").replace("\x0b", "").replace("\x0c", "").strip()
                # Chi xoa dong rong / page-break rong o cuoi file.
                if clean == "":
                    last_para.Range.Delete()
                else:
                    break
            except Exception:
                break

        # Tim lai heading vi range co the thay doi sau khi xoa paragraph.
        heading = _word_find_main_range(doc, DIRECTOR_HEADING, match_case=True)
        if heading is None:
            return

        try:
            p = heading.Paragraphs.Item(1)
            p.Format.PageBreakBefore = False
            p.Format.KeepWithNext = False
            p.Format.KeepTogether = False
            p.Format.SpaceBefore = 0
            p.Format.SpaceAfter = 0
            p.Format.Alignment = 1
        except Exception:
            pass

        # Chen ten vao cung paragraph. Bat dau voi khoang ky rong 7 dong;
        # neu Word van day sang trang sau thi tu dong giam dan den khi cung trang.
        heading_page = int(heading.Information(3))  # wdActiveEndPageNumber
        success = False

        for blank_lines in (7, 6, 5, 4, 3):
            # Bao dam khong con ban chen thu truoc.
            candidate = find_signature_name()
            if candidate is not None:
                try:
                    candidate.Text = ""
                except Exception:
                    pass

            heading = _word_find_main_range(doc, DIRECTOR_HEADING, match_case=True)
            if heading is None:
                break

            insert_at = doc.Range(Start=heading.End, End=heading.End)
            payload = ("\x0b" * blank_lines) + DIRECTOR_NAME  # manual line breaks
            insert_at.InsertAfter(payload)

            # Format ten vua chen.
            sig = find_signature_name()
            if sig is not None:
                try:
                    sig.Font.Bold = True
                    sig.Font.Name = "Times New Roman"
                    sig.Font.Size = 12
                    sig.Paragraphs.Item(1).Format.Alignment = 1
                    sig.Paragraphs.Item(1).Format.KeepTogether = False
                    sig.Paragraphs.Item(1).Format.KeepWithNext = False
                    sig.Paragraphs.Item(1).Format.PageBreakBefore = False
                except Exception:
                    pass

                doc.Repaginate()
                try:
                    name_page = int(sig.Information(3))
                except Exception:
                    name_page = heading_page + 1

                if name_page == heading_page:
                    success = True
                    break

                # Chua vua: xoa ten vua chen va thu it dong trong hon.
                try:
                    sig.Text = ""
                except Exception:
                    pass

        if not success:
            # Phuong an cuoi: 2 dong trong, uu tien chac chan cung trang.
            heading = _word_find_main_range(doc, DIRECTOR_HEADING, match_case=True)
            if heading is not None:
                insert_at = doc.Range(Start=heading.End, End=heading.End)
                insert_at.InsertAfter(("\x0b" * 2) + DIRECTOR_NAME)
                sig = find_signature_name()
                if sig is not None:
                    try:
                        sig.Font.Bold = True
                        sig.Font.Name = "Times New Roman"
                        sig.Font.Size = 12
                        sig.Paragraphs.Item(1).Format.Alignment = 1
                    except Exception:
                        pass

        # Đặt họ tên người lao động ở cột chữ ký bên trái, cùng hàng ngang với
        # tên Giám đốc. Word có thể thay đổi phân trang theo máy, nên đo vị trí
        # thật của hai tên và tự chọn số dòng trống gần nhất thay vì cố định.
        director_signature = find_signature_name()
        instruction = _word_find_main_range(doc, EMPLOYEE_INSTRUCTION, match_case=True)
        if not employee_name:
            raise ValueError("Họ tên người lao động trống, không thể tạo khối chữ ký")
        if director_signature is None or instruction is None:
            raise RuntimeError("Không tìm thấy khối chữ ký để căn tên người lao động")

        doc.Repaginate()
        director_page = int(director_signature.Information(3))
        director_y = float(director_signature.Information(6))  # wdVerticalPositionRelativeToPage
        best_blank_lines = None
        best_score = float("inf")
        best_employee_y = None
        best_employee_page = None

        for blank_lines in range(2, 10):
            instruction = _word_find_main_range(doc, EMPLOYEE_INSTRUCTION, match_case=True)
            if instruction is None:
                break
            insert_start = int(instruction.End)
            payload = ("\x0b" * blank_lines) + employee_name
            doc.Range(Start=insert_start, End=insert_start).InsertAfter(payload)
            employee_signature = doc.Range(
                Start=insert_start + blank_lines,
                End=insert_start + len(payload),
            )
            employee_signature.Font.Bold = True
            employee_signature.Font.Italic = False
            employee_signature.Font.Name = "Times New Roman"
            employee_signature.Font.Size = 12
            employee_signature.Paragraphs.Item(1).Format.Alignment = 1
            employee_signature.Paragraphs.Item(1).Format.KeepTogether = False
            employee_signature.Paragraphs.Item(1).Format.KeepWithNext = False
            employee_signature.Paragraphs.Item(1).Format.PageBreakBefore = False
            employee_signature.Paragraphs.Item(1).Format.SpaceBefore = 0
            employee_signature.Paragraphs.Item(1).Format.SpaceAfter = 0

            doc.Repaginate()
            employee_page = int(employee_signature.Information(3))
            employee_y = float(employee_signature.Information(6))
            score = abs(employee_y - director_y)
            if employee_page != director_page:
                score += 10000
            if score < best_score:
                best_score = score
                best_blank_lines = blank_lines
                best_employee_y = employee_y
                best_employee_page = employee_page

            # Xóa toàn bộ bản chèn thử, gồm cả các manual line-break.
            doc.Range(Start=insert_start, End=insert_start + len(payload)).Delete()

        if best_blank_lines is None or best_score > 8:
            raise RuntimeError(
                "Không thể căn họ tên người lao động ngang với tên Giám đốc "
                f"(sai lệch {best_score:.1f} pt; dòng {best_blank_lines}; "
                f"NLĐ trang {best_employee_page} y={best_employee_y}; "
                f"Giám đốc trang {director_page} y={director_y})"
            )

        instruction = _word_find_main_range(doc, EMPLOYEE_INSTRUCTION, match_case=True)
        if instruction is None:
            raise RuntimeError("Mất vị trí chèn tên người lao động sau khi căn chữ ký")
        insert_start = int(instruction.End)
        payload = ("\x0b" * best_blank_lines) + employee_name
        doc.Range(Start=insert_start, End=insert_start).InsertAfter(payload)
        employee_signature = doc.Range(
            Start=insert_start + best_blank_lines,
            End=insert_start + len(payload),
        )
        employee_signature.Font.Bold = True
        employee_signature.Font.Italic = False
        employee_signature.Font.Name = "Times New Roman"
        employee_signature.Font.Size = 12
        employee_signature.Paragraphs.Item(1).Format.Alignment = 1
        employee_signature.Paragraphs.Item(1).Format.KeepTogether = False
        employee_signature.Paragraphs.Item(1).Format.KeepWithNext = False
        employee_signature.Paragraphs.Item(1).Format.PageBreakBefore = False
        employee_signature.Paragraphs.Item(1).Format.SpaceBefore = 0
        employee_signature.Paragraphs.Item(1).Format.SpaceAfter = 0

        doc.Repaginate()
        final_employee_page = int(employee_signature.Information(3))
        final_employee_y = float(employee_signature.Information(6))
        director_signature = find_signature_name()
        if director_signature is None:
            raise RuntimeError("Không tìm thấy lại tên Giám đốc sau khi căn chữ ký")
        final_director_page = int(director_signature.Information(3))
        final_director_y = float(director_signature.Information(6))
        if (
            final_employee_page != final_director_page
            or abs(final_employee_y - final_director_y) > 8
        ):
            raise RuntimeError(
                "Tên người lao động chưa nằm ngang tên Giám đốc; dừng tạo PDF để tránh in sai"
            )

        # Can lai ten nguoi lao dong / nguoi su dung lao dong neu nam trong table.
        # Khong thay noi dung, chi can giua cac paragraph cua khoi chu ky.
        try:
            for label in ("NGƯỜI LAO ĐỘNG", "NGƯỜI SỬ DỤNG LAO ĐỘNG", DIRECTOR_HEADING):
                r = _word_find_main_range(doc, label, match_case=True)
                if r is not None:
                    r.Paragraphs.Item(1).Format.Alignment = 1
        except Exception:
            pass

        doc.Repaginate()

    except Exception as exc:
        logging.exception("Khong tu dong can duoc khoi chu ky V7")
        raise RuntimeError(
            f"Không thể hoàn thiện khối chữ ký cho nhân viên {employee_name}: {exc}"
        ) from exc


def _word_force_a4(doc) -> None:
    """Dam bao cac section trong hop dong dung kho giay A4."""
    try:
        for i in range(1, doc.Sections.Count + 1):
            try:
                doc.Sections.Item(i).PageSetup.PaperSize = 7  # wdPaperA4
            except Exception:
                pass
    except Exception:
        pass

def generate_contracts_word_com(employees: List[Employee], template_path: Path,
                                word_dir: Path, pdf_dir: Path, progress=None) -> List[Path]:
    """
    V5 - On dinh hon tren Windows:
    - python-docx thay du lieu Excel vao template Word.
    - Microsoft Word COM chi dung de mo DOCX da co du lieu va export PDF.
    - Neu con placeholder thi dung ngay, khong tao PDF de tranh in sai.
    """
    if os.name != "nt":
        raise RuntimeError("Microsoft Word COM chi ho tro Windows.")

    try:
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise RuntimeError("Thieu pywin32. Chay: pip install pywin32") from exc

    word_dir.mkdir(parents=True, exist_ok=True)
    pdf_dir.mkdir(parents=True, exist_ok=True)
    validate_output_targets(employees, word_dir, pdf_dir)

    word = None
    keeper_doc = None
    pdfs: List[Path] = []
    created_files: List[Path] = []

    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

        # Mot so cau hinh Word tu dong ket thuc Word.Application khi tai lieu
        # cuoi cung bi dong. Giu mot document trong nen de phien COM con song
        # trong suot ca lo; neu khong, nhan vien thu hai se loi RPC failed.
        keeper_doc = word.Documents.Add()
        total = len(employees)

        for i, emp in enumerate(employees, start=1):
            out_docx = word_dir / employee_filename(emp, "docx")
            out_pdf = pdf_dir / employee_filename(emp, "pdf")

            # Khong ghi de file nhan vien. Kiem tra lai tai thoi diem tao de
            # tranh truong hop file xuat hien sau buoc preflight.
            if out_docx.exists() or out_pdf.exists():
                existing = out_pdf if out_pdf.exists() else out_docx
                raise FileExistsError(
                    f"File nhân viên đã tồn tại, không thể ghi đè: {existing}"
                )

            # BUOC 1: thay placeholder bang python-docx.
            created_files.append(out_docx)
            create_contract_docx(template_path, emp, out_docx)

            # BUOC 2: kiem tra lai file Word vua sinh.
            check_doc = Document(out_docx)
            remaining = set()
            values = replacement_values(emp)

            for paragraph in iter_all_paragraphs(check_doc):
                full_text = "".join(run.text for run in paragraph.runs)
                for token in values.keys():
                    if token in full_text:
                        remaining.add(token)

            if remaining:
                raise ValueError(
                    "Hop dong Word van con placeholder sau khi nap Excel: "
                    + ", ".join(sorted(remaining))
                    + ". Khong tao PDF de tranh in sai."
                )

            # BUOC 3: Word Desktop chi dung de export PDF,
            # giu bo cuc giong file template tot nhat.
            # Ghi nhan truoc de neu Word loi giua luc export, file PDF dang do
            # cung duoc don sach va khong chan lan chay tiep theo.
            created_files.append(out_pdf)
            doc = word.Documents.Open(
                str(out_docx.resolve()),
                ReadOnly=False,
                AddToRecentFiles=False
            )
            try:
                try:
                    doc.Content.HighlightColorIndex = 0
                except Exception:
                    pass

                # V7: co dinh ten Giam doc trong cung paragraph chu ky va ep kho giay A4.
                _word_force_a4(doc)
                _word_fix_signature_layout(doc, emp.ho_ten)

                doc.Save()
                doc.ExportAsFixedFormat(
                    str(out_pdf.resolve()),
                    17  # wdExportFormatPDF
                )
            finally:
                doc.Close(False)

            if not out_pdf.exists() or out_pdf.stat().st_size == 0:
                raise RuntimeError(f"Khong tao duoc PDF cho {emp.ho_ten}")

            removed_pages = remove_trailing_blank_pdf_pages(out_pdf)
            if removed_pages:
                logging.info(
                    "Da bo %s trang trang cuoi PDF cua %s",
                    removed_pages,
                    emp.ho_ten,
                )

            pdfs.append(out_pdf)

            if progress:
                progress(
                    i, total,
                    f"Dang tao Word/PDF {i}/{total}: {emp.ho_ten}"
                )

        return pdfs

    except Exception as exc:
        cleanup_failed: List[str] = []
        for path in reversed(created_files):
            try:
                if path.exists():
                    path.unlink()
            except Exception as cleanup_exc:
                cleanup_failed.append(f"{path.name}: {cleanup_exc}")
        if cleanup_failed:
            logging.error(
                "Khong don sach duoc file dang do: %s",
                " | ".join(cleanup_failed),
            )
        raise RuntimeError(
            "Khong the tao hop dong bang Microsoft Word.\n"
            + str(exc)
            + (
                "\nKhong xoa duoc mot so file dang do: " + "; ".join(cleanup_failed)
                if cleanup_failed else ""
            )
        ) from exc

    finally:
        if keeper_doc is not None:
            try:
                keeper_doc.Close(False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass


def calculate_master_page_count(pdf_files: List[Path], copies: int,
                                duplex: bool = True) -> int:
    if copies < 1:
        raise ValueError("Số bản phải >= 1")
    total = 0
    for pdf in pdf_files:
        reader = PdfReader(str(pdf))
        if reader.is_encrypted:
            raise ValueError(f"PDF bị khóa: {pdf.name}")
        page_count = len(reader.pages)
        if page_count == 0:
            raise ValueError(f"PDF không có trang: {pdf.name}")
        pages_per_copy = page_count + (1 if duplex and page_count % 2 == 1 else 0)
        total += pages_per_copy * copies
    return total


def create_master_pdf(pdf_files: List[Path], copies: int, output_path: Path,
                      duplex: bool = True, progress=None) -> int:
    """
    Ghep PDF tong.

    Khi in 2 mat:
    - Moi hop dong/moi ban phai bat dau o MAT TRUOC cua mot to A4.
    - Neu hop dong co so trang le, tu dong chen 1 trang trang sau hop dong.
      Nho vay hop dong ke tiep khong bi bat dau o mat sau cua hop dong truoc.
    """
    if not pdf_files:
        raise ValueError("Khong co PDF de ghep")
    if copies < 1:
        raise ValueError("So ban phai >= 1")

    writer = PdfWriter()
    total_pages = 0
    total_steps = len(pdf_files) * copies
    step = 0

    for pdf in pdf_files:
        reader = PdfReader(str(pdf))
        if reader.is_encrypted:
            raise ValueError(f"PDF bi khoa: {pdf.name}")

        page_count = len(reader.pages)
        if page_count == 0:
            raise ValueError(f"PDF khong co trang: {pdf.name}")

        for _ in range(copies):
            for page in reader.pages:
                writer.add_page(page)
                total_pages += 1

            # Duplex safety: contract/copy with odd page count gets a blank back page.
            if duplex and page_count % 2 == 1:
                last_page = reader.pages[-1]
                writer.add_blank_page(
                    width=float(last_page.mediabox.width),
                    height=float(last_page.mediabox.height),
                )
                total_pages += 1

            step += 1
            if progress:
                progress(step, total_steps, f"Dang ghep {pdf.name}")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("wb") as f:
        writer.write(f)
    return total_pages


def _pdf_page_is_trailing_blank(page) -> bool:
    """True when a trailing Word page contains no content except its page number."""
    try:
        text = str(page.extract_text() or "")
    except Exception:
        return False

    lines = [line.strip() for line in text.splitlines() if line.strip()]
    page_number_only = re.compile(
        r"^(?:(?:trang|page)\s*)?(?:[-–—]\s*)?\d+(?:\s*/\s*\d+)?(?:\s*[-–—])?$",
        re.IGNORECASE,
    )
    if any(not page_number_only.fullmatch(line) for line in lines):
        return False

    # Never remove pages containing images/forms or interactive annotations.
    try:
        resources = page.get("/Resources")
        resources = resources.get_object() if resources is not None else {}
        xobjects = resources.get("/XObject") if resources else None
        xobjects = xobjects.get_object() if xobjects is not None else {}
        if xobjects:
            return False
    except Exception:
        return False

    try:
        if page.get("/Annots"):
            return False
    except Exception:
        return False

    return True


def remove_trailing_blank_pdf_pages(pdf_path: Path) -> int:
    """Remove Word's accidental blank page(s) at the end of an exported PDF."""
    reader = PdfReader(str(pdf_path))
    keep_pages = len(reader.pages)
    while keep_pages > 1 and _pdf_page_is_trailing_blank(reader.pages[keep_pages - 1]):
        keep_pages -= 1

    removed = len(reader.pages) - keep_pages
    if removed == 0:
        return 0

    writer = PdfWriter()
    for page in reader.pages[:keep_pages]:
        writer.add_page(page)

    temp_path = pdf_path.with_name(
        f".{pdf_path.stem}.{os.getpid()}.trim.tmp.pdf"
    )
    try:
        with temp_path.open("wb") as stream:
            writer.write(stream)
        os.replace(temp_path, pdf_path)
    finally:
        if temp_path.exists():
            try:
                temp_path.unlink()
            except Exception:
                pass
    return removed


def list_printers() -> Tuple[List[str], str]:
    if os.name != "nt":
        return [], ""
    try:
        import win32print  # type: ignore
        flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
        names = sorted({p[2] for p in win32print.EnumPrinters(flags)})
        try:
            default = win32print.GetDefaultPrinter()
        except Exception:
            default = ""
        return names, default
    except Exception:
        return [], ""


def printer_blocking_status(printer: str) -> List[str]:
    """Return Windows printer conditions that should block a large print job."""
    if os.name != "nt" or not printer:
        return []
    try:
        import win32print  # type: ignore
        handle = win32print.OpenPrinter(printer)
        try:
            info = win32print.GetPrinter(handle, 2)
        finally:
            win32print.ClosePrinter(handle)

        status = int(info.get("Status", 0) or 0)
        checks = (
            (getattr(win32print, "PRINTER_STATUS_ERROR", 0x00000002), "Máy in đang báo lỗi"),
            (getattr(win32print, "PRINTER_STATUS_PAPER_JAM", 0x00000008), "Máy in đang kẹt giấy"),
            (getattr(win32print, "PRINTER_STATUS_PAPER_OUT", 0x00000010), "Máy in hết giấy"),
            (getattr(win32print, "PRINTER_STATUS_OFFLINE", 0x00000080), "Máy in đang offline"),
            (getattr(win32print, "PRINTER_STATUS_OUTPUT_BIN_FULL", 0x00000800), "Khay giấy ra đang đầy"),
            (getattr(win32print, "PRINTER_STATUS_NOT_AVAILABLE", 0x00001000), "Máy in không sẵn sàng"),
            (getattr(win32print, "PRINTER_STATUS_NO_TONER", 0x00040000), "Máy in hết mực"),
            (getattr(win32print, "PRINTER_STATUS_USER_INTERVENTION", 0x00100000), "Máy in cần người xử lý"),
            (getattr(win32print, "PRINTER_STATUS_DOOR_OPEN", 0x00400000), "Cửa máy in đang mở"),
        )
        problems = [message for flag, message in checks if flag and status & flag]
        queued_jobs = int(info.get("cJobs", 0) or 0)
        if queued_jobs >= 2:
            problems.append(
                f"Hàng đợi đang có {queued_jobs} lệnh; hãy chờ còn dưới 2 lệnh"
            )
        return problems
    except Exception as exc:
        logging.warning("Khong doc duoc trang thai may in %s: %s", printer, exc)
        return []


def find_sumatra() -> Optional[Path]:
    candidates = [
        BASE_DIR / "SumatraPDF.exe",
        Path(os.environ.get("LOCALAPPDATA", "")) / "SumatraPDF" / "SumatraPDF.exe",
        Path(os.environ.get("ProgramFiles", "C:/Program Files")) / "SumatraPDF" / "SumatraPDF.exe",
        Path(os.environ.get("ProgramFiles(x86)", "C:/Program Files (x86)")) / "SumatraPDF" / "SumatraPDF.exe",
    ]
    for p in candidates:
        if str(p) and p.exists():
            return p
    return None


def print_pdf(pdf_path: Path, printer: str, duplex: bool = True) -> None:
    if os.name != "nt":
        raise RuntimeError("Chuc nang in chi ho tro Windows")
    if not pdf_path.exists():
        raise FileNotFoundError(pdf_path)
    if not printer:
        raise ValueError("Chua chon may in")

    sumatra = find_sumatra()
    if sumatra:
        # SumatraPDF ho tro ep kho A4 va duplex truc tiep.
        # duplexlong = in 2 mat, lat canh dai (dung cho hop dong A4 doc).
        settings = "paper=A4,fit,duplexlong" if duplex else "paper=A4,fit,simplex"
        cmd = [
            str(sumatra),
            "-print-to", printer,
            "-print-settings", settings,
            "-silent",
            str(pdf_path.resolve())
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
        if result.returncode != 0:
            raise RuntimeError(result.stderr.strip() or "SumatraPDF print failed")
        return

    if duplex:
        raise RuntimeError(
            "Che do IN 2 MAT A4 can SumatraPDF de ep duplex on dinh. "
            "Hay cai SumatraPDF hoac dat SumatraPDF.exe canh Tool."
        )

    # Fallback chi cho in 1 mat.
    try:
        import win32api  # type: ignore
        rc = win32api.ShellExecute(
            0, "printto", str(pdf_path.resolve()),
            f'"{printer}"', str(pdf_path.parent), 0
        )
        if rc <= 32:
            raise RuntimeError(f"ShellExecute printto error code {rc}")
    except Exception as exc:
        raise RuntimeError(
            "Khong tim thay SumatraPDF va Windows PDF handler khong in duoc. "
            "De in silent on dinh, cai SumatraPDF hoac dat SumatraPDF.exe canh Tool.\n"
            + str(exc)
        ) from exc


def write_print_log(output_dir: Path, printer: str, contracts: int, copies: int,
                    pages: int, master_pdf: Path, status: str, error: str = "") -> None:
    log_dir = output_dir / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    path = log_dir / "print_log.csv"
    exists = path.exists()
    with path.open("a", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        if not exists:
            w.writerow(["timestamp", "printer", "contracts", "copies", "total_pages", "master_pdf", "status", "error"])
        w.writerow([
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"), printer, contracts, copies,
            pages, str(master_pdf), status, error
        ])


def load_config() -> Dict[str, Any]:
    if CONFIG_PATH.exists():
        try:
            return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def save_config(data: Dict[str, Any]) -> None:
    try:
        CONFIG_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


class GenerationWorker(QThread):
    progress = Signal(int, int, str)
    message = Signal(str)
    finished_ok = Signal(list)
    failed = Signal(str)

    def __init__(self, employees: List[Employee], template: Path, output_dir: Path):
        super().__init__()
        self.employees = employees
        self.template = template
        self.output_dir = output_dir

    def run(self):
        try:
            word_dir = self.output_dir / "WORD"
            pdf_dir = self.output_dir / "PDF"
            word_dir.mkdir(parents=True, exist_ok=True)
            pdf_dir.mkdir(parents=True, exist_ok=True)

            def cb(cur, tot, msg):
                self.progress.emit(cur, tot, msg)

            if os.name == "nt":
                pdfs = generate_contracts_word_com(
                    self.employees, self.template, word_dir, pdf_dir, progress=cb
                )
                self.finished_ok.emit([str(p) for p in pdfs])
                return

            # Non-Windows fallback is only for development/QA. Production uses Word COM.
            docx_files: List[Path] = []
            total = len(self.employees)
            for i, emp in enumerate(self.employees, start=1):
                self.progress.emit(i, total, f"Dang tao Word {i}/{total}: {emp.ho_ten}")
                out_docx = word_dir / employee_filename(emp, "docx")
                create_contract_docx(self.template, emp, out_docx)
                docx_files.append(out_docx)
            self.message.emit("Dang chay khong phai Windows: da tao DOCX, bo qua Microsoft Word PDF export.")
            self.finished_ok.emit([str(p) for p in docx_files])
        except Exception:
            logging.exception("Generation failed")
            self.failed.emit(traceback.format_exc())


class MasterWorker(QThread):
    progress = Signal(int, int, str)
    finished_ok = Signal(str, int)
    failed = Signal(str)

    def __init__(self, pdfs: List[Path], copies: int, output_path: Path, duplex: bool = True):
        super().__init__()
        self.pdfs = pdfs
        self.copies = copies
        self.output_path = output_path
        self.duplex = duplex

    def run(self):
        try:
            def cb(cur, tot, msg):
                self.progress.emit(cur, tot, msg)
            pages = create_master_pdf(
                self.pdfs, self.copies, self.output_path,
                duplex=self.duplex, progress=cb
            )
            self.finished_ok.emit(str(self.output_path), pages)
        except Exception:
            logging.exception("Master PDF failed")
            self.failed.emit(traceback.format_exc())


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("TL CONTRACT TOOL V9 - Tạo & In HĐLĐ")
        self.resize(1500, 900)
        self.setMinimumSize(1200, 760)
        self.employees: List[Employee] = []
        self.generated_pdfs: List[Path] = []
        self.generated_employees: List[Employee] = []
        self.generated_template_path: Optional[Path] = None
        self.pending_generation_employees: List[Employee] = []
        self.master_pdf: Optional[Path] = None
        self.master_employees: List[Employee] = []
        self.master_pages: int = 0
        self.master_copies: int = 0
        self.master_duplex: bool = True
        self.master_sha256: str = ""
        self.master_job_key: str = ""
        self.master_batch_label: str = ""
        self.print_locked: bool = False
        self.reprint_allowed_job_key: Optional[str] = None
        self.auto_print_after_master: bool = False
        self.busy: bool = False
        self.worker: Optional[QThread] = None
        self._rendering_table: bool = False
        # Safety flags: only allow master/print from contracts generated in this app session.
        self.generation_ready: bool = False
        self.config = load_config()
        self._build_ui()
        self._load_initial_values()
        self.refresh_printers()
        self.load_excel_data(silent=True)
        self.on_print_settings_changed(self.copies_spin.value())
        self.refresh_action_states()

    def _build_ui(self):
        root = QWidget()
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.setSpacing(4)

        title = QLabel("THIỆN LONG TOOL HDLĐ")
        title.setStyleSheet("font-size: 22px; font-weight: 700; color: #1F4E78;")
        subtitle = QLabel("Sinh Hợp đồng lao động từ Excel + Word | PDF tổng | In 2 mặt A4")
        subtitle.setStyleSheet("font-size: 12px; color: #555;")
        layout.addWidget(title)
        layout.addWidget(subtitle)

        files_box = QGroupBox("Nguồn dữ liệu")
        grid = QGridLayout(files_box)
        self.excel_edit = QLineEdit()
        self.template_edit = QLineEdit()
        self.output_edit = QLineEdit()
        self.btn_choose_excel = QPushButton("Chọn Excel")
        self.btn_choose_template = QPushButton("Chọn mẫu Word")
        self.btn_choose_output = QPushButton("Chọn OUTPUT")
        self.btn_choose_excel.clicked.connect(self.choose_excel)
        self.btn_choose_template.clicked.connect(self.choose_template)
        self.btn_choose_output.clicked.connect(self.choose_output)
        grid.addWidget(QLabel("Excel nhân sự:"), 0, 0); grid.addWidget(self.excel_edit, 0, 1); grid.addWidget(self.btn_choose_excel, 0, 2)
        grid.addWidget(QLabel("Mẫu HĐ Word:"), 1, 0); grid.addWidget(self.template_edit, 1, 1); grid.addWidget(self.btn_choose_template, 1, 2)
        grid.addWidget(QLabel("Thư mục xuất:"), 2, 0); grid.addWidget(self.output_edit, 2, 1); grid.addWidget(self.btn_choose_output, 2, 2)
        layout.addWidget(files_box)

        action_row = QHBoxLayout()
        self.btn_load = QPushButton("1. NẠP & KIỂM TRA EXCEL")
        self.btn_generate = QPushButton("2. SINH HỢP ĐỒNG")
        self.btn_open_output = QPushButton("MỞ OUTPUT")
        self.btn_clear_output = QPushButton("XÓA FILE ĐÃ TẠO")
        self.btn_load.clicked.connect(lambda: self.load_excel_data(silent=False))
        self.btn_generate.clicked.connect(self.generate_contracts)
        self.btn_open_output.clicked.connect(self.open_output)
        self.btn_clear_output.clicked.connect(self.clear_generated_files)
        for b in (self.btn_load, self.btn_generate, self.btn_open_output, self.btn_clear_output):
            b.setMinimumHeight(38)
        self.btn_generate.setStyleSheet("font-weight:700; background:#1F4E78; color:white;")
        self.btn_clear_output.setStyleSheet("font-weight:700;")
        action_row.addWidget(self.btn_load)
        action_row.addWidget(self.btn_generate)
        action_row.addWidget(self.btn_open_output)
        action_row.addWidget(self.btn_clear_output)
        action_row.addStretch()
        layout.addLayout(action_row)

        selection_row = QHBoxLayout()
        self.summary = QLabel("Chưa nạp dữ liệu")
        self.summary.setStyleSheet("font-weight: 600;")
        self.btn_select_all = QPushButton("TICK TẤT CẢ (≤25)")
        self.btn_select_none = QPushButton("BỎ TICK")
        self.btn_select_all.clicked.connect(lambda: self.set_all_employee_selection(True))
        self.btn_select_none.clicked.connect(lambda: self.set_all_employee_selection(False))
        selection_row.addWidget(self.summary)
        selection_row.addStretch()
        selection_row.addWidget(self.btn_select_all)
        selection_row.addWidget(self.btn_select_none)
        layout.addLayout(selection_row)

        batch_row = QHBoxLayout()
        self.batch_size_spin = QSpinBox()
        self.batch_size_spin.setRange(5, MAX_BATCH_SIZE)
        self.batch_size_spin.setValue(DEFAULT_BATCH_SIZE)
        self.batch_combo = QComboBox()
        self.batch_combo.setMinimumWidth(330)
        self.btn_prev_batch = QPushButton("LÔ TRƯỚC")
        self.btn_apply_batch = QPushButton("ÁP DỤNG LÔ")
        self.btn_next_batch = QPushButton("LÔ TIẾP THEO")
        self.batch_status = QLabel("Chưa có dữ liệu lô")
        self.batch_status.setStyleSheet("font-weight:600; color:#1F4E78;")
        self.batch_size_spin.valueChanged.connect(self.on_batch_size_changed)
        self.btn_prev_batch.clicked.connect(lambda: self.move_batch(-1))
        self.btn_apply_batch.clicked.connect(self.apply_current_batch)
        self.btn_next_batch.clicked.connect(lambda: self.move_batch(1))
        batch_row.addWidget(QLabel("Nhân viên/lô:"))
        batch_row.addWidget(self.batch_size_spin)
        batch_row.addWidget(QLabel("Chọn lô:"))
        batch_row.addWidget(self.batch_combo, 1)
        batch_row.addWidget(self.btn_prev_batch)
        batch_row.addWidget(self.btn_apply_batch)
        batch_row.addWidget(self.btn_next_batch)
        batch_row.addWidget(self.batch_status)
        layout.addLayout(batch_row)

        headers = ["Chọn in"] + [label for _, label in TABLE_COLUMNS] + ["Trạng thái"]
        self.table = QTableWidget(0, len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.setWordWrap(False)
        self.table.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerItem)
        self.table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.table.itemChanged.connect(self.on_table_item_changed)
        hdr = self.table.horizontalHeader()
        hdr.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        hdr.setMinimumSectionSize(55)
        row_hdr = self.table.verticalHeader()
        row_hdr.setSectionResizeMode(QHeaderView.ResizeMode.Fixed)
        row_hdr.setDefaultSectionSize(TABLE_ROW_HEIGHT)
        row_hdr.setMinimumSectionSize(TABLE_ROW_HEIGHT)
        table_height = (
            hdr.sizeHint().height()
            + TABLE_VISIBLE_ROWS * TABLE_ROW_HEIGHT
            + self.table.horizontalScrollBar().sizeHint().height()
            + self.table.frameWidth() * 2
        )
        self.table.setFixedHeight(table_height)
        widths = [75, 55, 90, 190, 95, 85, 85, 260, 120, 120, 170, 120, 120, 170, 210, 125, 95, 95, 145, 190, 150]
        for column, width in enumerate(widths):
            self.table.setColumnWidth(column, width)
        layout.addWidget(self.table)

        print_box = QGroupBox("Tạo PDF tổng & In")
        pgrid = QGridLayout(print_box)
        self.printer_combo = QComboBox()
        self.copies_spin = QSpinBox(); self.copies_spin.setRange(1, 10); self.copies_spin.setValue(1)
        self.print_mode_combo = QComboBox()
        self.print_mode_combo.addItems([
            "2 mặt A4 - lật cạnh dài",
            "1 mặt A4"
        ])
        self.print_mode_combo.setCurrentIndex(0)
        self.btn_refresh_printer = QPushButton("Refresh máy in")
        self.btn_open_queue = QPushButton("MỞ HÀNG ĐỢI IN")
        self.btn_master = QPushButton("3. TẠO PDF TỔNG")
        self.btn_open_master = QPushButton("XEM PDF TỔNG")
        self.btn_print = QPushButton("4. IN NHÂN VIÊN ĐÃ TICK")
        self.btn_allow_reprint = QPushButton("CHO PHÉP IN LẠI 1 LẦN")
        self.btn_refresh_printer.clicked.connect(self.refresh_printers)
        self.btn_open_queue.clicked.connect(self.open_print_queue)
        self.btn_master.clicked.connect(self.create_master)
        self.btn_open_master.clicked.connect(self.open_master)
        self.btn_print.clicked.connect(self.print_master)
        self.btn_allow_reprint.clicked.connect(self.allow_reprint_once)
        self.btn_print.setStyleSheet("font-weight:700; background:#C00000; color:white;")
        self.btn_allow_reprint.setStyleSheet("font-weight:700; background:#F4B183;")
        self.copies_spin.valueChanged.connect(self.on_print_settings_changed)
        self.print_mode_combo.currentIndexChanged.connect(self.on_print_settings_changed)
        pgrid.addWidget(QLabel("Máy in:"), 0, 0); pgrid.addWidget(self.printer_combo, 0, 1); pgrid.addWidget(self.btn_refresh_printer, 0, 2)
        pgrid.addWidget(QLabel("Số bản / nhân viên:"), 1, 0); pgrid.addWidget(self.copies_spin, 1, 1); pgrid.addWidget(self.btn_open_queue, 1, 2)
        pgrid.addWidget(QLabel("Chế độ in:"), 2, 0); pgrid.addWidget(self.print_mode_combo, 2, 1, 1, 2)
        pgrid.addWidget(self.btn_master, 3, 0); pgrid.addWidget(self.btn_open_master, 3, 1); pgrid.addWidget(self.btn_print, 3, 2)
        self.print_safety_label = QLabel(
            f"Giới hạn an toàn: tối đa {MAX_BATCH_SIZE} nhân viên và "
            f"{MAX_PRINT_JOB_PAGES} trang cho mỗi lệnh in."
        )
        self.print_safety_label.setStyleSheet("color:#9C0006; font-weight:600;")
        pgrid.addWidget(self.print_safety_label, 4, 0, 1, 2)
        pgrid.addWidget(self.btn_allow_reprint, 4, 2)
        layout.addWidget(print_box)

        self.progress = QProgressBar(); self.progress.setRange(0, 100); self.progress.setValue(0)
        self.status = QLabel("Sẵn sàng")
        layout.addWidget(self.progress)
        layout.addWidget(self.status)

        self.log_box = QPlainTextEdit(); self.log_box.setReadOnly(True); self.log_box.setFixedHeight(55)
        layout.addWidget(self.log_box)

        # Keep the form itself fixed. Only the employee table scrolls. The
        # central widget's minimum height prevents Qt from overlapping panels.
        layout.activate()
        root.setMinimumHeight(layout.minimumSize().height())

    def _load_initial_values(self):
        configured_excel = Path(self.config.get("excel", str(DEFAULT_EXCEL)))
        configured_template = Path(self.config.get("template", str(DEFAULT_TEMPLATE)))

        # Cấu hình của các bản trước vẫn trỏ tới file cũ. Khi bộ file cập nhật
        # đã có sẵn, tự chuyển sang bản mới để người dùng không phải chọn lại.
        if configured_excel.name.casefold() == LEGACY_EXCEL_NAME.casefold() and DEFAULT_EXCEL.exists():
            configured_excel = DEFAULT_EXCEL
        if configured_template.name.casefold() in {
            LEGACY_TEMPLATE_NAME.casefold(), UPDATED_TEMPLATE_NAME.casefold()
        } and DEFAULT_TEMPLATE.exists():
            configured_template = DEFAULT_TEMPLATE

        self.excel_edit.setText(str(configured_excel))
        self.template_edit.setText(str(configured_template))
        self.output_edit.setText(self.config.get("output", str(DEFAULT_OUTPUT)))
        configured_batch = int(self.config.get("batch_size", DEFAULT_BATCH_SIZE) or DEFAULT_BATCH_SIZE)
        self.batch_size_spin.setValue(max(5, min(MAX_BATCH_SIZE, configured_batch)))
        self.copies_spin.setValue(1)  # Safety default: never start with 2 copies by accident
        self.print_mode_combo.setCurrentIndex(0 if self.config.get("duplex", True) else 1)

    def log(self, text: str):
        self.log_box.appendPlainText(f"[{datetime.now().strftime('%H:%M:%S')}] {text}")

    def choose_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "Chọn Excel", self.excel_edit.text(), "Excel (*.xlsx)")
        if path:
            self.excel_edit.setText(path)
            self.load_excel_data(silent=False)

    def choose_template(self):
        path, _ = QFileDialog.getOpenFileName(self, "Chọn mẫu Word", self.template_edit.text(), "Word (*.docx)")
        if path:
            changed = path != self.template_edit.text().strip()
            self.template_edit.setText(path)
            if changed:
                self.invalidate_generated_batch("Đã đổi mẫu Word; cần sinh lại hợp đồng.")

    def choose_output(self):
        path = QFileDialog.getExistingDirectory(self, "Chọn thư mục OUTPUT", self.output_edit.text())
        if path:
            changed = path != self.output_edit.text().strip()
            self.output_edit.setText(path)
            if changed:
                self.invalidate_generated_batch("Đã đổi thư mục OUTPUT; cần sinh lại hợp đồng.")

    def output_dir(self) -> Path:
        return Path(self.output_edit.text().strip())

    def persist(self):
        save_config({
            "excel": self.excel_edit.text().strip(),
            "template": self.template_edit.text().strip(),
            "output": self.output_edit.text().strip(),
            "copies": self.copies_spin.value(),
            "printer": self.printer_combo.currentText(),
            "duplex": self.print_mode_combo.currentIndex() == 0,
            "batch_size": self.batch_size_spin.value(),
        })

    def load_excel_data(self, silent=False):
        try:
            # Any Excel reload invalidates the previous generated batch.
            # This prevents printing stale PDFs from an older employee list.
            self.generation_ready = False
            self.generated_pdfs = []
            self.generated_employees = []
            self.generated_template_path = None
            self.pending_generation_employees = []
            self.clear_master_state()
            excel = Path(self.excel_edit.text().strip())
            self.employees = read_employees(excel)
            self.rebuild_batch_selector()
            if len(self.selected_employees()) > self.batch_size_spin.value():
                batch_size = self.batch_size_spin.value()
                for index, emp in enumerate(self.employees):
                    emp.selected = index < batch_size
                self.batch_combo.setCurrentIndex(0)
                self.log(
                    f"Excel tick quá {batch_size} nhân viên; Tool tự chọn lô đầu tiên để bảo vệ hệ thống."
                )
            self.restore_existing_generated_pdfs(excel)
            self.render_table()
            self.update_selection_summary()
            self.update_batch_status()
            if not silent:
                self.log(
                    f"Nạp Excel thành công: {len(self.employees)} dòng có dữ liệu, "
                    f"{len(self.selected_employees())} nhân viên được tick."
                )
            self.persist()
            self.refresh_action_states()
        except Exception as exc:
            self.employees = []
            self.table.setRowCount(0)
            self.rebuild_batch_selector()
            self.summary.setText("Lỗi nạp Excel")
            self.update_batch_status()
            if not silent:
                QMessageBox.critical(self, "Lỗi Excel", str(exc))
            self.log(str(exc))
            self.refresh_action_states()

    def render_table(self):
        self._rendering_table = True
        self.table.blockSignals(True)
        try:
            self.table.setRowCount(len(self.employees))
            for r, emp in enumerate(self.employees):
                check_item = QTableWidgetItem()
                check_item.setFlags(
                    Qt.ItemFlag.ItemIsEnabled
                    | Qt.ItemFlag.ItemIsSelectable
                    | Qt.ItemFlag.ItemIsUserCheckable
                )
                check_item.setCheckState(
                    Qt.CheckState.Checked if emp.selected else Qt.CheckState.Unchecked
                )
                check_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                check_item.setToolTip(f"Dòng Excel {emp.row_number}: tick để sinh và in hợp đồng")
                self.table.setItem(r, 0, check_item)
                for c, (field, _) in enumerate(TABLE_COLUMNS, start=1):
                    value = emp.data.get(field)
                    if field in DATE_FIELDS:
                        display = format_date(value)
                    elif field == "MUC_LUONG_CHINH":
                        display = format_money(value)
                    else:
                        display = text_value(value)
                    item = QTableWidgetItem(display)
                    item.setToolTip(display)
                    self.table.setItem(r, c, item)

                status = QTableWidgetItem("OK" if emp.valid else emp.error)
                status.setToolTip(status.text())
                self.table.setItem(r, len(TABLE_COLUMNS) + 1, status)
        finally:
            self.table.blockSignals(False)
            self._rendering_table = False

    def selected_employees(self) -> List[Employee]:
        return [emp for emp in self.employees if emp.selected]

    @staticmethod
    def employee_key(emp: Employee) -> Tuple[int, str]:
        return emp.stt, emp.ma_nv.casefold()

    def generated_pdf_map(self) -> Dict[Tuple[int, str], Path]:
        return {
            self.employee_key(emp): pdf
            for emp, pdf in zip(self.generated_employees, self.generated_pdfs)
            if pdf.exists()
        }

    def pdfs_for_employees(self, employees: List[Employee]) -> List[Path]:
        pdf_map = self.generated_pdf_map()
        return [
            pdf_map[self.employee_key(emp)]
            for emp in employees
            if self.employee_key(emp) in pdf_map
        ]

    def selection_has_generated_pdfs(self) -> bool:
        selected = self.selected_employees()
        if not selected or any(not emp.valid for emp in selected):
            return False
        return len(self.pdfs_for_employees(selected)) == len(selected)

    def restore_existing_generated_pdfs(self, excel_path: Path) -> None:
        """Restore fresh Tool outputs after restart so tick selection remains printable."""
        template = Path(self.template_edit.text().strip())
        if not template.exists() or not excel_path.exists():
            return
        try:
            freshness_cutoff = max(excel_path.stat().st_mtime, template.stat().st_mtime)
        except OSError:
            return

        pdf_dir = self.output_dir() / "PDF"
        word_dir = self.output_dir() / "WORD"
        restored_employees: List[Employee] = []
        restored_pdfs: List[Path] = []
        for emp in self.employees:
            if not emp.valid:
                continue
            pdf = pdf_dir / employee_filename(emp, "pdf")
            docx = word_dir / employee_filename(emp, "docx")
            try:
                if not pdf.is_file() or not docx.is_file():
                    continue
                if min(pdf.stat().st_mtime, docx.stat().st_mtime) < freshness_cutoff:
                    continue
                reader = PdfReader(str(pdf))
                if reader.is_encrypted or len(reader.pages) == 0:
                    continue
            except Exception:
                continue
            restored_employees.append(emp)
            restored_pdfs.append(pdf)

        if restored_pdfs:
            self.generated_employees = restored_employees
            self.generated_pdfs = restored_pdfs
            self.generated_template_path = template
            self.generation_ready = True
            self.log(
                f"Đã khôi phục {len(restored_pdfs)} PDF nhân viên còn mới trong OUTPUT; "
                "có thể tick để tạo lô in."
            )

    def rebuild_batch_selector(self):
        current = max(self.batch_combo.currentIndex(), 0)
        batch_size = self.batch_size_spin.value()
        total = len(self.employees)
        batch_count = (total + batch_size - 1) // batch_size if total else 0
        self.batch_combo.blockSignals(True)
        try:
            self.batch_combo.clear()
            for batch_index in range(batch_count):
                start = batch_index * batch_size
                end = min(start + batch_size, total)
                first_stt = self.employees[start].stt
                last_stt = self.employees[end - 1].stt
                self.batch_combo.addItem(
                    f"Lô {batch_index + 1}/{batch_count} | dòng {start + 1}-{end} "
                    f"| STT {first_stt:03d}-{last_stt:03d}",
                    (start, end),
                )
            if batch_count:
                self.batch_combo.setCurrentIndex(min(current, batch_count - 1))
        finally:
            self.batch_combo.blockSignals(False)

    def on_batch_size_changed(self, _value: int):
        self.rebuild_batch_selector()
        if self.employees:
            self.apply_current_batch()
        self.persist()

    def apply_current_batch(self):
        index = self.batch_combo.currentIndex()
        data = self.batch_combo.itemData(index) if index >= 0 else None
        if not data:
            return
        start, end = data
        changed = False
        for row, emp in enumerate(self.employees):
            selected = start <= row < end
            if emp.selected != selected:
                emp.selected = selected
                changed = True
        self.render_table()
        if changed:
            self.invalidate_master_for_selection("Đã chuyển sang lô nhân viên khác.")
        self.update_selection_summary()
        self.update_batch_status()
        if self.employees[start:end]:
            self.table.scrollToItem(self.table.item(start, 0))
        self.log(f"Đã chọn {self.current_batch_label()} ({end - start} nhân viên).")

    def move_batch(self, delta: int):
        if self.batch_combo.count() == 0:
            return
        target = max(0, min(self.batch_combo.count() - 1, self.batch_combo.currentIndex() + delta))
        self.batch_combo.setCurrentIndex(target)
        self.apply_current_batch()

    def current_batch_label(self) -> str:
        selected_rows = [index for index, emp in enumerate(self.employees) if emp.selected]
        if not selected_rows:
            return "Chưa chọn lô"
        for index in range(self.batch_combo.count()):
            data = self.batch_combo.itemData(index)
            if not data:
                continue
            start, end = data
            if selected_rows == list(range(start, end)):
                return f"Lô {index + 1}/{self.batch_combo.count()}"
        return f"Lô tùy chọn ({len(selected_rows)} NV)"

    def update_batch_status(self):
        label = self.current_batch_label()
        selected = self.selected_employees()
        if selected:
            self.batch_status.setText(
                f"{label}: {selected[0].ma_nv} → {selected[-1].ma_nv}"
            )
        else:
            self.batch_status.setText("Chưa chọn nhân viên")

    def update_selection_summary(self):
        selected = self.selected_employees()
        valid = sum(1 for emp in selected if emp.valid)
        invalid = len(selected) - valid
        self.summary.setText(
            f"Tổng Excel: {len(self.employees)} | Đã tick: {len(selected)} nhân viên "
            f"| Hợp lệ: {valid} | Lỗi: {invalid}"
        )

    def invalidate_generated_batch(self, reason: str = ""):
        had_generated_data = self.generation_ready or self.master_pdf is not None
        self.generated_pdfs = []
        self.generated_employees = []
        self.generated_template_path = None
        self.pending_generation_employees = []
        self.generation_ready = False
        self.clear_master_state()
        if had_generated_data:
            self.status.setText("Danh sách tick đã đổi - cần sinh hợp đồng mới")
            if reason:
                self.log(reason)
        self.refresh_action_states()

    def invalidate_master_for_selection(self, reason: str = ""):
        """Keep generated employee PDFs; only discard the master for the old ticks."""
        had_master = self.master_pdf is not None
        self.clear_master_state()
        if had_master:
            self.status.setText("Danh sách tick đã đổi - cần tạo lại PDF tổng")
        if reason:
            self.log(reason)
        self.refresh_action_states()

    def clear_master_state(self):
        self.master_pdf = None
        self.master_employees = []
        self.master_pages = 0
        self.master_copies = 0
        self.master_duplex = True
        self.master_sha256 = ""
        self.master_job_key = ""
        self.master_batch_label = ""
        self.print_locked = False
        self.reprint_allowed_job_key = None
        self.auto_print_after_master = False

    def on_print_settings_changed(self, _value: int):
        had_master = self.master_pdf is not None
        if had_master:
            self.clear_master_state()
            self.status.setText("Đã đổi số bản/chế độ in - cần tạo lại PDF tổng")
            self.log("Thiết lập in đã thay đổi; PDF tổng cũ bị khóa để tránh in sai số bản.")
        copies = max(self.copies_spin.value(), 1)
        recommended = max(1, min(MAX_BATCH_SIZE, MAX_PRINT_JOB_PAGES // (4 * copies)))
        self.print_safety_label.setText(
            f"Giới hạn {MAX_PRINT_JOB_PAGES} trang/lệnh. Với {copies} bản/người: "
            f"khuyến nghị tối đa {recommended} nhân viên/lô."
        )
        self.refresh_action_states()

    def on_table_item_changed(self, item: QTableWidgetItem):
        if self._rendering_table or item.column() != 0:
            return
        row = item.row()
        if not 0 <= row < len(self.employees):
            return
        selected = item.checkState() == Qt.CheckState.Checked
        if self.employees[row].selected == selected:
            return
        self.employees[row].selected = selected
        self.invalidate_master_for_selection("Đã thay đổi danh sách nhân viên được tick.")
        self.update_selection_summary()
        self.update_batch_status()

    def set_all_employee_selection(self, selected: bool):
        if selected and len(self.employees) > MAX_BATCH_SIZE:
            QMessageBox.information(
                self, "Giới hạn lô an toàn",
                f"Danh sách có {len(self.employees)} nhân viên. Tool chỉ cho chọn tối đa "
                f"{MAX_BATCH_SIZE} người/lô.\n\nĐã áp dụng lô đang chọn thay vì tick toàn bộ."
            )
            self.apply_current_batch()
            return
        changed = any(emp.selected != selected for emp in self.employees)
        if not changed:
            return
        for emp in self.employees:
            emp.selected = selected
        self.render_table()
        self.invalidate_master_for_selection("Đã thay đổi toàn bộ danh sách nhân viên được tick.")
        self.update_selection_summary()
        self.update_batch_status()

    def refresh_action_states(self):
        ready_master = bool(self.master_pdf and self.master_pdf.exists())
        selection_ready = self.selection_has_generated_pdfs()
        common_controls = (
            self.btn_load, self.btn_generate, self.btn_clear_output,
            self.btn_select_all, self.btn_select_none, self.btn_prev_batch,
            self.btn_apply_batch, self.btn_next_batch, self.batch_size_spin,
            self.batch_combo, self.copies_spin, self.print_mode_combo,
            self.printer_combo, self.btn_refresh_printer, self.btn_open_queue,
            self.btn_choose_excel, self.btn_choose_template, self.btn_choose_output,
        )
        for control in common_controls:
            control.setEnabled(not self.busy)
        self.btn_master.setEnabled(not self.busy and selection_ready)
        self.btn_open_master.setEnabled(not self.busy and ready_master)
        # Keep IN clickable so the user receives an exact missing-step message.
        # When employee PDFs are ready, clicking IN automatically creates master.
        self.btn_print.setEnabled(not self.busy and not self.print_locked)
        self.btn_allow_reprint.setEnabled(not self.busy and ready_master and self.print_locked)
        if ready_master:
            self.btn_print.setToolTip("In đúng PDF tổng của các nhân viên hiện đang được tick.")
        elif selection_ready:
            self.btn_print.setToolTip("Tool sẽ tự tạo PDF tổng từ các nhân viên đã tick rồi xác nhận in.")
        else:
            self.btn_print.setToolTip("Tick nhân viên và sinh hợp đồng trước; bấm để xem hướng dẫn cụ thể.")
        self.table.setEnabled(not self.busy)
        for edit in (self.excel_edit, self.template_edit, self.output_edit):
            edit.setReadOnly(self.busy)

    def set_busy(self, busy: bool):
        self.busy = busy
        self.refresh_action_states()

    def on_progress(self, cur: int, total: int, msg: str):
        pct = int(cur * 100 / max(total, 1))
        self.progress.setValue(pct)
        self.status.setText(msg)

    def generate_contracts(self):
        if not self.employees:
            self.load_excel_data(silent=True)
        if not self.employees:
            QMessageBox.warning(
                self, "Chưa có dữ liệu",
                "Excel không có dòng nhân viên nào. Hãy kiểm tra lại sheet NHAN_SU_HDLD."
            )
            return
        selected = self.selected_employees()
        if not selected:
            QMessageBox.warning(
                self, "Chưa tick nhân viên",
                "Hãy tick ít nhất một nhân viên trong cột 'Chọn in'."
            )
            return
        bad = [e for e in selected if not e.valid]
        if bad:
            names = "\n".join(f"- {e.ma_nv} {e.ho_ten}: {e.error}" for e in bad[:15])
            QMessageBox.warning(
                self, "Dữ liệu được tick chưa hợp lệ",
                "Hãy sửa Excel hoặc bỏ tick dòng lỗi trước khi sinh HĐ:\n" + names
            )
            return
        try:
            validate_employee_batch(selected, MAX_BATCH_SIZE)
        except ValueError as exc:
            self.log(str(exc))
            QMessageBox.critical(self, "Lô dữ liệu không an toàn", str(exc))
            return
        template = Path(self.template_edit.text().strip())
        if not template.exists():
            QMessageBox.critical(self, "Lỗi", "Không tìm thấy mẫu Word.")
            return
        out = self.output_dir()
        out.mkdir(parents=True, exist_ok=True)
        configure_logging(out)
        try:
            validate_output_targets(selected, out / "WORD", out / "PDF")
        except FileExistsError as exc:
            self.log(str(exc))
            QMessageBox.critical(self, "Trùng file nhân viên", str(exc))
            return
        self.pending_generation_employees = list(selected)
        self.generated_template_path = template
        self.persist(); self.set_busy(True); self.progress.setValue(0)
        self.worker = GenerationWorker(selected, template, out)
        self.worker.progress.connect(self.on_progress)
        self.worker.message.connect(self.log)
        self.worker.finished_ok.connect(self.generation_done)
        self.worker.failed.connect(self.generation_failed)
        self.worker.start()

    def generation_done(self, files: List[str]):
        self.set_busy(False); self.progress.setValue(100); self.status.setText("Sinh hợp đồng hoàn tất")
        paths = [Path(p) for p in files]
        new_pdfs = [p for p in paths if p.suffix.lower() == ".pdf"]
        new_employees = list(self.pending_generation_employees)
        self.pending_generation_employees = []
        if len(new_pdfs) == len(new_employees):
            inventory: Dict[Tuple[int, str], Tuple[Employee, Path]] = {
                self.employee_key(emp): (emp, pdf)
                for emp, pdf in zip(self.generated_employees, self.generated_pdfs)
                if pdf.exists()
            }
            for emp, pdf in zip(new_employees, new_pdfs):
                inventory[self.employee_key(emp)] = (emp, pdf)
            ordered = [
                inventory[self.employee_key(emp)]
                for emp in self.employees
                if self.employee_key(emp) in inventory
            ]
            self.generated_employees = [emp for emp, _ in ordered]
            self.generated_pdfs = [pdf for _, pdf in ordered]
        else:
            self.generated_employees = []
            self.generated_pdfs = []
        self.generation_ready = bool(self.generated_pdfs)
        self.clear_master_state()
        self.log(
            f"Đã sinh {len(paths)} file. PDF mới: {len(new_pdfs)} | "
            f"PDF sẵn sàng: {len(self.generated_pdfs)}"
        )
        self.refresh_action_states()
        if new_pdfs:
            QMessageBox.information(self, "Hoàn tất", f"Đã sinh {len(new_pdfs)} hợp đồng PDF.\nThư mục: {self.output_dir() / 'PDF'}")
        else:
            QMessageBox.information(self, "Hoàn tất DOCX", "Đã tạo Word. PDF sẽ được xuất khi chạy Tool trên Windows có Microsoft Word Desktop.")

    def generation_failed(self, detail: str):
        self.pending_generation_employees = []
        self.generated_pdfs = []
        self.generated_employees = []
        self.generated_template_path = None
        self.generation_ready = False
        self.clear_master_state()
        self.set_busy(False); self.progress.setValue(0); self.status.setText("Có lỗi")
        self.log(detail)
        last = detail.strip().splitlines()[-1] if detail.strip() else "Lỗi không xác định"
        QMessageBox.critical(self, "Lỗi", last)

    def master_failed(self, detail: str):
        self.clear_master_state()
        self.set_busy(False); self.progress.setValue(0); self.status.setText("Lỗi tạo PDF tổng")
        self.log(detail)
        last = detail.strip().splitlines()[-1] if detail.strip() else "Lỗi không xác định"
        QMessageBox.critical(
            self, "Lỗi tạo PDF tổng",
            last + "\n\nCác PDF nhân viên vẫn được giữ; bạn có thể tạo PDF tổng lại."
        )

    def current_pdfs(self) -> List[Path]:
        # Return PDFs in the same order as the employees currently ticked.
        if not self.generation_ready:
            return []
        return self.pdfs_for_employees(self.selected_employees())

    def create_master(self, auto_print: bool = False):
        selected = self.selected_employees()
        if not selected:
            QMessageBox.warning(
                self, "Chưa tick nhân viên",
                "Hãy tick ít nhất một nhân viên trong cột 'Chọn in'."
            )
            return
        bad = [emp for emp in selected if not emp.valid]
        if bad:
            names = "\n".join(f"- {emp.ma_nv} {emp.ho_ten}: {emp.error}" for emp in bad[:15])
            QMessageBox.warning(
                self, "Dữ liệu được tick chưa hợp lệ",
                "Không thể tạo lô in từ các dòng lỗi:\n" + names
            )
            return
        if not self.generation_ready:
            QMessageBox.warning(
                self, "Chưa sinh dữ liệu mới",
                "Các nhân viên đã tick chưa có PDF hợp đồng.\n\n"
                "Hãy bấm '2. SINH HỢP ĐỒNG' trước."
            )
            return
        pdfs = self.current_pdfs()
        if len(pdfs) != len(selected):
            pdf_map = self.generated_pdf_map()
            missing = [
                f"- {emp.ma_nv} {emp.ho_ten}"
                for emp in selected
                if self.employee_key(emp) not in pdf_map
            ]
            QMessageBox.warning(
                self, "Nhân viên chưa có PDF",
                "Chưa thể in vì các nhân viên sau chưa được sinh hợp đồng:\n"
                + "\n".join(missing[:15])
                + "\n\nHãy giữ nguyên tick và bấm '2. SINH HỢP ĐỒNG'."
            )
            return
        try:
            validate_employee_batch(selected, MAX_BATCH_SIZE)
        except ValueError as exc:
            QMessageBox.critical(self, "Lô dữ liệu không an toàn", str(exc))
            return
        copies = self.copies_spin.value()
        duplex = self.print_mode_combo.currentIndex() == 0
        try:
            projected_pages = calculate_master_page_count(pdfs, copies, duplex)
        except Exception as exc:
            QMessageBox.critical(self, "PDF không hợp lệ", str(exc))
            return
        if projected_pages > MAX_PRINT_JOB_PAGES:
            pages_per_employee = projected_pages / max(len(pdfs), 1)
            recommended = max(1, int(MAX_PRINT_JOB_PAGES // max(pages_per_employee, 1)))
            QMessageBox.critical(
                self, "Lô in quá lớn - đã dừng",
                f"Lô này sẽ tạo {projected_pages} trang, vượt giới hạn an toàn "
                f"{MAX_PRINT_JOB_PAGES} trang/lệnh.\n\n"
                f"Với thiết lập hiện tại, hãy chọn tối đa {recommended} nhân viên/lô."
            )
            return
        if not self.generated_template_path or not self.generated_template_path.exists():
            QMessageBox.critical(self, "Mất mẫu Word", "Không xác định được mẫu Word đã dùng để sinh lô này.")
            return
        current_template = Path(self.template_edit.text().strip())
        if current_template.resolve() != self.generated_template_path.resolve():
            QMessageBox.critical(
                self, "Mẫu Word đã thay đổi",
                "Mẫu Word hiện tại khác mẫu đã dùng để sinh PDF. Hãy sinh lại lô này."
            )
            return
        expected_pdf_dir = (self.output_dir() / "PDF").resolve()
        if any(pdf.parent.resolve() != expected_pdf_dir for pdf in pdfs):
            QMessageBox.critical(
                self, "OUTPUT đã thay đổi",
                "Thư mục OUTPUT hiện tại khác nơi đã sinh PDF. Tool dừng để tránh ghép nhầm file."
            )
            return
        try:
            self.master_job_key = build_print_job_key(
                selected, self.generated_template_path, copies, duplex
            )
        except Exception as exc:
            QMessageBox.critical(self, "Không tạo được khóa in", str(exc))
            return
        self.master_copies = copies
        self.master_duplex = duplex
        self.master_employees = list(selected)
        self.master_batch_label = self.current_batch_label()
        self.master_sha256 = ""
        self.print_locked = False
        self.reprint_allowed_job_key = None
        master_dir = self.output_dir() / "MASTER"
        batch_no = self.batch_combo.currentIndex() + 1
        batch_tag = f"LO_{batch_no:02d}" if "tùy chọn" not in self.master_batch_label.lower() else "LO_TUY_CHON"
        master = master_dir / (
            f"PRINT_{batch_tag}_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}.pdf"
        )
        self.auto_print_after_master = auto_print
        self.set_busy(True); self.progress.setValue(0)
        self.worker = MasterWorker(pdfs, copies, master, duplex=duplex)
        self.worker.progress.connect(self.on_progress)
        self.worker.finished_ok.connect(self.master_done)
        self.worker.failed.connect(self.master_failed)
        self.worker.start()

    def master_done(self, path: str, pages: int):
        self.master_pdf = Path(path)
        self.master_pages = pages
        try:
            actual_pages = len(PdfReader(str(self.master_pdf)).pages)
            if actual_pages != pages:
                raise ValueError(f"Khai báo {pages} trang nhưng file thực tế có {actual_pages} trang")
            if actual_pages > MAX_PRINT_JOB_PAGES:
                raise ValueError(
                    f"PDF tổng có {actual_pages} trang, vượt giới hạn {MAX_PRINT_JOB_PAGES} trang"
                )
            self.master_sha256 = sha256_file(self.master_pdf)
            previous = get_print_guard_record(self.output_dir(), self.master_job_key)
            self.print_locked = bool(
                previous and previous.get("status") in {"SENDING", "SENT", "UNKNOWN"}
            )
        except Exception as exc:
            self.clear_master_state()
            self.set_busy(False)
            self.progress.setValue(0)
            self.status.setText("PDF tổng không an toàn")
            QMessageBox.critical(self, "Không thể mở khóa in", str(exc))
            return

        self.set_busy(False); self.progress.setValue(100)
        self.status.setText(
            "PDF tổng đã sẵn sàng" if not self.print_locked
            else "Lô này đã từng gửi in - đang khóa chống in trùng"
        )
        self.log(f"Master PDF: {path} | {pages} trang")
        auto_print = self.auto_print_after_master
        self.auto_print_after_master = False
        lock_note = (
            "\n\nCẢNH BÁO: Lô dữ liệu này đã từng được gửi in. "
            "Tool đang khóa nút IN để tránh in trùng."
            if self.print_locked else ""
        )
        if auto_print and not self.print_locked:
            self.log("PDF tổng đã tạo tự động; chuyển sang bước xác nhận in.")
            self.print_master()
        else:
            QMessageBox.information(
                self, "PDF tổng đã tạo",
                f"{self.master_batch_label}\n"
                f"{len(self.master_employees)} nhân viên\n"
                f"{self.master_copies} bản/người\n"
                f"Chế độ: {'2 mặt A4' if self.master_duplex else '1 mặt A4'}\n"
                f"Tổng {pages} trang"
                + (f" = {(pages + 1) // 2} tờ A4\n\n" if self.master_duplex else "\n\n")
                + f"{path}"
                + lock_note
            )

    def open_master(self):
        if not self.master_pdf or not self.master_pdf.exists():
            QMessageBox.warning(self, "Chưa có", "Hãy tạo PDF tổng trước.")
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.master_pdf.resolve())))

    def clear_generated_files(self):
        out = self.output_dir()
        confirm = QMessageBox.question(
            self, "XÓA FILE ĐÃ TẠO",
            "Xóa toàn bộ file hợp đồng cũ trong WORD, PDF và MASTER?\n\n"
            "File log vẫn được giữ lại để đối chiếu.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if confirm != QMessageBox.StandardButton.Yes:
            return
        try:
            deleted = 0
            for folder_name in ("WORD", "PDF", "MASTER"):
                folder = out / folder_name
                if folder.exists():
                    for item in folder.iterdir():
                        if item.is_file() or item.is_symlink():
                            item.unlink()
                            deleted += 1
                        elif item.is_dir():
                            shutil.rmtree(item)
                            deleted += 1
                else:
                    folder.mkdir(parents=True, exist_ok=True)

            self.generated_pdfs = []
            self.generated_employees = []
            self.generated_template_path = None
            self.pending_generation_employees = []
            self.generation_ready = False
            self.clear_master_state()
            self.progress.setValue(0)
            self.status.setText("Đã xóa file cũ - cần sinh hợp đồng mới trước khi in")
            self.log(f"Đã xóa dữ liệu sinh cũ: {deleted} mục trong WORD/PDF/MASTER.")
            QMessageBox.information(
                self, "Đã xóa",
                "Đã xóa sạch file hợp đồng cũ.\n\nTool đã khóa chức năng in cho đến khi bạn sinh hợp đồng mới từ Excel.\n"
                "Lịch sử chống in trùng vẫn được giữ lại."
            )
            self.refresh_action_states()
        except Exception as exc:
            self.log(str(exc))
            QMessageBox.critical(self, "Lỗi xóa file", str(exc))

    def open_output(self):
        out = self.output_dir(); out.mkdir(parents=True, exist_ok=True)
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(out.resolve())))

    def refresh_printers(self):
        current = self.config.get("printer", "")
        names, default = list_printers()
        self.printer_combo.clear(); self.printer_combo.addItems(names)
        target = current if current in names else default
        if target in names:
            self.printer_combo.setCurrentText(target)
        if not names and os.name == "nt":
            self.log("Không đọc được danh sách máy in Windows.")

    def open_print_queue(self):
        printer = self.printer_combo.currentText().strip()
        if not printer:
            QMessageBox.warning(self, "Chưa có máy in", "Không có máy in nào được chọn.")
            return
        try:
            subprocess.Popen([
                "rundll32.exe", "printui.dll,PrintUIEntry", "/o", "/n", printer
            ])
        except Exception as exc:
            QMessageBox.critical(self, "Không mở được hàng đợi in", str(exc))

    def verify_master_before_print(self) -> None:
        if not self.generation_ready:
            raise RuntimeError("Chưa có lô hợp đồng được sinh thành công trong phiên hiện tại.")
        if not self.master_pdf or not self.master_pdf.exists():
            raise RuntimeError("PDF tổng không tồn tại. Hãy tạo lại PDF tổng.")
        if self.master_pdf.parent.resolve() != (self.output_dir() / "MASTER").resolve():
            raise RuntimeError("Thư mục OUTPUT đã thay đổi sau khi tạo PDF tổng.")
        if self.copies_spin.value() != self.master_copies:
            raise RuntimeError("Số bản hiện tại khác PDF tổng. Hãy tạo lại PDF tổng.")
        if (self.print_mode_combo.currentIndex() == 0) != self.master_duplex:
            raise RuntimeError("Chế độ in hiện tại khác PDF tổng. Hãy tạo lại PDF tổng.")
        if not self.master_employees:
            raise RuntimeError("Không xác định được danh sách nhân viên của PDF tổng.")
        selected_keys = [self.employee_key(emp) for emp in self.selected_employees()]
        master_keys = [self.employee_key(emp) for emp in self.master_employees]
        if selected_keys != master_keys:
            raise RuntimeError("Danh sách tick đã thay đổi. Hãy tạo lại PDF tổng.")
        if len(self.current_pdfs()) != len(self.master_employees):
            raise RuntimeError("Thiếu PDF nhân viên trong lô. Tool dừng để tránh in thiếu.")
        validate_employee_batch(self.master_employees, MAX_BATCH_SIZE)

        actual_pages = len(PdfReader(str(self.master_pdf)).pages)
        if actual_pages != self.master_pages:
            raise RuntimeError(
                f"PDF tổng đã thay đổi số trang ({self.master_pages} → {actual_pages})."
            )
        if actual_pages > MAX_PRINT_JOB_PAGES:
            raise RuntimeError(
                f"PDF tổng có {actual_pages} trang, vượt giới hạn {MAX_PRINT_JOB_PAGES} trang/lệnh."
            )
        if not self.master_sha256 or sha256_file(self.master_pdf) != self.master_sha256:
            raise RuntimeError("PDF tổng đã bị thay đổi sau khi tạo. Tool khóa in để tránh sai nội dung.")
        if not self.generated_template_path or not self.generated_template_path.exists():
            raise RuntimeError("Không tìm thấy mẫu Word đã dùng cho lô này.")
        if Path(self.template_edit.text().strip()).resolve() != self.generated_template_path.resolve():
            raise RuntimeError("Mẫu Word hiện tại khác mẫu đã dùng để sinh lô này.")

        current_key = build_print_job_key(
            self.master_employees,
            self.generated_template_path,
            self.master_copies,
            self.master_duplex,
        )
        if not self.master_job_key or current_key != self.master_job_key:
            raise RuntimeError("Dữ liệu hoặc mẫu Word đã thay đổi. Hãy sinh lại lô trước khi in.")

        previous = get_print_guard_record(self.output_dir(), self.master_job_key)
        if (
            previous
            and previous.get("status") in {"SENDING", "SENT", "UNKNOWN"}
            and self.reprint_allowed_job_key != self.master_job_key
        ):
            self.print_locked = True
            self.refresh_action_states()
            raise FileExistsError(
                "Lô này đã từng được gửi tới máy in và đang bị khóa chống in trùng.\n\n"
                f"Trạng thái: {previous.get('status', '')}\n"
                f"Thời gian: {previous.get('updated_at', '')}\n"
                f"Máy in: {previous.get('printer', '')}\n\n"
                "Hãy kiểm tra hàng đợi/máy in. Chỉ dùng 'CHO PHÉP IN LẠI 1 LẦN' "
                "khi chắc chắn cần in lại."
            )

    def allow_reprint_once(self):
        if not self.master_pdf or not self.master_job_key:
            QMessageBox.warning(self, "Chưa có lô", "Hãy tạo PDF tổng trước.")
            return
        try:
            previous = get_print_guard_record(self.output_dir(), self.master_job_key)
        except Exception as exc:
            QMessageBox.critical(self, "Không đọc được lịch sử in", str(exc))
            return
        if not previous:
            QMessageBox.information(self, "Không cần mở khóa", "Lô này chưa có lịch sử gửi in.")
            return
        confirm = QMessageBox.warning(
            self, "MỞ KHÓA IN LẠI MỘT LẦN",
            f"{self.master_batch_label}\n"
            f"Lần gần nhất: {previous.get('updated_at', '')}\n"
            f"Máy in: {previous.get('printer', '')}\n"
            f"Trạng thái: {previous.get('status', '')}\n\n"
            "Chỉ tiếp tục sau khi đã kiểm tra hàng đợi Windows và xác nhận lô cũ "
            "không còn in hoặc thật sự cần in thêm.\n\nCho phép in lại một lần?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if confirm != QMessageBox.StandardButton.Yes:
            return
        self.reprint_allowed_job_key = self.master_job_key
        self.print_locked = False
        self.status.setText("Đã mở khóa in lại 1 lần - hãy kiểm tra kỹ trước khi bấm IN")
        self.log(f"Đã mở khóa in lại một lần cho {self.master_batch_label}.")
        self.refresh_action_states()

    def print_master(self):
        if not self.master_pdf or not self.master_pdf.exists():
            selected = self.selected_employees()
            if not selected:
                QMessageBox.warning(
                    self, "Chưa tick nhân viên",
                    "Hãy tick nhân viên cần in hoặc bấm 'TICK TẤT CẢ', sau đó bấm IN."
                )
                return
            bad = [emp for emp in selected if not emp.valid]
            if bad:
                names = "\n".join(
                    f"- {emp.ma_nv} {emp.ho_ten}: {emp.error}" for emp in bad[:15]
                )
                QMessageBox.warning(
                    self, "Dữ liệu được tick chưa hợp lệ",
                    "Tool chưa in vì các dòng sau đang lỗi:\n" + names
                )
                return
            if not self.selection_has_generated_pdfs():
                pdf_map = self.generated_pdf_map()
                missing = [
                    f"- {emp.ma_nv} {emp.ho_ten}"
                    for emp in selected
                    if self.employee_key(emp) not in pdf_map
                ]
                QMessageBox.warning(
                    self, "Chưa sinh đủ hợp đồng",
                    "Nhân viên đã tick nhưng chưa có PDF hợp đồng:\n"
                    + "\n".join(missing[:15])
                    + "\n\nHãy giữ nguyên các dấu tick và bấm '2. SINH HỢP ĐỒNG', "
                      "sau đó bấm IN lại."
                )
                return
            self.status.setText("Đang tạo PDF tổng từ các nhân viên đã tick...")
            self.log(
                f"Bắt đầu lệnh in nhanh: tự tạo PDF tổng cho {len(selected)} nhân viên đã tick."
            )
            self.create_master(auto_print=True)
            return

        try:
            self.verify_master_before_print()
        except Exception as exc:
            self.log(str(exc))
            QMessageBox.critical(self, "ĐÃ DỪNG LỆNH IN", str(exc))
            return
        printer = self.printer_combo.currentText().strip()
        if not printer:
            QMessageBox.warning(self, "Chưa có máy in", "Không có máy in nào được chọn.")
            return
        printer_errors = printer_blocking_status(printer)
        if printer_errors:
            detail = "\n".join(f"- {message}" for message in printer_errors)
            self.log("Dừng in do trạng thái máy in: " + "; ".join(printer_errors))
            QMessageBox.critical(
                self, "MÁY IN CHƯA SẴN SÀNG",
                detail + "\n\nTool chưa gửi bất kỳ trang nào. Hãy xử lý máy in rồi thử lại."
            )
            return
        contracts = len(self.master_employees)
        copies = self.master_copies
        duplex = self.master_duplex
        print_mode_text = "2 mặt A4 - lật cạnh dài" if duplex else "1 mặt A4"
        sheets_text = f"\nSố tờ A4 dự kiến: {(self.master_pages + 1) // 2}" if duplex else ""
        confirm = QMessageBox.question(
            self, "XÁC NHẬN IN",
            f"{self.master_batch_label}\n"
            f"Máy in: {printer}\n"
            f"Số nhân viên: {contracts}\n"
            f"Số bản / người: {copies}\n"
            f"Chế độ in: {print_mode_text}\n"
            f"Tổng số trang: {self.master_pages}"
            f"{sheets_text}\n\n"
            "Tool sẽ khóa lô ngay khi gửi để chống in trùng.\n"
            "Bạn chắc chắn muốn gửi lệnh in?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if confirm != QMessageBox.StandardButton.Yes:
            return

        guard_record = {
            "status": "SENDING",
            "printer": printer,
            "batch": self.master_batch_label,
            "contracts": contracts,
            "copies": copies,
            "pages": self.master_pages,
            "master_pdf": str(self.master_pdf),
            "master_sha256": self.master_sha256,
        }
        try:
            save_print_guard_record(self.output_dir(), self.master_job_key, guard_record)
        except Exception as exc:
            self.print_locked = False
            self.refresh_action_states()
            QMessageBox.critical(
                self, "KHÔNG GHI ĐƯỢC KHÓA IN",
                f"Tool chưa gửi lệnh in vì không ghi được lịch sử chống in trùng.\n\n{exc}"
            )
            return

        try:
            self.set_busy(True); self.status.setText("Đang gửi lệnh in...")
            print_pdf(self.master_pdf, printer, duplex=duplex)
            guard_record["status"] = "SENT"
            save_print_guard_record(self.output_dir(), self.master_job_key, guard_record)
            write_print_log(self.output_dir(), printer, contracts, copies, self.master_pages, self.master_pdf, "SENT")
            self.print_locked = True
            self.reprint_allowed_job_key = None
            self.log(f"Đã gửi lệnh in tới: {printer}")
            QMessageBox.information(
                self, "Đã gửi và khóa lô",
                "PDF tổng đã được gửi tới máy in. Tool đã khóa lô này để tránh in trùng.\n\n"
                "Hãy kiểm tra hàng đợi Windows Printer trước khi chuyển sang lô tiếp theo."
            )
            self.status.setText("Đã gửi in và khóa lô - có thể chuyển sang LÔ TIẾP THEO")
            self.batch_status.setText(self.master_batch_label + " - ĐÃ GỬI IN")
        except Exception as exc:
            logging.exception("Print failed")
            guard_record["status"] = "UNKNOWN"
            guard_record["error"] = str(exc)
            try:
                save_print_guard_record(self.output_dir(), self.master_job_key, guard_record)
            except Exception:
                logging.exception("Khong cap nhat duoc print guard sau loi in")
            self.print_locked = True
            self.reprint_allowed_job_key = None
            write_print_log(self.output_dir(), printer, contracts, copies, self.master_pages, self.master_pdf, "UNKNOWN", str(exc))
            self.log(str(exc))
            QMessageBox.critical(
                self, "TRẠNG THÁI IN KHÔNG CHẮC CHẮN - ĐÃ KHÓA",
                str(exc)
                + "\n\nTool đã khóa in lại để tránh gửi trùng. Hãy kiểm tra hàng đợi Windows "
                  "và máy in trước khi dùng nút cho phép in lại."
            )
            self.status.setText("Lỗi/không chắc chắn - kiểm tra hàng đợi trước khi in lại")
        finally:
            self.set_busy(False)

    def closeEvent(self, event):
        self.persist()
        event.accept()


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    window = MainWindow()
    window.showMaximized()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
