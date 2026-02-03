import sys, math, datetime
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Tuple

import numpy as np
import openpyxl
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QFileDialog, QMessageBox,
    QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QSpinBox,
    QComboBox, QCheckBox, QFrame, QScrollArea, QGridLayout
)

APP_TITLE = "Commissioning Budget Tool"

# Business rules
TRAINING_MACHINES_PER_DAY = 3  # 1 training day per 3 machines (ceil)
DEFAULT_INSTALL_WINDOW = 7
MIN_INSTALL_WINDOW = 3
MAX_INSTALL_WINDOW = 14
TRAVEL_DAYS_PER_PERSON = 2  # travel-in + travel-out

# Requested overrides
OVERRIDE_AIRFARE_PER_PERSON = 1500.0
OVERRIDE_BAGGAGE_PER_DAY_PER_PERSON = 150.0

ASSETS_DIR = Path(__file__).resolve().parent / "assets"
DEFAULT_EXCEL = ASSETS_DIR / "Tech days and quote rates.xlsx"
LOGO_PATH = ASSETS_DIR / "Pearson Logo.png"


def ceil_int(x: float) -> int:
    return int(math.ceil(float(x)))


def balanced_allocate(total_days: int, headcount: int) -> List[int]:
    """Balance integer days to minimize the maximum assigned days."""
    if headcount <= 0:
        return []
    loads = [0] * headcount
    for _ in range(int(total_days)):
        i = int(np.argmin(loads))
        loads[i] += 1
    loads.sort(reverse=True)
    return loads


@dataclass
class ModelInfo:
    item: str
    tech_install_days_per_machine: int   # install-only (no training baked in)
    eng_days_per_machine: int


@dataclass
class LineSelection:
    model: str
    qty: int
    training_required: bool


@dataclass
class RoleTotals:
    headcount: int
    total_onsite_days: int
    onsite_days_by_person: List[int]
    day_rate: float
    labor_cost: float


@dataclass
class ExpenseLine:
    description: str
    quantity: float
    unit_price: float
    extended: float
    details: str


class ExcelData:
    def __init__(self, path: Path):
        self.path = path
        self.models: Dict[str, ModelInfo] = {}
        self.rates: Dict[str, Dict[str, object]] = {}
        self.requirements: List[str] = []
        self._load()

    def _load(self):
        wb = openpyxl.load_workbook(self.path, data_only=True)

        # Models: Instal days by Model
        if "Instal days by Model" not in wb.sheetnames:
            raise ValueError("Missing sheet: 'Instal days by Model'")
        ws = wb["Instal days by Model"]

        # Expected headers in row 1
        headers = {str(ws.cell(1, c).value).strip(): c for c in range(1, ws.max_column + 1) if ws.cell(1, c).value is not None}

        def find_col(pred):
            for k, c in headers.items():
                if pred(k.lower()):
                    return c
            return None

        col_item = find_col(lambda s: s in ["item", "model", "machine", "machine type"])
        col_tech = find_col(lambda s: "technician" in s and "day" in s)
        col_eng = find_col(lambda s: "engineer" in s and "day" in s)

        if col_item is None or col_tech is None or col_eng is None:
            raise ValueError("Model sheet columns not found. Expected: Item, Technician Days Required, Field Engineer Days Required.")

        for r in range(2, ws.max_row + 1):
            item = ws.cell(r, col_item).value
            if item is None:
                continue
            item = str(item).strip()
            if not item:
                continue
            tech = ws.cell(r, col_tech).value or 0
            eng = ws.cell(r, col_eng).value or 0
            try:
                tech_i = int(float(tech))
            except Exception:
                tech_i = 0
            try:
                eng_i = int(float(eng))
            except Exception:
                eng_i = 0
            self.models[item] = ModelInfo(item=item, tech_install_days_per_machine=tech_i, eng_days_per_machine=eng_i)

        # Rates: Service Rates
        if "Service Rates" not in wb.sheetnames:
            raise ValueError("Missing sheet: 'Service Rates'")
        ws = wb["Service Rates"]

        # Detect header row by "Item" and "Description"
        header_row = None
        for r in range(1, 12):
            if ws.cell(r, 2).value == "Item" and ws.cell(r, 3).value == "Description":
                header_row = r
                break
        if header_row is None:
            header_row = 3

        for r in range(header_row + 1, ws.max_row + 1):
            desc = ws.cell(r, 3).value
            if desc is None:
                continue
            desc_s = str(desc).strip()
            if not desc_s:
                continue
            unit = ws.cell(r, 6).value
            notes = ws.cell(r, 7).value
            try:
                unit_f = float(unit)
            except Exception:
                continue
            self.rates[desc_s.lower()] = {
                "description": desc_s,
                "unit_price": unit_f,
                "notes": str(notes).strip() if notes is not None else ""
            }

        # Requirements & Assumptions
        if "Requirements and Assumptions" in wb.sheetnames:
            ws = wb["Requirements and Assumptions"]
            out = []
            for r in range(1, ws.max_row + 1):
                v = ws.cell(r, 3).value
                if v is None:
                    continue
                s = str(v).strip()
                if s and not s.lower().startswith("assumptions and requirements"):
                    out.append(s)
            self.requirements = out

    def get_rate(self, key: str) -> Tuple[float, str]:
        k = key.lower().strip()
        if k in self.rates:
            rv = self.rates[k]
            return float(rv["unit_price"]), str(rv["description"])
        # fuzzy contains
        for rk, rv in self.rates.items():
            if k in rk:
                return float(rv["unit_price"]), str(rv["description"])
        raise KeyError(f"Rate not found for '{key}'")


class MachineLine(QWidget):
    def __init__(self, models: List[str], on_change, on_delete):
        super().__init__()
        self.on_change = on_change
        self.on_delete = on_delete

        row = QHBoxLayout(self)
        row.setContentsMargins(0, 0, 0, 0)

        self.cmb_model = QComboBox()
        self.cmb_model.addItems(models)
        self.cmb_model.currentIndexChanged.connect(self._changed)

        self.spin_qty = QSpinBox()
        self.spin_qty.setRange(0, 999)
        self.spin_qty.setValue(1)
        self.spin_qty.valueChanged.connect(self._changed)

        self.chk_training = QCheckBox("Training Required")
        self.chk_training.setChecked(True)
        self.chk_training.stateChanged.connect(self._changed)

        self.btn_delete = QPushButton("🗑")
        self.btn_delete.setFixedWidth(40)
        self.btn_delete.clicked.connect(self._delete)

        row.addWidget(QLabel("Model"))
        row.addWidget(self.cmb_model, 2)
        row.addWidget(QLabel("Qty"))
        row.addWidget(self.spin_qty)
        row.addWidget(self.chk_training, 1)
        row.addWidget(self.btn_delete)

    def _changed(self, *_):
        self.on_change()

    def _delete(self):
        self.on_delete(self)

    def value(self) -> LineSelection:
        return LineSelection(
            model=self.cmb_model.currentText().strip(),
            qty=int(self.spin_qty.value()),
            training_required=bool(self.chk_training.isChecked())
        )


class Card(QFrame):
    def __init__(self, title: str):
        super().__init__()
        self.setObjectName("card")
        lay = QVBoxLayout(self)
        lay.setContentsMargins(14, 12, 14, 12)
        self.lbl_title = QLabel(title)
        self.lbl_title.setObjectName("cardTitle")
        self.lbl_value = QLabel("—")
        self.lbl_value.setObjectName("cardValue")
        self.lbl_sub = QLabel("")
        self.lbl_sub.setObjectName("cardSub")
        self.lbl_sub.setWordWrap(True)
        lay.addWidget(self.lbl_title)
        lay.addWidget(self.lbl_value)
        lay.addWidget(self.lbl_sub)

    def set_value(self, value: str, sub: str = ""):
        self.lbl_value.setText(value)
        self.lbl_sub.setText(sub or "")


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(1320, 780)

        self.data = ExcelData(DEFAULT_EXCEL)
        self.models_sorted = sorted(self.data.models.keys())
        self.lines: List[MachineLine] = []

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)

        header = QFrame()
        header.setObjectName("header")
        h = QHBoxLayout(header)
        h.setContentsMargins(14, 10, 14, 10)
        h.addWidget(QLabel("☰  " + APP_TITLE))
        h.addStretch(1)
        btn_excel = QPushButton("Open Excel…")
        btn_excel.clicked.connect(self.open_excel)
        h.addWidget(btn_excel)
        root.addWidget(header)

        body = QHBoxLayout()
        body.setContentsMargins(14, 14, 14, 14)
        root.addLayout(body, 1)

        # Left panel
        left = QFrame()
        left.setObjectName("panel")
        left_l = QVBoxLayout(left)
        left_l.setContentsMargins(14, 14, 14, 14)

        title = QLabel("Machine Configuration")
        title.setObjectName("sectionTitle")
        left_l.addWidget(title)
        left_l.addWidget(QLabel(
            "Add machines to estimate commissioning requirements.\n"
            "Each machine type requires personnel with different skills (no sharing across types)."
        ))

        win_box = QFrame()
        win_box.setObjectName("softBox")
        win_l = QHBoxLayout(win_box)
        win_l.setContentsMargins(12, 10, 12, 10)
        win_l.addWidget(QLabel("Customer Install Window"))
        self.spin_window = QSpinBox()
        self.spin_window.setRange(MIN_INSTALL_WINDOW, MAX_INSTALL_WINDOW)
        self.spin_window.setValue(DEFAULT_INSTALL_WINDOW)
        self.spin_window.valueChanged.connect(self.recalc)
        win_l.addStretch(1)
        win_l.addWidget(self.spin_window)
        win_l.addWidget(QLabel("days"))
        left_l.addWidget(win_box)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        container = QWidget()
        self.lines_layout = QVBoxLayout(container)
        self.lines_layout.setContentsMargins(0, 0, 0, 0)
        self.lines_layout.setSpacing(10)
        self.scroll.setWidget(container)
        left_l.addWidget(self.scroll, 1)

        btn_add = QPushButton("+  Add Machine")
        btn_add.setObjectName("primary")
        btn_add.clicked.connect(self.add_line)
        left_l.addWidget(btn_add)

        note = QLabel("Note: Unchecking “Training Required” should only be done by customer request.")
        note.setObjectName("note")
        note.setWordWrap(True)
        left_l.addWidget(note)

        body.addWidget(left, 2)

        # Right panel
        right = QVBoxLayout()
        body.addLayout(right, 3)

        cards = QHBoxLayout()
        self.card_tech = Card("Technicians")
        self.card_eng = Card("Engineers")
        self.card_window = Card("Customer Install Window")
        self.card_total = Card("Total Cost")
        cards.addWidget(self.card_tech)
        cards.addWidget(self.card_eng)
        cards.addWidget(self.card_window)
        cards.addWidget(self.card_total)
        right.addLayout(cards)

        self.breakdown = QFrame()
        self.breakdown.setObjectName("panel")
        b = QVBoxLayout(self.breakdown)
        b.setContentsMargins(14, 14, 14, 14)

        self.lbl_break_title = QLabel("Cost Breakdown")
        self.lbl_break_title.setObjectName("sectionTitle")
        b.addWidget(self.lbl_break_title)

        self.lbl_summary = QLabel("Add at least one machine line.")
        self.lbl_summary.setWordWrap(True)
        b.addWidget(self.lbl_summary)

        grid = QGridLayout()
        grid.setHorizontalSpacing(14)
        grid.setVerticalSpacing(10)

        for i, t in enumerate(["Machine Configuration", "Labor", "Estimated Expenses"]):
            lbl = QLabel(t)
            lbl.setObjectName("subTitle")
            grid.addWidget(lbl, 0, i)

        self.tbl_machines = QLabel("")
        self.tbl_labor = QLabel("")
        self.tbl_exp = QLabel("")
        for w in [self.tbl_machines, self.tbl_labor, self.tbl_exp]:
            w.setObjectName("mono")
            w.setTextInteractionFlags(Qt.TextSelectableByMouse)

        grid.addWidget(self.tbl_machines, 1, 0)
        grid.addWidget(self.tbl_labor, 1, 1)
        grid.addWidget(self.tbl_exp, 1, 2)

        lbl_assign = QLabel("Workload Distribution")
        lbl_assign.setObjectName("subTitle")
        grid.addWidget(lbl_assign, 2, 0, 1, 3)
        self.tbl_assign = QLabel("")
        self.tbl_assign.setObjectName("mono")
        self.tbl_assign.setTextInteractionFlags(Qt.TextSelectableByMouse)
        grid.addWidget(self.tbl_assign, 3, 0, 1, 3)

        b.addLayout(grid)

        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        self.btn_print = QPushButton("Print Quote…")
        self.btn_print.clicked.connect(self.print_quote)
        self.btn_print.setEnabled(False)
        btn_row.addWidget(self.btn_print)
        b.addLayout(btn_row)

        right.addWidget(self.breakdown, 1)

        self.add_line()
        self.apply_theme()

    def apply_theme(self):
        blue = "#0B3D66"
        gold = "#D39A2C"
        self.setStyleSheet(f"""
        QFrame#header {{ background: {blue}; color: white; border: none; }}
        QFrame#panel {{ background: white; border: 1px solid #E6E8EB; border-radius: 14px; }}
        QFrame#softBox {{ background: #FFF7EA; border: 1px solid #F0D8A8; border-radius: 12px; }}
        QLabel#sectionTitle {{ font-size: 16px; font-weight: 700; color: #0F172A; }}
        QLabel#subTitle {{ font-size: 13px; font-weight: 700; color: #0F172A; }}
        QLabel#note {{ color: #334155; font-size: 12px; }}
        QPushButton#primary {{
            background: {gold}; border: 0px; color: #0B1B2A;
            padding: 10px 12px; border-radius: 10px; font-weight: 700;
        }}
        QPushButton {{
            padding: 8px 10px; border-radius: 10px;
            border: 1px solid #D6D9DD; background: #F8FAFC;
        }}
        QPushButton:disabled {{ color: #94A3B8; background: #F1F5F9; }}
        QFrame#card {{ background: #FFFFFF; border: 1px solid #E6E8EB; border-radius: 14px; min-width: 190px; }}
        QLabel#cardTitle {{ font-size: 12px; color: #475569; font-weight: 600; }}
        QLabel#cardValue {{ font-size: 24px; color: #0F172A; font-weight: 800; }}
        QLabel#cardSub {{ font-size: 12px; color: #475569; }}
        QLabel#mono {{
            font-family: Consolas, 'Courier New', monospace; font-size: 12px; color: #0F172A;
            background: #FAFAFB; border: 1px solid #EEF0F2; border-radius: 10px; padding: 10px;
        }}
        """)

    def add_line(self):
        ln = MachineLine(self.models_sorted, on_change=self.recalc, on_delete=self.delete_line)
        self.lines.append(ln)
        self.lines_layout.addWidget(ln)
        self.recalc()

    def delete_line(self, ln: MachineLine):
        if len(self.lines) <= 1:
            return
        self.lines.remove(ln)
        ln.setParent(None)
        ln.deleteLater()
        self.recalc()

    def open_excel(self):
        fp, _ = QFileDialog.getOpenFileName(self, "Select Excel file", "", "Excel (*.xlsx)")
        if not fp:
            return
        try:
            self.data = ExcelData(Path(fp))
            self.models_sorted = sorted(self.data.models.keys())
            for ln in self.lines:
                cur = ln.cmb_model.currentText()
                ln.cmb_model.blockSignals(True)
                ln.cmb_model.clear()
                ln.cmb_model.addItems(self.models_sorted)
                if cur in self.models_sorted:
                    ln.cmb_model.setCurrentIndex(self.models_sorted.index(cur))
                ln.cmb_model.blockSignals(False)
            self.recalc()
        except Exception as e:
            QMessageBox.critical(self, "Excel load error", str(e))

    # ---- core calculations ----
    def calc(self):
        selections = [ln.value() for ln in self.lines]
        selections = [s for s in selections if s.qty > 0]
        if not selections:
            raise ValueError("No machines selected.")

        window = int(self.spin_window.value())

        # Rates: Tech/Eng regular time are hourly, converted to daily at 8 hrs/day
        tech_hr, _ = self.data.get_rate("tech. regular time")
        eng_hr, _ = self.data.get_rate("eng. regular time")
        hours_per_day = 8
        tech_day_rate = tech_hr * hours_per_day
        eng_day_rate = eng_hr * hours_per_day

        tech_person_days: List[int] = []
        eng_person_days: List[int] = []
        machine_rows = []

        for s in selections:
            mi = self.data.models[s.model]

            training_days = ceil_int(s.qty / TRAINING_MACHINES_PER_DAY) if s.training_required else 0

            # INSTALL-ONLY baseline from Excel (no training baked in)
            tech_install_total = mi.tech_install_days_per_machine * s.qty
            tech_total = tech_install_total + training_days

            eng_total = mi.eng_days_per_machine * s.qty

            # Validation: single-machine commissioning (install + training if checked) must fit window
            single_training = 1 if s.training_required else 0  # ceil(1/3)=1
            if mi.tech_install_days_per_machine + single_training > window:
                raise ValueError(
                    f"{s.model}: Install ({mi.tech_install_days_per_machine}) + Training ({single_training}) exceeds the Customer Install Window ({window})."
                )
            if mi.eng_days_per_machine > window and mi.eng_days_per_machine > 0:
                raise ValueError(
                    f"{s.model}: Engineer days for a single machine ({mi.eng_days_per_machine}) exceeds the Customer Install Window ({window})."
                )

            # Dedicated staffing per model (no sharing across model types)
            if tech_total > 0:
                tech_headcount = ceil_int(tech_total / window)
                tech_person_days.extend(balanced_allocate(tech_total, tech_headcount))
            if eng_total > 0:
                eng_headcount = ceil_int(eng_total / window)
                eng_person_days.extend(balanced_allocate(eng_total, eng_headcount))

            machine_rows.append({
                "model": s.model,
                "qty": s.qty,
                "tech_install_per_machine": mi.tech_install_days_per_machine,
                "training_days": training_days,
                "training_required": s.training_required,
                "tech_total": tech_total,
                "eng_per_machine": mi.eng_days_per_machine,
                "eng_total": eng_total
            })

        tech = RoleTotals(
            headcount=len(tech_person_days),
            total_onsite_days=sum(tech_person_days),
            onsite_days_by_person=tech_person_days,
            day_rate=tech_day_rate,
            labor_cost=float(sum(tech_person_days)) * tech_day_rate
        )
        eng = RoleTotals(
            headcount=len(eng_person_days),
            total_onsite_days=sum(eng_person_days),
            onsite_days_by_person=eng_person_days,
            day_rate=eng_day_rate,
            labor_cost=float(sum(eng_person_days)) * eng_day_rate
        )

        # Expenses are based on person-days including travel in/out per person
        trip_days_by_person = [d + TRAVEL_DAYS_PER_PERSON for d in tech_person_days] + [d + TRAVEL_DAYS_PER_PERSON for d in eng_person_days]
        n_people = len(trip_days_by_person)
        total_trip_days = sum(trip_days_by_person)  # person-days
        total_hotel_nights = sum(max(d - 1, 0) for d in trip_days_by_person)

        def add_exp(lines, name, qty, unit, detail):
            lines.append(ExpenseLine(name, float(qty), float(unit), float(qty) * float(unit), detail))

        exp_lines: List[ExpenseLine] = []

        # Airfare override
        try:
            _, _ = self.data.get_rate("airfare")
        except KeyError:
            pass
        add_exp(exp_lines, "Airfare", n_people, OVERRIDE_AIRFARE_PER_PERSON, f"{n_people} person(s) × ${OVERRIDE_AIRFARE_PER_PERSON:,.0f}")

        # Baggage override (per day per person) — displayed as total person-days × rate
        add_exp(exp_lines, "Baggage", total_trip_days, OVERRIDE_BAGGAGE_PER_DAY_PER_PERSON, f"{int(total_trip_days)} day(s) × ${OVERRIDE_BAGGAGE_PER_DAY_PER_PERSON:,.0f}")

        # Other expense lines from sheet (per day/person -> person-days)
        parking, _ = self.data.get_rate("parking")
        car, _ = self.data.get_rate("car rental")
        hotel, _ = self.data.get_rate("hotel")
        per_diem, _ = self.data.get_rate("per diem weekday")
        prep, _ = self.data.get_rate("pre/post trip prep")
        travel_time_rate, _ = self.data.get_rate("travel time")

        add_exp(exp_lines, "Parking", total_trip_days, parking, f"{int(total_trip_days)} day(s) × ${parking:,.0f}")
        add_exp(exp_lines, "Car Rental", total_trip_days, car, f"{int(total_trip_days)} day(s) × ${car:,.0f}")
        add_exp(exp_lines, "Hotel", total_hotel_nights, hotel, f"{int(total_hotel_nights)} night(s) × ${hotel:,.0f}")
        add_exp(exp_lines, "Per Diem", total_trip_days, per_diem, f"{int(total_trip_days)} day(s) × ${per_diem:,.0f}")
        add_exp(exp_lines, "Pre/Post Trip Prep", n_people, prep, f"{n_people} person(s) × ${prep:,.0f}")

        # Travel time: 16 hours per person
        travel_hours = 16 * n_people
        add_exp(exp_lines, "Travel Time", travel_hours, travel_time_rate, f"{travel_hours} hr(s) × ${travel_time_rate:,.0f}")

        exp_total = sum(l.extended for l in exp_lines)
        max_onsite = max(tech_person_days + eng_person_days) if (tech_person_days or eng_person_days) else 0
        grand_total = exp_total + tech.labor_cost + eng.labor_cost

        meta = {
            "machine_rows": machine_rows,
            "window": window,
            "max_onsite": max_onsite,
            "n_people": n_people,
            "total_trip_days": total_trip_days,
            "exp_total": exp_total,
            "grand_total": grand_total
        }
        return tech, eng, exp_lines, meta

    def recalc(self):
        try:
            tech, eng, exp_lines, meta = self.calc()

            self.card_tech.set_value(str(tech.headcount), f"{tech.total_onsite_days} total onsite days")
            self.card_eng.set_value(str(eng.headcount), f"{eng.total_onsite_days} total onsite days")
            self.card_window.set_value(f"{meta['window']} days", f"Estimated duration: {meta['max_onsite']} days onsite + {TRAVEL_DAYS_PER_PERSON} travel days")
            self.card_total.set_value(f"${meta['grand_total']:,.0f}", "labor + expenses")

            self.lbl_summary.setText(
                f"Expenses include {TRAVEL_DAYS_PER_PERSON} travel days per person (travel-in + travel-out). "
                f"Machine types require specialized skills: personnel are not shared across different machine types."
            )

            # Machine block
            lines = []
            for r in meta["machine_rows"]:
                tech_disp = f"{r['tech_install_per_machine']} (incl. Train)" if r["training_required"] else f"{r['tech_install_per_machine']} (training excluded)"
                train_note = (f"Train days: {r['training_days']} (1 per {TRAINING_MACHINES_PER_DAY} machines)"
                              if r["training_required"] else "Train days: 0 (excluded by customer request)")
                lines.append(
                    f"- {r['model']}  Qty {r['qty']}  |  Tech Install Days: {tech_disp}  |  Eng Days: {r['eng_per_machine']}\n"
                    f"  Tech total: {r['tech_total']}  ({train_note})  |  Eng total: {r['eng_total']}"
                )
            self.tbl_machines.setText("\n".join(lines))

            labor_txt = (
                f"Tech. Regular Time: {tech.total_onsite_days} day(s) × ${tech.day_rate:,.0f}/day = ${tech.labor_cost:,.0f}\n"
                f"Eng. Regular Time:  {eng.total_onsite_days} day(s) × ${eng.day_rate:,.0f}/day = ${eng.labor_cost:,.0f}\n"
                f"Labor Subtotal: ${tech.labor_cost + eng.labor_cost:,.0f}"
            )
            self.tbl_labor.setText(labor_txt)

            exp_txt = [f"Includes {meta['total_trip_days']:.0f} total trip day(s) across {meta['n_people']} person(s)"]
            for l in exp_lines:
                exp_txt.append(f"{l.description}: {l.details} = ${l.extended:,.0f}")
            exp_txt.append(f"Expenses Subtotal: ${meta['exp_total']:,.0f}")
            self.tbl_exp.setText("\n".join(exp_txt))

            assign_lines = []
            if tech.headcount:
                assign_lines.append("Technicians:")
                for i, d in enumerate(tech.onsite_days_by_person, 1):
                    assign_lines.append(f"  Tech {i}: {d} day(s) onsite")
            if eng.headcount:
                assign_lines.append("Engineers:")
                for i, d in enumerate(eng.onsite_days_by_person, 1):
                    assign_lines.append(f"  Eng {i}: {d} day(s) onsite")
            self.tbl_assign.setText("\n".join(assign_lines))

            self.btn_print.setEnabled(True)

        except Exception as e:
            self.card_tech.set_value("—", "")
            self.card_eng.set_value("—", "")
            self.card_window.set_value(f"{self.spin_window.value()} days", "")
            self.card_total.set_value("—", "")
            self.tbl_machines.setText("")
            self.tbl_labor.setText("")
            self.tbl_exp.setText("")
            self.tbl_assign.setText("")
            self.lbl_summary.setText(str(e))
            self.btn_print.setEnabled(False)

    # ---- PDF quote ----
    def print_quote(self):
        try:
            tech, eng, exp_lines, meta = self.calc()
        except Exception as e:
            QMessageBox.critical(self, "Cannot print", str(e))
            return

        fp, _ = QFileDialog.getSaveFileName(self, "Save PDF", "Commissioning Budget Quote.pdf", "PDF (*.pdf)")
        if not fp:
            return
        try:
            self._build_pdf(Path(fp), tech, eng, exp_lines, meta)
            QMessageBox.information(self, "Saved", f"Quote saved:\n{fp}")
        except Exception as e:
            QMessageBox.critical(self, "PDF error", str(e))

    def _build_pdf(self, out: Path, tech: RoleTotals, eng: RoleTotals, exp_lines: List[ExpenseLine], meta: Dict[str, object]):
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle("Title", parent=styles["Title"], fontSize=22, leading=26, textColor=colors.HexColor("#0B3D66"))
        h_style = ParagraphStyle("H", parent=styles["Heading2"], fontSize=13, textColor=colors.HexColor("#0B3D66"), spaceBefore=10, spaceAfter=6)
        small = ParagraphStyle("Small", parent=styles["Normal"], fontSize=9, leading=12, textColor=colors.HexColor("#334155"))
        normal = ParagraphStyle("Normal2", parent=styles["Normal"], fontSize=10, leading=13)

        doc = SimpleDocTemplate(str(out), pagesize=letter, leftMargin=0.75*inch, rightMargin=0.75*inch, topMargin=0.65*inch, bottomMargin=0.65*inch)
        story = []

        # Header with logo
        left = Paragraph("<b>Service Estimate</b><br/>Commissioning Budget Quote", title_style)
        logo = RLImage(str(LOGO_PATH), width=2.2*inch, height=0.75*inch) if LOGO_PATH.exists() else Paragraph("", normal)
        t = Table([[left, logo]], colWidths=[4.4*inch, 2.6*inch])
        t.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP"), ("ALIGN",(1,0),(1,0),"RIGHT"), ("BOTTOMPADDING",(0,0),(-1,-1),8)]))
        story.append(t)
        story.append(Spacer(1, 6))
        story.append(Table([[""]], colWidths=[7.0*inch], style=[("LINEBELOW",(0,0),(-1,-1),1,colors.HexColor("#0B3D66"))]))
        story.append(Spacer(1, 16))

        today = datetime.date.today()
        validity = today + datetime.timedelta(days=30)

        total_personnel = f"{tech.headcount + eng.headcount} ({tech.headcount} Tech, {eng.headcount} Eng)"
        est_dur = f"{meta['max_onsite']} days onsite + {TRAVEL_DAYS_PER_PERSON} travel days"

        info2 = Table([
            [Paragraph("<b>DATE</b><br/>" + today.strftime("%B %-d, %Y"), normal),
             Paragraph("<b>QUOTE VALIDITY</b><br/>" + validity.strftime("%B %-d, %Y"), normal)],
            [Paragraph("<b>TOTAL PERSONNEL</b><br/>" + total_personnel, normal),
             Paragraph("<b>ESTIMATED DURATION</b><br/>" + est_dur, normal)],
        ], colWidths=[3.4*inch, 3.4*inch])
        info2.setStyle(TableStyle([("FONTSIZE",(0,0),(-1,-1),10), ("BOTTOMPADDING",(0,0),(-1,-1),10)]))
        story.append(info2)
        story.append(Spacer(1, 10))

        scope_text = (f"This quote reflects <b>{tech.headcount}</b> technician(s) and <b>{eng.headcount}</b> engineer(s) working "
                      f"<b>{meta['max_onsite']}</b> day(s) onsite for commissioning and training of the equipment listed below. "
                      f"Expenses include {TRAVEL_DAYS_PER_PERSON} travel days (travel-in and travel-out).")
        scope_tbl = Table([[Paragraph("<b>Scope of Work</b>", h_style)], [Paragraph(scope_text, normal)]], colWidths=[6.8*inch])
        scope_tbl.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(0,0),colors.HexColor("#FFF7EA")),
            ("BOX",(0,0),(-1,-1),1,colors.HexColor("#D39A2C")),
            ("BACKGROUND",(0,1),(0,1),colors.HexColor("#FFFDF7")),
            ("LEFTPADDING",(0,0),(-1,-1),10),
            ("RIGHTPADDING",(0,0),(-1,-1),10),
            ("TOPPADDING",(0,0),(-1,-1),8),
            ("BOTTOMPADDING",(0,0),(-1,-1),8),
        ]))
        story.append(scope_tbl)
        story.append(Spacer(1, 14))

        # Machine table
        story.append(Paragraph("Machine Configuration", h_style))
        rows = [["Model", "Qty", "Tech Install Days", "Eng Days", "Personnel"]]
        window = int(meta["window"])
        for r in meta["machine_rows"]:
            tech_disp = f"{r['tech_install_per_machine']} (incl. Train)" if r["training_required"] else f"{r['tech_install_per_machine']} (training excluded)"
            eng_disp = "—" if r["eng_per_machine"] == 0 else str(r["eng_per_machine"])
            tech_head = ceil_int(r["tech_total"]/window) if r["tech_total"] > 0 else 0
            eng_head = ceil_int(r["eng_total"]/window) if r["eng_total"] > 0 else 0
            pers = []
            if tech_head: pers.append(f"{tech_head} Tech")
            if eng_head: pers.append(f"{eng_head} Eng")
            rows.append([r["model"], str(r["qty"]), tech_disp, eng_disp, ", ".join(pers) if pers else "—"])
        mt = Table(rows, colWidths=[1.4*inch, 0.8*inch, 2.2*inch, 1.2*inch, 1.2*inch])
        mt.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F1F5F9")),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
            ("FONTSIZE",(0,0),(-1,-1),9),
            ("GRID",(0,0),(-1,-1),0.5,colors.HexColor("#E2E8F0")),
            ("ALIGN",(1,1),(1,-1),"CENTER"),
            ("ALIGN",(3,1),(4,-1),"CENTER"),
            ("BOTTOMPADDING",(0,0),(-1,-1),6),
            ("TOPPADDING",(0,0),(-1,-1),6),
        ]))
        story.append(mt)
        story.append(Spacer(1, 14))

        # Labor
        story.append(Paragraph("Labor Costs", h_style))
        labor_rows = [["Item", "Quantity", "Unit Price", "Extended Price"]]
        if tech.total_onsite_days:
            labor_rows.append(["Tech. Regular Time", f"{tech.total_onsite_days} days", f"${tech.day_rate:,.0f}/day", f"${tech.labor_cost:,.0f}"])
        if eng.total_onsite_days:
            labor_rows.append(["Eng. Regular Time", f"{eng.total_onsite_days} days", f"${eng.day_rate:,.0f}/day", f"${eng.labor_cost:,.0f}"])
        labor_sub = tech.labor_cost + eng.labor_cost
        labor_rows.append(["", "", "<b>Labor Subtotal</b>", f"<b>${labor_sub:,.0f}</b>"])
        lt = Table(labor_rows, colWidths=[2.6*inch, 1.3*inch, 1.4*inch, 1.5*inch])
        lt.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F1F5F9")),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
            ("FONTSIZE",(0,0),(-1,-1),9),
            ("GRID",(0,0),(-1,-1),0.5,colors.HexColor("#E2E8F0")),
            ("ALIGN",(1,1),(3,-2),"RIGHT"),
            ("ALIGN",(2,-1),(3,-1),"RIGHT"),
        ]))
        story.append(lt)
        story.append(Spacer(1, 14))

        # Expenses
        story.append(Paragraph("Estimated Expenses", h_style))
        story.append(Paragraph(f"Includes {int(meta['total_trip_days'])} total trip day(s) across personnel (onsite + travel days).", small))
        story.append(Spacer(1, 6))
        exp_rows = [["Item", "Details", "Extended Price"]]
        for l in exp_lines:
            exp_rows.append([l.description, l.details, f"${l.extended:,.0f}"])
        exp_rows.append(["", "<b>Expenses Subtotal</b>", f"<b>${meta['exp_total']:,.0f}</b>"])
        et = Table(exp_rows, colWidths=[1.7*inch, 3.9*inch, 1.2*inch])
        et.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F1F5F9")),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
            ("FONTSIZE",(0,0),(-1,-1),9),
            ("GRID",(0,0),(-1,-1),0.5,colors.HexColor("#E2E8F0")),
            ("ALIGN",(2,1),(2,-1),"RIGHT"),
        ]))
        story.append(et)
        story.append(Spacer(1, 14))

        # Total
        total_tbl = Table([[Paragraph("<b>ESTIMATED TOTAL</b>", h_style), Paragraph(f"<b>${meta['grand_total']:,.0f}</b>", h_style)]], colWidths=[5.3*inch, 1.5*inch])
        total_tbl.setStyle(TableStyle([("ALIGN",(1,0),(1,0),"RIGHT"), ("LINEABOVE",(0,0),(-1,0),1,colors.HexColor("#E2E8F0"))]))
        story.append(total_tbl)
        story.append(Spacer(1, 10))

        # Terms (no payment terms)
        story.append(Paragraph("Terms & Conditions", h_style))
        tc_lines = [
            "<b>Pricing & Quote Expiration:</b> Prices shown reflect an estimate of hours and expenses. Any additional hours or days will be billed at the rates shown above. Quote valid for 30 days.",
            f"<b>Customer Install Window:</b> No individual technician or engineer is assigned more than {meta['window']} onsite days per trip.",
            f"<b>Training:</b> Training days are calculated at 1 day per {TRAINING_MACHINES_PER_DAY} machines of the same model type. Training can be excluded per machine if not required (customer request only).",
            "<b>Machine-Specific Skills:</b> Each machine type requires technicians with specialized skills. Personnel are not shared across different machine types.",
            f"<b>Travel Days:</b> Expenses include {TRAVEL_DAYS_PER_PERSON} travel days (1 day travel-in + 1 day travel-out) in addition to onsite work days.",
        ]
        for s in tc_lines:
            story.append(Paragraph(s, normal))
            story.append(Spacer(1, 4))

        if self.data.requirements:
            story.append(Spacer(1, 8))
            story.append(Paragraph("Requirements & Assumptions", h_style))
            for s in self.data.requirements:
                story.append(Paragraph(s, small))
                story.append(Spacer(1, 2))

        doc.build(story)


def main():
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
