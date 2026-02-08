import sys, math
from PySide6.QtGui import QDesktopServices
from PySide6.QtCore import QUrl
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Tuple, Dict as TDict


def resolve_excel_path(expected_name: str = "Tech days and quote rates.xlsx") -> Path | None:
    """Find the default Excel workbook inside assets without prompting the user unless missing."""
    assets = resolve_assets_dir()

    exact = (assets / expected_name)
    try:
        if exact.exists():
            return exact.resolve()
    except Exception:
        pass

    # Fuzzy match (handles minor renames like '(1)' etc.)
    try:
        for f in assets.glob("*.xlsx"):
            n = f.name.lower()
            if "tech" in n and "quote" in n and "rate" in n:
                return f.resolve()
    except Exception:
        pass

    # If exactly one xlsx exists, use it
    try:
        xls = list(assets.glob("*.xlsx"))
        if len(xls) == 1:
            return xls[0].resolve()
    except Exception:
        pass

    return None

QUAL_MATRIX_FILENAME = "Machine Qualifications for PCP Quoting.xlsx"

def resolve_qual_matrix_path() -> Path | None:
    """Return path to the qualifications matrix if present in the bundled assets (or alongside the EXE)."""
    try:
        assets = resolve_assets_dir()
        p = assets / QUAL_MATRIX_FILENAME
        if p.exists():
            return p
    except Exception:
        pass
    # Fallback: look next to this script / exe
    try:
        root = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
        p2 = root / "assets" / QUAL_MATRIX_FILENAME
        if p2.exists():
            return p2
    except Exception:
        pass
    return None


class QualificationMatrix:
    """Loads a technician qualification matrix (T1/T2/T3) by model.

    Supported formats:

    **New format (preferred)**
      - Row 1: headers with model names starting in column C.
      - Column A: "Resource Type" (Technician/Engineer)
      - Column B: "Resource" (name/id like "Tech 1", "Eng 3")
      - Body cells: T1/T2/T3 (case-insensitive). Blank = not qualified.

    **Legacy format**
      - Row 1: headers with model names starting in column B.
      - Column A (rows 2+): technician names.
      - Body cells: T1/T2/T3 (case-insensitive). Blank = not qualified.

    Notes:
      - Only Technician rows are used for grouping logic.
      - Engineer rows (if present) are ignored.
    """

    def __init__(self, path: Path):
        self.path = Path(path)
        self.models: set[str] = set()
        self._ratings: dict[str, dict[str, int]] = {}  # tech -> model -> level (1-3)

        wb = load_workbook(self.path, data_only=True)
        ws = wb.active

        def norm(s) -> str:
            return str(s).strip() if s is not None else ""

        def parse_level(val) -> int:
            if val is None:
                return 0
            s = str(val).strip().upper()
            if not s:
                return 0
            if s.startswith("T") and s[1:].isdigit():
                n = int(s[1:])
                return n if n in (1, 2, 3) else 0
            return 0

        # Detect layout
        h1 = norm(ws.cell(row=1, column=1).value).lower()
        h2 = norm(ws.cell(row=1, column=2).value).lower()
        is_new = ("resource" in h1 and "type" in h1) and (h2 in ("resource", "name", "person"))

        model_start_col = 3 if is_new else 2
        name_col = 2 if is_new else 1
        type_col = 1 if is_new else None

        # Headers (models)
        headers: list[str] = []
        for c in range(model_start_col, ws.max_column + 1):
            v = ws.cell(row=1, column=c).value
            if v is None or str(v).strip() == "":
                continue
            headers.append(str(v).strip())
        self.models = set(headers)

        # Rows
        for r in range(2, ws.max_row + 1):
            if is_new:
                rtype = norm(ws.cell(row=r, column=type_col).value).lower()
                if not rtype:
                    continue
                if not rtype.startswith("tech"):  # ignore engineers and anything else
                    continue
                tech_cell = ws.cell(row=r, column=name_col).value
            else:
                tech_cell = ws.cell(row=r, column=name_col).value

            if tech_cell is None or str(tech_cell).strip() == "":
                continue
            tech_name = str(tech_cell).strip()

            row_map: dict[str, int] = {}
            for ci, model in enumerate(headers, start=model_start_col):
                lvl = parse_level(ws.cell(row=r, column=ci).value)
                if lvl:
                    row_map[model] = lvl

            # Store even if empty (keeps row count stable for debugging), but does not affect grouping.
            self._ratings[tech_name] = row_map

    @staticmethod
    def try_load() -> "QualificationMatrix | None":
        p = resolve_qual_matrix_path()
        if p and p.exists():
            try:
                return QualificationMatrix(p)
            except Exception:
                return None
        return None
class ModelInfo:
    item: str
    tech_install_days_per_machine: int   # install-only (no training baked in)
    eng_days_per_machine: int
    training_applicable: bool = True


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


@dataclass
class Assignment:
    model: str
    role: str  # "Technician" or "Engineer"
    person_num: int
    onsite_days: int
    cost: float


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

        headers = {
            str(ws.cell(1, c).value).strip(): c
            for c in range(1, ws.max_column + 1)
            if ws.cell(1, c).value is not None
        }

        def find_col(pred):
            for k, c in headers.items():
                if pred(k.lower()):
                    return c
            return None

        col_item = find_col(lambda s: s in ["item", "model", "machine", "machine type"])
        col_tech = find_col(lambda s: "technician" in s and "day" in s)
        col_eng = find_col(lambda s: ("engineer" in s and "day" in s) or ("field engineer" in s and "day" in s))
        col_train_app = find_col(lambda s: ("training required" in s))

        if col_item is None or col_tech is None or col_eng is None:
            raise ValueError("Model sheet columns not found. Expected: Item, Technician Days Required, Field Engineer Days Required.")

        def _as_bool(v, default=True):
            if v is None:
                return default
            if isinstance(v, bool):
                return v
            s = str(v).strip().lower()
            if s in ("true","t","yes","y","1"):
                return True
            if s in ("false","f","no","n","0"):
                return False
            return default

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
            train_app = _as_bool(ws.cell(r, col_train_app).value, default=True) if col_train_app is not None else True
            self.models[item] = ModelInfo(item=item, tech_install_days_per_machine=tech_i, eng_days_per_machine=eng_i, training_applicable=train_app)

        # Rates: Service Rates
        if "Service Rates" not in wb.sheetnames:
            raise ValueError("Missing sheet: 'Service Rates'")
        ws = wb["Service Rates"]

        header_row = None
        for r in range(1, 15):
            if str(ws.cell(r, 2).value).strip().lower() == "item" and str(ws.cell(r, 3).value).strip().lower() == "description":
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
        for rk, rv in self.rates.items():
            if k in rk:
                return float(rv["unit_price"]), str(rv["description"])
        raise KeyError(f"Rate not found for '{key}'")


    def get(self, *args):
        """Compatibility helper.

        Supports:
          - get(section, key)
          - get(section, key, default)
          - get(key)  (tries common rate keys)
        """
        if len(args) == 1:
            key = args[0]
            # try common places
            for sec in ("Service Rates", "Rates", "rates"):
                if sec in self.rates and key in self.rates[sec]:
                    return self.rates[sec][key]
            # flat scan
            for sec, d in self.rates.items():
                if key in d:
                    return d[key]
            raise KeyError(key)
        elif len(args) == 2:
            section, key = args
            if section in self.rates:
                return self.rates[section].get(key)
            raise KeyError(section)
        elif len(args) == 3:
            section, key, default = args
            if section in self.rates:
                return self.rates[section].get(key, default)
            return default
        else:
            raise TypeError("ExcelData.get() expects 1 to 3 arguments")

class MachineLine(QFrame):
    def __init__(self, models: List[str], training_applicable_map: Dict[str, bool], on_change, on_delete):
        super().__init__()
        self.on_change = on_change
        self.on_delete = on_delete
        self.training_applicable_map = training_applicable_map

        self.setObjectName("machineLine")
        row = QHBoxLayout(self)
        row.setContentsMargins(10, 10, 10, 10)
        row.setSpacing(10)

        self.cmb_model = QComboBox()
        self.cmb_model.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.cmb_model.addItem("— Select —")
        self.cmb_model.addItems(models)
        self.cmb_model.setCurrentIndex(-1)
        self.cmb_model.currentIndexChanged.connect(self._model_changed)

        self.spin_qty = QSpinBox()
        self.spin_qty.setRange(0, 999)
        self.spin_qty.setValue(1)
        self.spin_qty.valueChanged.connect(self._changed)

        self.chk_training = QCheckBox("Training Required")
        self.chk_training.setChecked(True)
        self.chk_training.stateChanged.connect(self._changed)
        self.chk_training.setToolTip("Uncheck only by customer request (training is normally included).")

        self.btn_delete = QPushButton("🗑")
        self.btn_delete.setFixedWidth(40)
        self.btn_delete.clicked.connect(self._delete)

        row.addWidget(QLabel("Machine Model"))
        row.addWidget(self.cmb_model, 2)
        row.addWidget(QLabel("Qty"))
        row.addWidget(self.spin_qty)
        row.addWidget(self.chk_training, 1)
        row.addWidget(self.btn_delete)

        self._model_changed()
    def _model_changed(self, *_):
        model = self.cmb_model.currentText().strip()
        if model == "— Select —":
            model = ""
        if not model:
            self.chk_training.hide()
            self.chk_training.setChecked(False)
        else:
            applicable = bool(self.training_applicable_map.get(model, True))
            if not applicable:
                self.chk_training.hide()
                self.chk_training.setChecked(False)
            else:
                self.chk_training.show()
                if not self.chk_training.isChecked():
                    self.chk_training.setChecked(True)
        self.on_change()

    def _changed(self, *_):

        self.on_change()

    def _delete(self):
        self.on_delete(self)

    def value(self) -> LineSelection:
        model = self.cmb_model.currentText().strip()
        if model == "— Select —":
            model = ""
        return LineSelection(
            model=model,
            qty=int(self.spin_qty.value()) if model else 0,
            training_required=(bool(self.chk_training.isChecked()) if self.chk_training.isVisible() else False)
        )


class Card(QFrame):
    def __init__(self, title: str, icon_text: str):
        super().__init__()
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.setObjectName("card")
        lay = QVBoxLayout(self)
        lay.setContentsMargins(14, 12, 14, 12)

        top = QHBoxLayout()
        ic = QLabel(icon_text)
        ic.setObjectName("cardIcon")
        ic.setFixedSize(QSize(34, 34))
        ic.setAlignment(Qt.AlignCenter)
        top.addWidget(ic)

        v = QVBoxLayout()
        self.lbl_title = QLabel(title)
        self.lbl_title.setObjectName("cardTitle")
        self.lbl_value = QLabel("—")
        self.lbl_value.setObjectName("cardValue")
        self.lbl_sub = QLabel("")
        self.lbl_sub.setObjectName("cardSub")
        self.lbl_sub.setWordWrap(True)
        v.addWidget(self.lbl_title)
        v.addWidget(self.lbl_value)
        top.addLayout(v, 1)

        lay.addLayout(top)
        lay.addWidget(self.lbl_sub)

    def set_value(self, value: str, sub: str = ""):
        self.lbl_value.setText(value)
        self.lbl_sub.setText(sub or "")


class Section(QFrame):
    def __init__(self, title: str, subtitle: str = "", icon_text: str = ""):
        super().__init__()
        self.setObjectName("section")
        lay = QVBoxLayout(self)
        lay.setContentsMargins(14, 14, 14, 14)
        lay.setSpacing(10)

        head = QHBoxLayout()
        if icon_text:
            ic = QLabel(icon_text)
            ic.setObjectName("sectionIcon")
            ic.setFixedSize(QSize(28, 28))
            ic.setAlignment(Qt.AlignCenter)
            head.addWidget(ic)

        title_box = QVBoxLayout()
        t = QLabel(title)
        t.setObjectName("sectionTitle")
        title_box.addWidget(t)
        if subtitle:
            s = QLabel(subtitle)
            s.setObjectName("sectionSub")
            s.setWordWrap(True)
            title_box.addWidget(s)
        head.addLayout(title_box, 1)
        lay.addLayout(head)

        self.content = QWidget()
        self.content_layout = QVBoxLayout(self.content)
        self.content_layout.setContentsMargins(0, 0, 0, 0)
        self.content_layout.setSpacing(10)
        lay.addWidget(self.content)


def money(x: float) -> str:
    return f"${x:,.0f}"


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(1920, 1200)

        self.data = ExcelData(DEFAULT_EXCEL)
        self.models_sorted = sorted(self.data.models.keys())
        self.training_app_map = {k: bool(v.training_applicable) for k, v in self.data.models.items()}
        self.qual = QualificationMatrix.try_load()
        self.lines: List[MachineLine] = []

        central_container = QWidget()
        root = QVBoxLayout(central_container)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # Outer scroll area enables whole-window scrolling in stacked (single-column) mode
        self.outer_scroll = QScrollArea()
        self.outer_scroll.setObjectName("outerScroll")
        self.outer_scroll.setWidgetResizable(True)
        self.outer_scroll.setFrameShape(QFrame.NoFrame)
        self.outer_scroll.setWidget(central_container)
        self.setCentralWidget(self.outer_scroll)


        header = QFrame()
        header.setObjectName("header")
        h = QHBoxLayout(header)
        h.setContentsMargins(14, 10, 14, 10)
        self.lbl_title = QLabel("🧾  " + APP_TITLE)
        self.lbl_title.setObjectName("appTitle")
        h.addWidget(self.lbl_title)
        h.addStretch(1)
        btn_load = QPushButton("Load Excel…")
        btn_load.setToolTip("Load a different Excel workbook (will replace the bundled rates/models for this session).")
        btn_load.clicked.connect(self.open_excel)
        h.addWidget(btn_load)

        btn_open_bundled = QPushButton("Open Bundled Excel")
        btn_open_bundled.setToolTip("Open the Excel workbook that was bundled into this EXE (for verification).")
        btn_open_bundled.clicked.connect(self.open_bundled_excel)
        h.addWidget(btn_open_bundled)
        root.addWidget(header)

        splitter = QSplitter(Qt.Horizontal)
        splitter.setChildrenCollapsible(False)
        self.splitter = splitter
        root.addWidget(splitter, 1)

        # LEFT
        left = QFrame()
        left.setObjectName("panel")
        left_l = QVBoxLayout(left)
        left_l.setContentsMargins(14, 14, 14, 14)
        left_l.setSpacing(12)

        t = QLabel("Machine Configuration")
        t.setObjectName("panelTitle")
        left_l.addWidget(t)

        left_l.addWidget(QLabel(
            "Add machines to estimate commissioning requirements.\\n"
            "Each machine type requires dedicated personnel (no sharing across types)."
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

        self.empty_hint = QLabel("No machines added.\\nClick “Add Machine” to begin.")
        self.empty_hint.setObjectName("emptyHint")
        self.empty_hint.setAlignment(Qt.AlignCenter)
        self.empty_hint.setMinimumHeight(120)
        self.lines_layout.addWidget(self.empty_hint)

        self.scroll.setWidget(container)
        left_l.addWidget(self.scroll, 1)

        btn_add = QPushButton("+  Add Machine")
        btn_add.setObjectName("addMachine")
        btn_add.clicked.connect(self.add_line)
        left_l.addWidget(btn_add)

        note = QLabel("Note: Unchecking “Training Required” should only be done by customer request.")
        note.setObjectName("note")
        note.setWordWrap(True)
        left_l.addWidget(note)

        splitter.addWidget(left)

        # RIGHT (scrollable)
        right_wrap = QWidget()
        self.right_wrap = right_wrap
        right_layout = QVBoxLayout(right_wrap)
        right_layout.setContentsMargins(0, 0, 0, 0)

        right_scroll = QScrollArea()
        right_scroll.setWidgetResizable(True)
        self.right_scroll = right_scroll
        right_layout.addWidget(right_scroll)

        right = QWidget()
        self.right_content = right
        right_scroll.setWidget(right)
        right_l = QVBoxLayout(right)
        right_l.setContentsMargins(14, 14, 14, 14)
        right_l.setSpacing(12)

        cards = QHBoxLayout()
        self.card_tech = Card("Technicians", "🧰")
        self.card_eng = Card("Engineers", "🧑‍💻")
        self.card_window = Card("Max Onsite", "⏱")
        self.card_total = Card("Total Cost", "💲")
        cards.addWidget(self.card_tech, 1)
        cards.addWidget(self.card_eng, 1)
        cards.addWidget(self.card_window, 1)
        cards.addWidget(self.card_total, 1)
        right_l.addLayout(cards)

        self.alert = QLabel("")
        self.alert.setObjectName("alert")
        self.alert.setWordWrap(True)
        self.alert.hide()
        right_l.addWidget(self.alert)

        self.tbl_breakdown = self.make_table(["Model", "Qty", "Tech Days", "Eng Days", "Technicians", "Engineers"])
        self.tbl_assign = self.make_table(["Machine Type", "Role", "Person #", "Assigned Days", "Cost"])
        self.tbl_labor = self.make_table(["Role", "Daily Rate", "Total Days", "Personnel", "Total Cost"])
        self.tbl_exp = self.make_table(["Expense", "Details", "Amount"])
        self.tbl_exp.setMinimumHeight(0)

        sec_breakdown = Section("Machine Breakdown", "Days and personnel required per machine model", "🧩")
        sec_breakdown.content_layout.addWidget(self.tbl_breakdown)

        sec_assign = Section("Personnel Assignments", "Each machine type has dedicated personnel.", "👥")
        sec_assign.content_layout.addWidget(self.tbl_assign)

        sec_labor = Section("Labor Costs", "Labor costs by role at daily rates (8 hours/day).", "🛠")
        sec_labor.content_layout.addWidget(self.tbl_labor)

        # Workload bar chart (bonus visual)
        self.chart = QChart()
        self.chart_view = QChartView(self.chart)
        self.chart_view.setRenderHint(QPainter.Antialiasing)
        self.chart_view.setMinimumHeight(300)
        sec_chart = Section("Workload", "Days onsite per person (T=Tech, E=Engineer).", "📊")
        sec_chart.content_layout.addWidget(self.chart_view)

        # Left side: put chart under Machine Configuration so the right-side widgets stay readable
        left_l.addWidget(sec_chart)

        right_l.addWidget(sec_breakdown)
        right_l.addWidget(sec_assign)
        right_l.addWidget(sec_labor)

        sec_exp = Section("Estimated Expenses", "", "🧳")
        self.lbl_exp_hdr = QLabel("")
        self.lbl_exp_hdr.setObjectName("sectionSub")
        self.lbl_exp_hdr.setWordWrap(True)
        sec_exp.content_layout.addWidget(self.lbl_exp_hdr)
        sec_exp.content_layout.addWidget(self.tbl_exp)
        right_l.addWidget(sec_exp)

        bottom = QFrame()
        bottom.setObjectName("totalBar")
        bl = QHBoxLayout(bottom)
        bl.setContentsMargins(14, 12, 14, 12)
        self.lbl_total = QLabel("Estimated Total")
        self.lbl_total.setObjectName("totalLabel")
        self.lbl_total_val = QLabel("—")
        self.lbl_total_val.setObjectName("totalValue")
        bl.addWidget(self.lbl_total)
        bl.addStretch(1)
        bl.addWidget(self.lbl_total_val)
        self.btn_print = QPushButton("Print Quote…")
        self.btn_print.clicked.connect(self.print_quote_preview)
        self.btn_print.setEnabled(False)
        bl.addWidget(self.btn_print)
        right_l.addWidget(bottom)

        splitter.addWidget(right_wrap)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 2)

        self.apply_theme()
        self.reset_views()

        # Responsive scaling baseline (designed for 1920x1200)
        self._base_font_pt = float(self.font().pointSizeF() or 10.0)
        self._apply_scale()

        # Responsive layout: two-column w/ right scroll on large screens; single stacked w/ full-window scroll on small screens
        self._stack_threshold = 1280
        self._is_stacked = False
        self._apply_responsive_layout()

    def make_table(self, headers: List[str]) -> QTableWidget:
        tbl = QTableWidget(0, len(headers))
        tbl.setHorizontalHeaderLabels(headers)
        tbl.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        tbl.verticalHeader().setVisible(False)
        tbl.setEditTriggers(QAbstractItemView.NoEditTriggers)
        tbl.setSelectionBehavior(QAbstractItemView.SelectRows)
        tbl.setSelectionMode(QAbstractItemView.SingleSelection)
        tbl.setAlternatingRowColors(True)
        tbl.setObjectName("table")
        tbl.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        tbl.setMinimumHeight(120)
        return tbl

    def apply_theme(self):
        # Pearson-ish palette (navy + orange + neutral)
        blue = "#4B4F54"   # charcoal gray (logo text)
        gold = "#F05A28"   # Pearson orange
        neutral = "#6D6E71"
        red = "#D6453D"    # Pearson red accent
        css = """
        QFrame#header { background: __BLUE__; color: white; border: none; }
        QLabel#appTitle { color: white; font-size: 20px; font-weight: 800; }
        QFrame#panel { background: white; border: 1px solid #E6E8EB; border-radius: 14px; }
        QLabel#panelTitle { font-size: 16px; font-weight: 800; color: #0F172A; }
        QFrame#softBox { background: #FFF7EA; border: 1px solid #F0D8A8; border-radius: 12px; }
        QLabel#note { color: #334155; font-size: 12px; }
        QLabel#emptyHint { color: __NEUTRAL__; background: #F8FAFC; border: 1px dashed #CBD5E1; border-radius: 12px; }
        QPushButton#primary {
            background: __GOLD__; border: 0px; color: #0B1B2A;
            padding: 10px 12px; border-radius: 10px; font-weight: 800;
        }
        QPushButton#addMachine {
            background: #bebebe; border: 0px; color: #0B1B2A;
            padding: 10px 12px; border-radius: 10px; font-weight: 800;
        }
        QPushButton#addMachine:hover { background: #D6D9DD; }
        QPushButton#addMachine:pressed { background: #CBD5E1; }
        QPushButton {
            padding: 8px 10px; border-radius: 10px;
            border: 1px solid #D6D9DD; background: #F8FAFC;
        }
        QPushButton:disabled { color: #94A3B8; background: #F1F5F9; }
        QFrame#card { background: #FFFFFF; border: 1px solid #E6E8EB; border-radius: 14px; }
        QLabel#cardIcon { background: #EEF2F7; border-radius: 10px; font-size: 16px; }
        QLabel#cardTitle { font-size: 12px; color: __NEUTRAL__; font-weight: 700; }
        QLabel#cardValue { font-size: 24px; color: #0F172A; font-weight: 900; }
        QLabel#cardSub { font-size: 12px; color: __NEUTRAL__; }
        QFrame#section { background: #FFFFFF; border: 1px solid #E6E8EB; border-radius: 14px; }
        QLabel#sectionIcon { background: #EEF2F7; border-radius: 10px; font-size: 14px; }
        QLabel#sectionTitle { font-size: 15px; font-weight: 900; color: #0F172A; }
        QLabel#sectionSub { font-size: 12px; color: __NEUTRAL__; }
        QTableWidget#table {
            background: #FFFFFF;
            border: 1px solid #EEF0F2;
            border-radius: 12px;
            gridline-color: #E2E8F0;
            selection-background-color: #DBEAFE;
        }
        QHeaderView::section {
            background: #343551;
            color: white;
            padding: 8px;
            border: 0px;
            border-bottom: 1px solid #E2E8F0;
            font-weight: 800;
        }
        QFrame#totalBar { background: #F8FAFC; border: 1px solid #E6E8EB; border-radius: 14px; }
        QLabel#totalLabel { font-size: 13px; font-weight: 800; color: #0F172A; }
        QLabel#totalValue { font-size: 22px; font-weight: 900; color: __BLUE__; }
        QLabel#alert {
            background: #FEF2F2;
            border: 1px solid #FCA5A5;
            color: #7F1D1D;
            padding: 10px;
            border-radius: 12px;
        }
        QFrame#machineLine { background: #FFFFFF; border: 1px solid #E6E8EB; border-radius: 12px; }
        """

        css = css.replace("__BLUE__", blue).replace("__GOLD__", gold).replace("__NEUTRAL__", neutral).replace("__RED__", red)
        self.setStyleSheet(css)

    def reset_views(self):
        self.card_tech.set_value("0", "0 total days")
        self.card_eng.set_value("0", "0 total days")
        self.card_window.set_value(f"{self.spin_window.value()}", "")
        self.card_total.set_value("—", "labor + expenses")
        self.lbl_total_val.setText("—")
        for tbl in [self.tbl_breakdown, self.tbl_assign, self.tbl_labor, self.tbl_exp]:
            tbl.setRowCount(0)
        self.lbl_exp_hdr.setText("")
        self.btn_print.setEnabled(False)
        self.alert.hide()
        self.alert.setText("")
        if hasattr(self, 'chart'):
            try:
                self.chart.removeAllSeries()
                self.chart.setTitle('Workload (onsite days)')
            except Exception:
                pass

    def add_line(self):
        if self.empty_hint is not None:
            self.empty_hint.hide()
        ln = MachineLine(self.models_sorted, self.training_app_map, on_change=self.recalc, on_delete=self.delete_line)
        self.lines.append(ln)
        self.lines_layout.addWidget(ln)
        self.recalc()

    def delete_line(self, ln: MachineLine):
        self.lines.remove(ln)
        ln.setParent(None)
        ln.deleteLater()
        if len(self.lines) == 0:
            self.empty_hint.show()
            self.reset_views()
        else:
            self.recalc()

    def open_bundled_excel(self):
        """Open the bundled Excel workbook in the user's default spreadsheet app."""
        try:
            p = resolve_excel_path()
            if not p or not p.exists():
                QMessageBox.warning(self, "Bundled Excel not found",
                                    "The bundled Excel file could not be found inside the app assets.")
                return
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(p)))
        except Exception as e:
            QMessageBox.critical(self, "Open error", str(e))


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

    def calc(self):
        selections = [ln.value() for ln in self.lines if ln.value().model and ln.value().qty > 0]

        window = int(self.spin_window.value())
        if window <= 0:
            window = 1

        # Rates come from the bundled Excel "Service Rates" tab.
        # Older drafts used editable spin-box controls (spin_tech_rate/spin_eng_rate
        # and travel-in/out). Those controls are not part of the current UI, so
        # we read rates directly from the workbook and default travel-in/out to 1 day
        # each when travel is required.
        tech_day_rate = float(self.data.get_rate("Tech. Regular Time", 155.0))
        eng_day_rate = float(self.data.get_rate("Eng. Regular Time", 206.0))

        travel_in = 1
        travel_out = 1

        qual = getattr(self, "qual", None)

        def training_days_for(sel: Selection) -> int:
            if not self.training_app_map.get(sel.model, False):
                return 0
            if not sel.training_required:
                return 0
            return int(math.ceil(sel.qty / 3.0))

        # -------------------------------------------------
        # TECHNICIAN ALLOCATION (with qualification grouping)
        # -------------------------------------------------
        model_to_tech_hc: dict[str, int] = {}
        tech_days_per_person: dict[int, int] = {}
        tech_next_id = 1
        assignments: list[Assignment] = []

        # Categorize selections
        rpc_sels = [s for s in selections if s.model.upper().startswith("RPC")]
        eng_required_sels = [s for s in selections if (not s.model.upper().startswith("RPC")) and float(self.data.models[s.model].engineer_days_per_machine) > 0.0]
        tech_only_sels = [s for s in selections if (not s.model.upper().startswith("RPC")) and float(self.data.models[s.model].engineer_days_per_machine) <= 0.0]

        matrix_sels = [s for s in tech_only_sels if qual and (s.model in getattr(qual, "models", set()))]
        universal_sels = [s for s in tech_only_sels if (not qual) or (s.model not in getattr(qual, "models", set()))]

        def tech_workload(sel: Selection) -> int:
            d = self.data.models[sel.model]
            base = int(sel.qty * float(d.technician_days_per_machine))
            return int(base + (training_days_for(sel) if base > 0 else 0))

        # Greedy grouping among matrix-covered tech-only models
        groups: list[list[Selection]] = []
        if qual and matrix_sels:
            for s in sorted(matrix_sels, key=lambda x: -tech_workload(x)):
                placed = False
                for g in groups:
                    if qual.group_ok([x.model for x in g] + [s.model]):
                        g.append(s)
                        placed = True
                        break
                if not placed:
                    groups.append([s])
        else:
            groups = [[s] for s in matrix_sels]

        # If we only have "universal" tech-only items, keep them on a single existing crew
        if not groups and universal_sels:
            groups = [[]]

        # Distribute universal tech-only items across existing groups to avoid spawning new crews
        group_loads = [sum(tech_workload(s) for s in g) for g in groups]
        for s in sorted(universal_sels, key=lambda x: -tech_workload(x)):
            if not groups:
                groups = [[]]
                group_loads = [0]
            idx = min(range(len(groups)), key=lambda i: group_loads[i])
            groups[idx].append(s)
            group_loads[idx] += tech_workload(s)

        def allocate_group(label: str, sels_in_group: list[Selection]):
            nonlocal tech_next_id
            total_days = sum(tech_workload(s) for s in sels_in_group)
            if total_days <= 0:
                return
            headcount = max(1, int(math.ceil(total_days / float(window))))
            days_list = balanced_allocate(int(total_days), headcount)
            for i, d in enumerate(days_list, start=1):
                tech_days_per_person[tech_next_id] = int(d)
                assignments.append(Assignment(label, "Technician", tech_next_id, int(d), float(d) * tech_day_rate))
                tech_next_id += 1
            for s in sels_in_group:
                model_to_tech_hc[s.model] = headcount

        # Allocate grouped tech-only crews
        for gi, g in enumerate(groups):
            if not g:
                continue
            label_models = ", ".join(sorted({s.model for s in g}))
            allocate_group(f"Tech Crew {gi+1} ({label_models})", g)

        # Allocate RPC tech crew as its own group (robot cell skillset)
        if rpc_sels:
            label_models = ", ".join(sorted({s.model for s in rpc_sels}))
            allocate_group(f"Robot Cell Crew (Tech) ({label_models})", rpc_sels)

        # Allocate technician workload for engineer-required non-RPC models (separate)
        for s in eng_required_sels:
            if s.model in model_to_tech_hc:
                continue
            allocate_group(f"{s.model} (Tech)", [s])

        # -------------------------------------------------
        # ENGINEER ALLOCATION (RPC-C and RPC-DF share crew)
        # -------------------------------------------------
        model_to_eng_hc: dict[str, int] = {}
        eng_days_per_person: dict[int, int] = {}
        eng_next_id = 1

        def eng_workload(sel: Selection) -> int:
            d = self.data.models[sel.model]
            base = int(sel.qty * float(d.engineer_days_per_machine))
            if base <= 0:
                return 0
            return int(base + training_days_for(sel))

        def allocate_eng_group(label: str, sels_in_group: list[Selection]):
            nonlocal eng_next_id
            total_days = sum(eng_workload(s) for s in sels_in_group)
            if total_days <= 0:
                return
            headcount = max(1, int(math.ceil(total_days / float(window))))
            days_list = balanced_allocate(int(total_days), headcount)
            for i, d in enumerate(days_list, start=1):
                eng_days_per_person[eng_next_id] = int(d)
                assignments.append(Assignment(label, "Engineer", eng_next_id, int(d), float(d) * eng_day_rate))
                eng_next_id += 1
            for s in sels_in_group:
                model_to_eng_hc[s.model] = headcount

        if rpc_sels:
            label_models = ", ".join(sorted({s.model for s in rpc_sels}))
            allocate_eng_group(f"Robot Cell Crew (Eng) ({label_models})", rpc_sels)

        for s in eng_required_sels:
            if s.model in model_to_eng_hc:
                allocate_eng_group(f"{s.model} (Eng)", [s])  # still make sure we count it
            else:
                allocate_eng_group(f"{s.model} (Eng)", [s])

        # -------------------------------------------------
        # Build machine breakdown rows (per selection/model)
        # -------------------------------------------------
        machine_rows = []
        for s in selections:
            d = self.data.models[s.model]
            base_training = training_days_for(s)

            base_tech = int(s.qty * float(d.technician_days_per_machine))
            base_eng = int(s.qty * float(d.engineer_days_per_machine))

            tech_total = base_tech + (base_training if base_tech > 0 else 0)
            eng_total = base_eng + (base_training if base_eng > 0 else 0)

            if self.training_app_map.get(s.model, False):
                if s.training_required:
                    s_train = f"(incl. {base_training} Train)"
                else:
                    s_train = "(training excluded)"
            else:
                s_train = ""

            tech_hc = int(model_to_tech_hc.get(s.model, math.ceil(tech_total / float(window)) if tech_total else 0))
            eng_hc = int(model_to_eng_hc.get(s.model, math.ceil(eng_total / float(window)) if eng_total else 0))

            machine_rows.append((
                s.model,
                s.qty,
                f"{tech_total} {s_train}".strip(),
                tech_hc,
                (f"{eng_total} {s_train}".strip() if eng_total else ""),
                eng_hc
            ))

        tech_all = list(tech_days_per_person.values())
        eng_all = list(eng_days_per_person.values())

        tech = RoleTotals(len(tech_all), sum(tech_all), tech_all, tech_day_rate, sum(float(d) * tech_day_rate for d in tech_all))
        eng = RoleTotals(len(eng_all), sum(eng_all), eng_all, eng_day_rate, sum(float(d) * eng_day_rate for d in eng_all))

        # -------------------------------------------------
        # Expenses
        # -------------------------------------------------
        trip_days_by_person = {a.person_id: (travel_in + travel_out) for a in assignments}
        expenses_rows, expenses_total = self.expenses_calc(len(assignments), travel_in, travel_out)

        labor_lines = [(a.role, a.person_label(), a.days, a.cost) for a in assignments]
        exp_lines = [(a.role, a.person_label(), trip_days_by_person.get(a.person_id, 0)) for a in assignments]

        meta = {
            "machine_rows": machine_rows,
            "assignments": exp_lines,
            "labor_lines": labor_lines,
            "expenses_rows": expenses_rows,
            "expenses_total": float(expenses_total),
        }

        return tech, eng, exp_lines, meta

    def recalc(self):
        if len(self.lines) == 0:
            self.reset_views()
            return
        try:
            tech, eng, exp_lines, meta = self.calc()
            self.alert.hide()

            self.card_tech.set_value(str(tech.headcount), f"{tech.total_onsite_days} total days")
            self.card_eng.set_value(str(eng.headcount), f"{eng.total_onsite_days} total days")
            self.card_window.set_value(f"{meta['max_onsite']} days", f"install window {meta['window']} days")
            self.card_total.set_value(money(meta["grand_total"]), "labor + expenses")

            self.update_workload_chart(tech, eng)
            self.lbl_total_val.setText(money(meta["grand_total"]))

            rows = meta["machine_rows"]
            self.tbl_breakdown.setRowCount(len(rows))
            for r_i, r in enumerate(rows):
                # Training display rules:
                # - If training is not applicable for this model, hide all training UI/labels.
                # - If applicable but user unchecked training, show “(training excluded)”.
                if not r.get("training_applicable", True):
                    tech_disp = str(r["tech_total"])
                else:
                    if r.get("training_required", True):
                        tech_disp = f"{r['tech_total']} (incl. {r['training_days']} Train)" if r.get("training_days", 0) > 0 else str(r["tech_total"])
                    else:
                        tech_disp = f"{r['tech_total']} (training excluded)"

                if r["eng_total"] == 0:
                    eng_disp = "—"
                elif not r.get("training_applicable", True):
                    eng_disp = str(r["eng_total"])
                else:
                    eng_tp = r.get("eng_training_potential", 0)
                    eng_td = r.get("eng_training_days", 0)
                    if r.get("training_required", True):
                        eng_disp = f"{r['eng_total']} (incl. {eng_td} Train)" if (eng_tp > 0 and eng_td > 0) else str(r["eng_total"])
                    else:
                        eng_disp = f"{r['eng_total']} (training excluded)" if eng_tp > 0 else str(r["eng_total"])

                vals = [r["model"], str(r["qty"]), tech_disp, eng_disp,
                        "—" if r["tech_headcount"] == 0 else str(r["tech_headcount"]),
                        "—" if r["eng_headcount"] == 0 else str(r["eng_headcount"])]
                for c, v in enumerate(vals):
                    it = QTableWidgetItem(v)
                    if c in [1, 4, 5]:
                        it.setTextAlignment(Qt.AlignCenter)
                    if c == 2 and r["training_required"]:
                        it.setForeground(Qt.darkYellow)
                    self.tbl_breakdown.setItem(r_i, c, it)

            assigns: List[Assignment] = meta["assignments"]
            self.tbl_assign.setRowCount(len(assigns))
            for i, a in enumerate(assigns):
                vals = [a.model, a.role, str(a.person_num), str(a.onsite_days), money(a.cost)]
                for c, v in enumerate(vals):
                    it = QTableWidgetItem(v)
                    if c in [2, 3]:
                        it.setTextAlignment(Qt.AlignCenter)
                    if c == 4:
                        it.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                    self.tbl_assign.setItem(i, c, it)

            self.tbl_labor.setRowCount(3)
            labor_rows = [
                ("Technician", money(tech.day_rate) + "/day", str(tech.total_onsite_days), str(tech.headcount), money(tech.labor_cost)),
                ("Engineer", money(eng.day_rate) + "/day", str(eng.total_onsite_days), str(eng.headcount), money(eng.labor_cost)),
            ]
            for r_i, row in enumerate(labor_rows):
                for c, v in enumerate(row):
                    it = QTableWidgetItem(v)
                    if c in [2, 3]:
                        it.setTextAlignment(Qt.AlignCenter)
                    if c == 4:
                        it.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                    self.tbl_labor.setItem(r_i, c, it)


            # Subtotal row
            labor_subtotal = tech.labor_cost + eng.labor_cost
            self.tbl_labor.setItem(2, 0, QTableWidgetItem("Subtotal"))
            self.tbl_labor.setItem(2, 1, QTableWidgetItem(""))
            self.tbl_labor.setItem(2, 2, QTableWidgetItem(""))
            self.tbl_labor.setItem(2, 3, QTableWidgetItem(""))
            it = QTableWidgetItem(money(labor_subtotal))
            it.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.tbl_labor.setItem(2, 4, it)

            self.lbl_exp_hdr.setText(
                f"Expenses are calculated using person-days, including {TRAVEL_DAYS_PER_PERSON} travel days per person."
            )
            self.tbl_exp.setRowCount(len(exp_lines) + 1)
            for i, l in enumerate(exp_lines):
                vals = [l.description, l.details, money(l.extended)]
                for c, v in enumerate(vals):
                    it = QTableWidgetItem(v)
                    if c == 2:
                        it.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                    self.tbl_exp.setItem(i, c, it)

            sub_row = len(exp_lines)
            self.tbl_exp.setItem(sub_row, 0, QTableWidgetItem("Expenses Subtotal"))
            self.tbl_exp.setItem(sub_row, 1, QTableWidgetItem("—"))
            it = QTableWidgetItem(money(meta["exp_total"]))
            it.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.tbl_exp.setItem(sub_row, 2, it)

            self.btn_print.setEnabled(True)

        except Exception as e:
            self.reset_views()
            self.alert.setText(str(e))
            self.alert.show()

    def build_quote_html(self, tech: RoleTotals, eng: RoleTotals, exp_lines: List[ExpenseLine], meta: TDict[str, object]) -> str:
        from datetime import date, timedelta
        today = date.today()
        validity = today + timedelta(days=30)
        date_str = f"{today:%B} {today.day}, {today:%Y}"
        valid_str = f"{validity:%B} {validity.day}, {validity:%Y}"

        logo_html = ""
        if LOGO_PATH.exists():
            try:
                b = LOGO_PATH.read_bytes()
                b64 = base64.b64encode(b).decode("ascii")
                logo_html = f'<img src="data:image/png;base64,{b64}" height="36" style="height:36px;" />'
            except Exception:
                logo_html = ""

        mr = []
        for r in meta["machine_rows"]:
            if not r.get("training_applicable", True):
                tech_disp = str(r["tech_total"])
            else:
                if r.get("training_required", True):
                    tech_disp = f"{r['tech_total']} (incl. {r['training_days']} Train)" if r.get("training_days", 0) > 0 else str(r["tech_total"])
                else:
                    tech_disp = f"{r['tech_total']} (training excluded)"

            if r["eng_total"] == 0:
                eng_disp = "—"
            elif not r.get("training_applicable", True):
                eng_disp = str(r["eng_total"])
            else:
                eng_tp = r.get("eng_training_potential", 0)
                eng_td = r.get("eng_training_days", 0)
                if r.get("training_required", True):
                    eng_disp = f"{r['eng_total']} (incl. {eng_td} Train)" if (eng_tp > 0 and eng_td > 0) else str(r["eng_total"])
                else:
                    eng_disp = f"{r['eng_total']} (training excluded)" if eng_tp > 0 else str(r["eng_total"])

            mr.append(f"""<tr>
                <td>{r['model']}</td>
                <td style="text-align:center;">{r['qty']}</td>
                <td>{tech_disp}</td>
                <td style="text-align:center;">{eng_disp}</td>
                <td style="text-align:center;">{r['tech_headcount'] if r['tech_headcount'] else "—"}</td>
                <td style="text-align:center;">{r['eng_headcount'] if r['eng_headcount'] else "—"}</td>
            </tr>""")

        exp_rows = []
        for l in exp_lines:
            exp_rows.append(f"""<tr>
                <td>{l.description}</td>
                <td>{l.details}</td>
                <td style="text-align:right;">{money(l.extended)}</td>
            </tr>""")

        labor_sub = tech.labor_cost + eng.labor_cost

        req_html = ""
        if self.data.requirements:
            li = "".join([f"<li>{x}</li>" for x in self.data.requirements])
            req_html = f"<h3>Requirements & Assumptions</h3><ul>{li}</ul>"

        html = f"""<html><head><meta charset="utf-8" />
        <style>
            body {{ font-family: Arial, Helvetica, sans-serif; font-size: 10pt; color: #0F172A; }}
            .topbar {{ display:flex; align-items:flex-start; justify-content:space-between; border-bottom: 3px solid #F05A28; padding-bottom: 10px; margin-bottom: 14px; }}
            .logo {{ text-align:right; }}
            .title {{ font-size: 18pt; font-weight: 800; color: #4c4b4c; margin: 0; }}
            .subtitle {{ margin: 4px 0 0 0; color: #6D6E71; }}
            .grid {{ width: 100%; border-collapse: collapse; margin-top: 10px; }}
            .grid th {{ background: #343551; color: white; text-align: left; padding: 8px; border-bottom: 1px solid #E2E8F0; }}
            .grid td {{ padding: 8px; border-bottom: 1px solid #E2E8F0; }}
            .box {{ border: 1px solid #E6E8EB; border-radius: 10px; padding: 10px; background: rgba(103,144,160,0.18); }}
            .two {{ display: table; width: 100%; }}
            .two > div {{ display: table-cell; width: 50%; vertical-align: top; padding-right: 10px; }}
            h3 {{ color: #4c4b4c; margin: 18px 0 8px 0; }}
            .right {{ text-align: right; }}
            .muted {{ color: #6D6E71; }}
            .total {{ font-size: 16pt; font-weight: 900; color: #4c4b4c; }}
        </style></head><body>
            <div class="topbar">
                <div>
                    <p class="title">Commissioning Budget Quote</p>
                    <p class="subtitle muted">Service Estimate</p>
                </div>
                <div class="logo">{logo_html}</div>
            </div>

            <div class="two">
                <div class="box">
                    <b>DATE</b><br/>{date_str}<br/><br/>
                    <b>TOTAL PERSONNEL</b><br/>{tech.headcount + eng.headcount} ({tech.headcount} Tech, {eng.headcount} Eng)
                </div>
                <div class="box">
                    <b>QUOTE VALIDITY</b><br/>{valid_str}<br/><br/>
                    <b>ESTIMATED DURATION</b><br/>{meta["max_onsite"]} days onsite + {TRAVEL_DAYS_PER_PERSON} travel days
                </div>
            </div>
            <div class="section-spacer"></div>

            <h3>Machine Breakdown</h3>
            <table class="grid">
                <tr><th>Model</th><th style="text-align:center;">Qty</th><th>Tech Days</th><th style="text-align:center;">Eng Days</th>
                    <th style="text-align:center;">Technicians</th><th style="text-align:center;">Engineers</th></tr>
                {''.join(mr)}
            </table>

            <h3>Labor Costs</h3>
            <table class="grid">
                <tr><th>Item</th><th class="right">Extended</th></tr>
                <tr><td>Tech. Regular Time ({tech.total_onsite_days} days × {money(tech.day_rate)}/day)</td><td class="right">{money(tech.labor_cost)}</td></tr>
                <tr><td>Eng. Regular Time ({eng.total_onsite_days} days × {money(eng.day_rate)}/day)</td><td class="right">{money(eng.labor_cost)}</td></tr>
                <tr><td><b>Labor Subtotal</b></td><td class="right"><b>{money(labor_sub)}</b></td></tr>
            </table>

            <h3>Estimated Expenses</h3>
            <div class="muted">Includes {int(meta["total_trip_days"])} total trip day(s) across personnel (onsite + travel days).</div>
            <table class="grid">
                <tr><th>Expense</th><th>Details</th><th class="right">Amount</th></tr>
                {''.join(exp_rows)}
                <tr><td><b>Expenses Subtotal</b></td><td>—</td><td class="right"><b>{money(meta["exp_total"])}</b></td></tr>
            </table>

            <h3>Estimated Total</h3>
            <div class="box">
                <span class="total">{money(meta["grand_total"])}</span><br/>
                <span class="muted">Labor ({money(labor_sub)}) + Expenses ({money(meta["exp_total"])})</span>
            </div>

            <h3>Terms & Conditions</h3>
            <ul>
                <li><b>Pricing & Quote Expiration:</b> Prices shown reflect an estimate of days and expenses. Any additional time will be billed at the rates shown. Quote valid for 30 days.</li>
                <li><b>Customer Install Window:</b> No individual technician or engineer is assigned more than {meta["window"]} onsite days per trip.</li>
                <li><b>Training:</b> Training days are calculated at 1 day per {TRAINING_MACHINES_PER_DAY} machines of the same model type. Training can be excluded per machine if not required (customer request only).</li>
                <li><b>Machine-Specific Skills:</b> Each machine type requires technicians with specialized skills. Personnel are not shared across different machine types.</li>
                <li><b>Travel Days:</b> Expenses include {TRAVEL_DAYS_PER_PERSON} travel days (1 day travel-in + 1 day travel-out) in addition to onsite work days.</li>
            </ul>
            {req_html}
        </body></html>"""
        return html

    def print_quote_preview(self):
        try:
            tech, eng, exp_lines, meta = self.calc()
        except Exception as e:
            QMessageBox.critical(self, "Cannot print", str(e))
            return

        try:
            html = self.build_quote_html(tech, eng, exp_lines, meta)
            doc = QTextDocument()
            doc.setHtml(html)

            printer = QPrinter(QPrinter.HighResolution)
            # Letter page (8.5x11)
            try:
                printer.setPageSize(QPageSize(QPageSize.Letter))
            except Exception:
                pass

            preview = QPrintPreviewDialog(printer, self)
            preview.setWindowTitle("Print Preview - Commissioning Budget Quote")
            preview.setWindowModality(Qt.ApplicationModal)
            preview.resize(1100, 800)
            preview.paintRequested.connect(lambda p: doc.print_(p))

            # Show the dialog reliably
            preview.exec()
        except Exception as e:
            QMessageBox.critical(self, "Print error", str(e))
            return

    
    def _update_right_scroll_height_if_stacked(self):
        """When stacked, expand the right scroll area to its content so the OUTER scroll handles scrolling."""
        if not getattr(self, "_is_stacked", False):
            return
        try:
            if hasattr(self, "right_scroll") and hasattr(self, "right_content"):
                h = int(self.right_content.sizeHint().height()) + 80
                self.right_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
                self.right_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
                self.right_scroll.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
                self.right_scroll.setMinimumHeight(h)
                self.right_scroll.setMaximumHeight(h)
        except Exception:
            pass

    def _apply_responsive_layout(self):
        """Large screens: 2 columns w/ right-side scroll. Small screens: single stacked w/ full-window scroll."""
        if not hasattr(self, "splitter") or not hasattr(self, "outer_scroll") or not hasattr(self, "right_scroll"):
            return

        w = int(self.width())
        stacked = w < getattr(self, "_stack_threshold", 1280)

        if stacked and not getattr(self, "_is_stacked", False):
            self._is_stacked = True
            self.splitter.setOrientation(Qt.Vertical)
            # Enable whole-window scrolling; disable inner right scrolling
            self.outer_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
            self.outer_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            self._update_right_scroll_height_if_stacked()
            # Give left pane enough room; right pane will follow under it
            try:
                self.splitter.setSizes([650, 1000])
            except Exception:
                pass

            # In stacked (single-column) mode, make the machine configuration area taller so
            # multiple machine lines are visible without feeling cramped.
            try:
                if hasattr(self, "scroll"):
                    self.scroll.setMinimumHeight(320)
                    self.scroll.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
            except Exception:
                pass

        elif (not stacked) and getattr(self, "_is_stacked", False):
            self._is_stacked = False
            self.splitter.setOrientation(Qt.Horizontal)
            # Disable whole-window scrolling; allow right column to scroll
            self.outer_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            self.outer_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            self.right_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
            self.right_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            self.right_scroll.setMinimumHeight(0)
            self.right_scroll.setMaximumHeight(16777215)
            self.right_scroll.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
            try:
                self.splitter.setSizes([520, 1040])
            except Exception:
                pass

            # Restore default sizing for wide (two-column) mode.
            try:
                if hasattr(self, "scroll"):
                    self.scroll.setMinimumHeight(0)
            except Exception:
                pass

        elif stacked:
            # Still stacked; keep heights updated as content changes
            self._update_right_scroll_height_if_stacked()
            try:
                if hasattr(self, "scroll"):
                    self.scroll.setMinimumHeight(320)
            except Exception:
                pass

    def _apply_scale(self):
        # Scale UI typography modestly with window size; keep within sensible bounds.
        w = max(self.width(), 1)
        h = max(self.height(), 1)
        # Use width as primary driver; clamp to avoid extremes.
        scale = w / 1920.0
        scale = 0.85 if scale < 0.85 else (1.25 if scale > 1.25 else scale)
        pt = self._base_font_pt * scale
        f = QFont(self.font())
        f.setPointSizeF(pt)
        self.setFont(f)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._apply_responsive_layout()
        self._apply_scale()
    def closeEvent(self, event):

        event.accept()


def main():
    app = QApplication(sys.argv)
    w = MainWindow()
    w.showMaximized()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()