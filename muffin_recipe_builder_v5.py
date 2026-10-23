"""
Muffin / Brownie / Cookie / Granola Cluster Macro Calculator (v5a)
--------------------------------------------------------------
Same CustomTkinter app as v4, with the ingredient updater built in.

v5a changes (everything else is identical to v5):
- After picking an ingredient, it now stays in the Search bar, so the
  Search bar shows what is selected. The results list narrows to match.
  Clicking into the Search bar selects all its text, so typing replaces
  the old name.
- The "Selected:" line turns purple with light grey text once an
  ingredient is picked. Before that it stays plain.
- After "Add to Recipe", everything resets: the Search bar is blank, the
  Selected line goes back to plain "(none selected)", and the full
  ingredient list shows again. This was not the case before.
- Clearing the Search bar by hand always brings the full list back.

v5 now merges the separate script (update_ingredients.py) in v4 to turn
ingredients.csv into ingredients.json. The app now reads
ingredients.csv directly every time it opens, so any edit to the CSV
shows up the next time you launch it. There is no JSON file anymore.
As background, v4 switched from regular Tkinter to CustomTkinter, 
which is a separate package. 

How to use:
- Pick an ingredient, enter grams (or cups, tbsp, tsp, pieces, where known),
  and it looks up the macros the same way as the previous Excel workbook VLOOKUP.
- Add several ingredients to build a full recipe.
- Set the number of servings to get per-serving macros.
- Save the finished recipe as a NEW SHEET in an Excel workbook, just
  like excel.

This is the ground-truth file needed in the same folder as this script:
    ingredients.csv
Make sure it has the correct numbers. 
    
Expected CSV columns (case-sensitive, any order):
    Food, Meal Type, Grams total, Calories, Fat g, Protein g,
    Carbs g, Fiber g, Sugar g,
    Grams per tbsp, Grams per tsp, Grams per cup, Grams per piece,
    Default g/volume, Food type

"Grams total" is the amount that all the macro columns are based on.
It does NOT need to be 100g. The app converts to per-gram automatically.

The app will try to convert the volumes to grams using the column 
`Default g/volume`, so try to fill this up if adding a new ingredient.
The four "Grams per {volume}" columns will be calculated by the app. 

If a row has a problem (missing number, text where a number belongs,
duplicate name), the app skips that row, opens normally, and shows a
message listing which rows were skipped and why.

Extra dependencies for this version:
    pip install customtkinter openpyxl

Run with:
    python muffin_recipe_builder_v5.py
"""

import csv
import io
import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime

import customtkinter as ctk

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

APP_DIR = os.path.dirname(os.path.abspath(__file__))
CSV_FILE = os.path.join(APP_DIR, "ingredients.csv")

# ---------------------------------------------------------- csv columns
REQUIRED_COLUMNS = [
    "Food",
    "Grams total",
    "Calories",
    "Fat g",
    "Protein g",
    "Carbs g",
    "Fiber g",
    "Sugar g",
]
OPTIONAL_TEXT_COLUMNS = ["Meal Type", "Default g/volume", "Food type"]
UNIT_COLUMNS = {
    "Grams per tbsp": "tbsp",
    "Grams per tsp": "tsp",
    "Grams per cup": "cup",
    "Grams per piece": "piece",
}
COLUMN_TO_KEY = {
    "Grams total": "base_grams",
    "Calories": "calories",
    "Fat g": "fat_g",
    "Protein g": "protein_g",
    "Carbs g": "carbs_g",
    "Fiber g": "fiber_g",
    "Sugar g": "sugar_g",
    "Default g/volume": "default_g_volume",
    "Food type": "food_type",
    "Meal Type": "meal_type",
}
NUMERIC_COLUMNS = ["Grams total", "Calories", "Fat g", "Protein g", "Carbs g", "Fiber g", "Sugar g"]

MACRO_FIELDS = [
    ("calories", "Calories"),
    ("fat_g", "Fat (g)"),
    ("protein_g", "Protein (g)"),
    ("carbs_g", "Carbs (g)"),
    ("fiber_g", "Fiber (g)"),
    ("sugar_g", "Sugar (g)"),
]

# Order units should appear in the dropdown, and how they're labeled.
UNIT_LABELS = [("g", "g"), ("tbsp", "tbsp"), ("tsp", "tsp"), ("cup", "cup"), ("piece", "piece")]

ACCENT = "#3B8ED0"
ACCENT_HOVER = "#3679AC"

SEARCH_PLACEHOLDER = "Start typing…"

# "Selected:" bar once an ingredient is picked (fixed colors, same in light and dark mode)
SELECTED_BG = "#6B3FA0"   # purple
SELECTED_FG = "#E4E4E4"   # light grey


# ---------------------------------------------------------- ingredient csv loader
class CsvLoadError(Exception):
    """Raised when ingredients.csv can't be used at all (missing file, missing columns, no valid rows)."""


def read_csv_text(path):
    """
    Excel on Windows usually saves CSV files as cp1252 ('ANSI'), not UTF-8.
    Decode LINE BY LINE (not the whole file at once): if we decoded the whole
    file together, a single bad byte on one row would fail UTF-8 for every
    row and silently fall back to cp1252 for the entire file, mangling every
    other row that actually was correct UTF-8 (curly quotes turning into
    garbage like 'â€™'). Per-line decoding keeps good rows good and only
    falls back on the specific row that needs it.
    """
    with open(path, "rb") as f:
        raw = f.read()

    # Strip a UTF-8 BOM if present, since we're decoding manually now.
    if raw.startswith(b"\xef\xbb\xbf"):
        raw = raw[3:]

    encodings_to_try = ["utf-8", "cp1252", "latin-1"]
    lines_out = []
    for line_num, raw_line in enumerate(raw.split(b"\n"), start=1):
        line_text = None
        for enc in encodings_to_try:
            try:
                line_text = raw_line.decode(enc)
                if enc != "utf-8":
                    print(f"(ingredients.csv line {line_num} wasn't valid UTF-8, read it as {enc} instead.)")
                break
            except UnicodeDecodeError:
                continue
        if line_text is None:
            # Should be unreachable, latin-1 can decode any byte, but just in case:
            line_text = raw_line.decode("latin-1", errors="replace")
        lines_out.append(line_text.rstrip("\r"))  # undo manual \r\n split

    return "\n".join(lines_out)


def load_ingredients_from_csv(path):
    """
    Reads ingredients.csv and returns (ingredients, problems).

    ingredients: dict of {food name: data dict}, ready for Ingredient(name, data).
    problems:    list of plain-English messages, one per skipped row or skipped unit.

    Rows with a problem are skipped and reported. Everything valid still loads.
    Raises CsvLoadError only when the whole file is unusable.
    """
    if not os.path.exists(path):
        raise CsvLoadError(
            f"Could not find ingredients.csv in:\n{os.path.dirname(path)}\n\n"
            "Put ingredients.csv in the same folder as this script and try again."
        )

    try:
        csv_text = read_csv_text(path)
    except PermissionError:
        raise CsvLoadError(
            "Could not open ingredients.csv. If it's open in Excel, close it and try again."
        )

    reader = csv.DictReader(io.StringIO(csv_text))
    header = reader.fieldnames or []
    missing = [col for col in REQUIRED_COLUMNS if col not in header]
    if missing:
        raise CsvLoadError(
            "ingredients.csv is missing required column(s):\n  "
            + ", ".join(missing)
            + "\n\nFound columns:\n  "
            + ", ".join(header)
        )

    ingredients = {}
    problems = []

    for row_num, row in enumerate(reader, start=2):  # row 1 is the header
        food = (row.get("Food") or "").strip()
        if not food:
            continue  # blank row

        grams_raw = (row.get("Grams total") or "").strip()
        if not grams_raw:
            problems.append(f"Row {row_num} ('{food}'): missing 'Grams total'. Row skipped.")
            continue

        try:
            grams = float(grams_raw)
            if grams <= 0:
                raise ValueError
        except ValueError:
            problems.append(
                f"Row {row_num} ('{food}'): 'Grams total' must be a number greater than 0 "
                f"(got '{grams_raw}'). Row skipped."
            )
            continue

        entry = {"base_grams": grams}
        row_ok = True
        for col in NUMERIC_COLUMNS:
            if col == "Grams total":
                continue
            raw_val = (row.get(col) or "").strip()
            if raw_val == "":
                entry[COLUMN_TO_KEY[col]] = 0.0
                continue
            try:
                entry[COLUMN_TO_KEY[col]] = float(raw_val)
            except ValueError:
                problems.append(
                    f"Row {row_num} ('{food}'): '{col}' should be a number "
                    f"(got '{raw_val}'). Row skipped."
                )
                row_ok = False
                break
        if not row_ok:
            continue

        for col in OPTIONAL_TEXT_COLUMNS:
            entry[COLUMN_TO_KEY[col]] = (row.get(col) or "").strip()

        # Optional volume conversions (grams per 1 tbsp / tsp / cup / piece)
        units = {}
        for col, unit_key in UNIT_COLUMNS.items():
            raw_val = (row.get(col) or "").strip() if col in header else ""
            if raw_val == "":
                continue
            try:
                val = float(raw_val)
                if val <= 0:
                    raise ValueError
                units[unit_key] = val
            except ValueError:
                problems.append(
                    f"Row {row_num} ('{food}'): '{col}' should be a number greater than 0 "
                    f"(got '{raw_val}'). That unit was skipped for this ingredient."
                )
        entry["units"] = units

        if food in ingredients:
            problems.append(f"Row {row_num}: duplicate ingredient name '{food}'. Kept the earlier one.")
            continue

        ingredients[food] = entry

    if not ingredients:
        raise CsvLoadError("No valid ingredient rows found in ingredients.csv. Check the file and try again.")

    return ingredients, problems


class Ingredient:
    def __init__(self, name, data):
        self.name = name
        self.base_grams = data["base_grams"]
        self.per_gram = {
            key: data.get(key, 0.0) / self.base_grams for key, _ in MACRO_FIELDS
        }
        self.default_g_volume = data.get("default_g_volume", "")
        self.food_type = data.get("food_type", "")
        self.meal_type = data.get("meal_type", "")
        # units: dict like {"tbsp": 14.2, "cup": 96.0} = grams per 1 of that unit,
        # filled in explicitly in ingredients.csv (no guessing from free text).
        self.units = data.get("units", {}) or {}

    def available_units(self):
        units = ["g"] + [u for u, _ in UNIT_LABELS if u != "g" and u in self.units]
        return units

    def grams_for(self, amount, unit):
        if unit == "g":
            return amount
        grams_per_unit = self.units.get(unit)
        if not grams_per_unit:
            return None
        return amount * grams_per_unit

    def macros_for_grams(self, grams):
        return {key: self.per_gram[key] * grams for key, _ in MACRO_FIELDS}


class MacroCalculatorApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Muffin / Brownie / Cookie Macro Calculator")
        self.geometry("1150x780")
        self.minsize(980, 700)

        self.font_body = ctk.CTkFont(size=12)
        self.font_bold = ctk.CTkFont(size=12, weight="bold")
        self.font_title = ctk.CTkFont(size=16, weight="bold")
        self.font_subtitle = ctk.CTkFont(size=11)
        self.font_mono = ctk.CTkFont(size=12, family="Consolas")

        self.ingredients = {}
        self.recipe_rows = []  # list of dicts: name, grams, macros
        self.workbook_path = None

        # search-result state (replaces the old tk.Listbox selection model)
        self._current_names = []
        self._result_buttons = []
        self._cursor_index = -1

        self._startup_problems = []

        self._load_ingredients()
        self._build_ui()

        if self._startup_problems:
            # Wait until the window is drawn so the message appears on top of it.
            self.after(300, self._show_startup_problems)

    # ---------------------------------------------------------- data
    def _load_ingredients(self):
        """Reads ingredients.csv straight into Ingredient objects. No JSON step."""
        try:
            raw, problems = load_ingredients_from_csv(CSV_FILE)
        except CsvLoadError as exc:
            self.ingredients = {}
            self.after(300, lambda: messagebox.showerror("Ingredients file problem", str(exc)))
            return
        self.ingredients = {name: Ingredient(name, data) for name, data in raw.items()}
        self._startup_problems = problems

    def _show_startup_problems(self):
        shown = self._startup_problems[:15]
        extra = len(self._startup_problems) - len(shown)
        msg = (
            f"Loaded {len(self.ingredients)} ingredients from ingredients.csv.\n\n"
            f"{len(self._startup_problems)} problem(s) found. Everything else loaded fine:\n\n- "
            + "\n- ".join(shown)
        )
        if extra > 0:
            msg += f"\n\n...and {extra} more."
        messagebox.showwarning("Some ingredient rows were skipped", msg)

    # ---------------------------------------------------------- ui
    def _build_ui(self):
        # ---------- header ----------
        header = ctk.CTkFrame(self, corner_radius=0, fg_color=("gray92", "gray14"))
        header.pack(fill="x")
        ctk.CTkLabel(
            header, text="Macro Calculator", font=self.font_title, anchor="w"
        ).pack(side="left", padx=16, pady=(8, 2))
        ctk.CTkLabel(
            header,
            text="muffins · brownies · cookies · granola clusters",
            font=self.font_subtitle,
            text_color=("gray35", "gray65"),
            anchor="w",
        ).pack(side="left", padx=(0, 16), pady=(11, 3))

        self.appearance_switch = ctk.CTkSwitch(
            header, text="Dark mode", command=self._toggle_appearance, font=self.font_body
        )
        self.appearance_switch.pack(side="right", padx=16, pady=6)
        if ctk.get_appearance_mode() == "Dark":
            self.appearance_switch.select()

        # ---------- scrollable body ----------
        # Everything below the header lives in here. v4's extra header, padding,
        # and switch made the full stack taller than v3's; on smaller screens or
        # windows this pushed the totals panel below the visible area with no way
        # to reach it. Scrolling guarantees it's always reachable.
        self.body = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.body.pack(fill="both", expand=True)

        # ---------- top: search + amount/unit/add ----------
        top = ctk.CTkFrame(self.body, corner_radius=14)
        top.pack(fill="x", padx=16, pady=(10, 6))
        top.grid_columnconfigure(1, weight=1)

        # -- search column --
        search_col = ctk.CTkFrame(top, fg_color="transparent")
        search_col.grid(row=0, column=0, sticky="nw", padx=10, pady=8)

        ctk.CTkLabel(search_col, text="Search ingredient", font=self.font_bold).pack(anchor="w")
        self.search_var = tk.StringVar()
        search_entry = ctk.CTkEntry(
            search_col, textvariable=self.search_var, width=340, height=24,
            placeholder_text=SEARCH_PLACEHOLDER,
        )
        search_entry.pack(anchor="w", pady=(3, 4))
        search_entry.bind("<KeyRelease>", self._filter_ingredients)
        search_entry.bind("<Down>", lambda e: self._move_cursor(1))
        search_entry.bind("<Up>", lambda e: self._move_cursor(-1))
        search_entry.bind("<Return>", self._confirm_cursor)
        search_entry.bind("<FocusIn>", self._select_all_in_search)
        self.search_entry = search_entry

        self.results_frame = ctk.CTkScrollableFrame(
            search_col, width=340, height=84, corner_radius=10, 
            fg_color=("gray95", "gray17"), label_text="",
        )
        self.results_frame.pack(anchor="w")
        # CTkScrollableFrame creates its internal vertical scrollbar with no
        # height, which defaults to 200px. Since the scrollbar and canvas
        # share one grid row (sticky="nsew"), that 200px default otherwise
        # wins and the box never gets shorter than ~200px no matter what
        # height is passed above. Shrink it to match.
        self.results_frame._scrollbar.configure(height=1)

        self.selected_row = ctk.CTkFrame(search_col, fg_color="transparent", corner_radius=8)
        self.selected_row.pack(anchor="w", fill="x", pady=(4, 0))
        self.selected_prefix_label = ctk.CTkLabel(
            self.selected_row, text="Selected:", font=self.font_body
        )
        self.selected_prefix_label.pack(side="left", padx=(8, 0), pady=3)
        self._prefix_default_color = self.selected_prefix_label.cget("text_color")
        self.selected_ingredient_var = tk.StringVar(value="(none selected)")
        self.selected_value_label = ctk.CTkLabel(
            self.selected_row, textvariable=self.selected_ingredient_var, font=self.font_bold,
            text_color=ACCENT,
        )
        self.selected_value_label.pack(side="left", padx=(6, 8), pady=3)

        self.default_note_var = tk.StringVar(value="")
        ctk.CTkLabel(
            search_col, textvariable=self.default_note_var, font=self.font_subtitle,
            text_color=("gray35", "gray65"), anchor="w", justify="left", wraplength=340,
        ).pack(anchor="w", pady=(1, 0))

        # -- amount / unit / add column --
        entry_col = ctk.CTkFrame(top, fg_color="transparent")
        entry_col.grid(row=0, column=1, sticky="n", padx=10, pady=8)

        row1 = ctk.CTkFrame(entry_col, fg_color="transparent")
        row1.pack(anchor="w")

        amount_box = ctk.CTkFrame(row1, fg_color="transparent")
        amount_box.pack(side="left", padx=(0, 12))
        ctk.CTkLabel(amount_box, text="Amount", font=self.font_bold).pack(anchor="w")
        self.amount_var = tk.StringVar(value="0")
        self.amount_entry = ctk.CTkEntry(amount_box, textvariable=self.amount_var, width=110, height=26)
        self.amount_entry.pack(anchor="w", pady=(4, 0))

        unit_box = ctk.CTkFrame(row1, fg_color="transparent")
        unit_box.pack(side="left")
        ctk.CTkLabel(unit_box, text="Unit", font=self.font_bold).pack(anchor="w")
        self.unit_var = tk.StringVar(value="g")
        self.unit_menu = ctk.CTkOptionMenu(unit_box, variable=self.unit_var, values=["g"], width=110, height=26)
        self.unit_menu.pack(anchor="w", pady=(4, 0))

        ctk.CTkButton(
            entry_col, text="+  Add to Recipe", command=self._add_ingredient,
            fg_color=ACCENT, hover_color=ACCENT_HOVER, height=32, corner_radius=10,
            font=self.font_bold,
        ).pack(anchor="w", pady=(12, 0), fill="x")

        self._populate_results(sorted(self.ingredients.keys()))

        # ---------- recipe table ----------
        mid = ctk.CTkFrame(self.body, corner_radius=14)
        mid.pack(fill="both", expand=True, padx=16, pady=6)
        mid.grid_rowconfigure(1, weight=1)
        mid.grid_columnconfigure(0, weight=1)

        table_header = ctk.CTkFrame(mid, fg_color="transparent")
        table_header.grid(row=0, column=0, sticky="ew", padx=10, pady=(8, 3))
        ctk.CTkLabel(table_header, text="Recipe", font=self.font_bold).pack(side="left")
        ctk.CTkLabel(
            table_header, text="Right-click, Delete, or Backspace to remove a row",
            font=self.font_subtitle, text_color=("gray35", "gray65"),
        ).pack(side="right")

        tree_holder = ctk.CTkFrame(mid, corner_radius=10, fg_color=("gray95", "gray17"))
        tree_holder.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        tree_holder.grid_rowconfigure(0, weight=1)
        tree_holder.grid_columnconfigure(0, weight=1)

        columns = ["food", "entered", "grams"] + [key for key, _ in MACRO_FIELDS]
        headings = ["Food", "You Entered", "Grams"] + [label for _, label in MACRO_FIELDS]
        self.tree = ttk.Treeview(
            tree_holder, columns=columns, show="headings", height=12, selectmode="extended"
        )
        for col, head in zip(columns, headings):
            self.tree.heading(col, text=head)
            self.tree.column(col, width=95 if col != "food" else 190, anchor="center")
        self.tree.column("food", anchor="w")
        self.tree.grid(row=0, column=0, sticky="nsew", padx=(2, 0), pady=2)

        scrollbar = ttk.Scrollbar(tree_holder, orient="vertical", command=self.tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns", pady=2)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self._style_treeview()

        # Right-click context menu (like Excel's "Delete Row")
        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label="Delete Selected Row(s)", command=self._remove_selected)
        self.tree.bind("<Button-3>", self._show_context_menu)  # Windows / Linux right-click
        self.tree.bind("<Button-2>", self._show_context_menu)  # Mac right-click / two-finger click
        self.tree.bind("<Delete>", lambda e: self._remove_selected())
        self.tree.bind("<BackSpace>", lambda e: self._remove_selected())

        # ---------- bottom controls ----------
        bottom = ctk.CTkFrame(self.body, corner_radius=14)
        bottom.pack(fill="x", padx=16, pady=(0, 6))

        controls = ctk.CTkFrame(bottom, fg_color="transparent")
        controls.pack(fill="x", padx=10, pady=8)

        ctk.CTkButton(
            controls, text="Remove Selected", command=self._remove_selected,
            fg_color="transparent", border_width=1, text_color=("gray10", "gray90"),
            hover_color=("gray85", "gray25"), corner_radius=8,
        ).pack(side="left")

        ctk.CTkLabel(controls, text="Servings:", font=self.font_bold).pack(side="left", padx=(24, 6))
        self.servings_var = tk.StringVar(value="8")
        servings_entry = ctk.CTkEntry(controls, textvariable=self.servings_var, width=60)
        servings_entry.pack(side="left")
        servings_entry.bind("<KeyRelease>", lambda e: self._refresh_totals())

        ctk.CTkButton(
            controls, text="Save Recipe to Excel", command=self._save_recipe,
            fg_color=ACCENT, hover_color=ACCENT_HOVER, corner_radius=8,
        ).pack(side="right", padx=(8, 0))
        ctk.CTkButton(
            controls, text="Load Recipe from Excel…", command=self._load_recipe_from_excel,
            fg_color="transparent", border_width=1, text_color=("gray10", "gray90"),
            hover_color=("gray85", "gray25"), corner_radius=8,
        ).pack(side="right", padx=(8, 0))
        ctk.CTkButton(
            controls, text="Load Workbook…", command=self._pick_workbook_button,
            fg_color="transparent", border_width=1, text_color=("gray10", "gray90"),
            hover_color=("gray85", "gray25"), corner_radius=8,
        ).pack(side="right", padx=(8, 0))

        # ---------- totals ----------
        totals_frame = ctk.CTkFrame(self.body, corner_radius=14)
        totals_frame.pack(fill="x", padx=16, pady=(0, 6))

        inner = ctk.CTkFrame(totals_frame, fg_color="transparent")
        inner.pack(fill="x", padx=10, pady=8)
        ctk.CTkLabel(inner, text="Totals & Per Serving", font=self.font_bold).grid(
            row=0, column=0, sticky="w", pady=(0, 4)
        )

        for i, (key, label) in enumerate(MACRO_FIELDS):
            ctk.CTkLabel(inner, text=label, font=self.font_bold).grid(
                row=1, column=i + 1, padx=10
            )

        ctk.CTkLabel(inner, text="Recipe Total", font=self.font_body).grid(
            row=2, column=0, sticky="w", pady=2
        )
        ctk.CTkLabel(inner, text="Per Serving", font=self.font_body).grid(
            row=3, column=0, sticky="w", pady=2
        )

        self.total_vars = {key: tk.StringVar(value="0") for key, _ in MACRO_FIELDS}
        self.per_serving_vars = {key: tk.StringVar(value="0") for key, _ in MACRO_FIELDS}
        for i, (key, _) in enumerate(MACRO_FIELDS):
            ctk.CTkLabel(inner, textvariable=self.total_vars[key], font=self.font_mono).grid(
                row=2, column=i + 1
            )
            ctk.CTkLabel(inner, textvariable=self.per_serving_vars[key], font=self.font_mono).grid(
                row=3, column=i + 1
            )

        self.workbook_label_var = tk.StringVar(value="Workbook: (none selected — Save will ask)")
        ctk.CTkLabel(
            self.body, textvariable=self.workbook_label_var, font=self.font_subtitle,
            text_color=("gray35", "gray65"),
        ).pack(anchor="w", padx=20, pady=(0, 6))

    # ------------------------------------------------------ appearance
    def _toggle_appearance(self):
        ctk.set_appearance_mode("Dark" if self.appearance_switch.get() else "Light")
        self._style_treeview()

    def _style_treeview(self):
        dark = ctk.get_appearance_mode() == "Dark"
        bg = "#1C1C30" if dark else "#F2F2F2"
        fg = "#F2F2F2" if dark else "#2D1111"
        head_bg = "#17284F" if dark else "#E2E2E2"
        select_bg = ACCENT

        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "Treeview", background=bg, foreground=fg, fieldbackground=bg,
            borderwidth=0, rowheight=28, font=("", 11),
        )
        style.configure(
            "Treeview.Heading", background=head_bg, foreground=fg,
            borderwidth=0, relief="flat", font=("", 11, "bold"),
        )
        style.map("Treeview", background=[("selected", select_bg)], foreground=[("selected", "white")])
        style.layout("Treeview", [("Treeview.treearea", {"sticky": "nswe"})])

    # ------------------------------------------------------ search results
    def _populate_results(self, names):
        for btn in self._result_buttons:
            btn.destroy()
        self._result_buttons = []
        self._current_names = names
        self._cursor_index = -1

        for name in names:
            btn = ctk.CTkButton(
                self.results_frame, text=name, anchor="w", fg_color="transparent",
                text_color=("gray10", "gray90"), hover_color=("gray85", "gray25"),
                corner_radius=6, height=28, font=self.font_body,
                command=lambda n=name: self._select_ingredient(n),
            )
            btn.pack(fill="x", padx=4, pady=1)
            self._result_buttons.append(btn)

    def _filter_ingredients(self, event=None):
        # Arrow keys / Enter shouldn't re-run the filter — the cursor/confirm
        # handlers below take care of those.
        if event is not None and event.keysym in ("Down", "Up", "Return", "Escape"):
            return
        typed = self.search_var.get().lower().strip()
        if typed == SEARCH_PLACEHOLDER.lower().strip():
            typed = ""  # the grey hint text is not a real search
        names = sorted(self.ingredients.keys())
        if typed:
            names = [n for n in names if typed in n.lower()]
        self._populate_results(names)

    def _move_cursor(self, delta):
        if not self._current_names:
            return "break"
        if self._cursor_index == -1:
            new_index = 0
        else:
            new_index = min(max(self._cursor_index + delta, 0), len(self._current_names) - 1)
        self._highlight_cursor(new_index)
        return "break"

    def _highlight_cursor(self, index):
        if 0 <= self._cursor_index < len(self._result_buttons):
            self._result_buttons[self._cursor_index].configure(fg_color="transparent")
        self._cursor_index = index
        if 0 <= index < len(self._result_buttons):
            self._result_buttons[index].configure(fg_color=("gray80", "gray30"))
            self.results_frame._parent_canvas.yview_moveto(index / max(len(self._result_buttons), 1))

    def _confirm_cursor(self, event=None):
        if 0 <= self._cursor_index < len(self._current_names):
            self._select_ingredient(self._current_names[self._cursor_index])
        self.amount_entry.focus_set()
        self.amount_entry.select_range(0, "end")
        return "break"

    def _select_ingredient(self, name):
        ing = self.ingredients.get(name)
        if not ing:
            return
        self.selected_ingredient_var.set(name)
        self._highlight_selected_row()
        self.unit_menu.configure(values=ing.available_units())
        self.unit_var.set("g")
        self.default_note_var.set(f"Reference: {ing.default_g_volume}" if ing.default_g_volume else "")

        if name in self._current_names:
            self._highlight_cursor(self._current_names.index(name))

        # Put the chosen name in the Search bar. This is done a moment later
        # because a mouse click comes from one of the result buttons, and the
        # refresh below deletes those buttons. Waiting 150 ms lets the button
        # finish its own click animation first.
        self.after(150, lambda n=name: self._sync_search_to_selection(n))

    def _sync_search_to_selection(self, name):
        if self.selected_ingredient_var.get() != name:
            return  # a different ingredient was picked during the delay
        self.search_var.set(name)
        self._filter_ingredients()
        if name in self._current_names:
            self._highlight_cursor(self._current_names.index(name))

    def _reset_selection(self):
        """Clean slate after an ingredient is added: blank search, plain Selected line, full list."""
        self.selected_ingredient_var.set("(none selected)")
        self.selected_row.configure(fg_color="transparent")
        self.selected_prefix_label.configure(fg_color="transparent", text_color=self._prefix_default_color)
        self.selected_value_label.configure(fg_color="transparent", text_color=ACCENT)
        self.default_note_var.set("")
        self.unit_menu.configure(values=["g"])
        self.unit_var.set("g")
        self.search_var.set("")
        # Rebuild the list directly instead of reading the search box back.
        self._populate_results(sorted(self.ingredients.keys()))
        self.results_frame._parent_canvas.yview_moveto(0)

    def _highlight_selected_row(self):
        self.selected_row.configure(fg_color=SELECTED_BG)
        for label in (self.selected_prefix_label, self.selected_value_label):
            label.configure(fg_color=SELECTED_BG, text_color=SELECTED_FG)

    def _select_all_in_search(self, event=None):
        # Short delay so Tk finishes its own click handling first. Otherwise it
        # clears the selection right after we make it.
        self.after(10, lambda: self.search_entry.select_range(0, "end"))

    # ------------------------------------------------------ recipe rows
    def _add_ingredient(self):
        name = self.selected_ingredient_var.get().strip()
        ing = self.ingredients.get(name)
        if not ing:
            messagebox.showwarning("Pick an ingredient", "Search and click an ingredient from the list first.")
            return
        try:
            amount = float(self.amount_var.get())
        except ValueError:
            messagebox.showwarning("Invalid amount", "Enter a number for the amount.")
            return
        if amount <= 0:
            messagebox.showwarning("Invalid amount", "Amount must be greater than zero.")
            return

        unit = self.unit_var.get()
        grams = ing.grams_for(amount, unit)
        if grams is None:
            messagebox.showwarning(
                "No conversion available",
                f"'{name}' doesn't have a known '{unit}' conversion. Use grams, "
                "or add the conversion in ingredients.csv and reopen the app.",
            )
            return

        macros = ing.macros_for_grams(grams)
        entered_label = f"{amount:g} {unit}"
        self.recipe_rows.append(
            {"name": name, "amount": amount, "unit": unit, "entered": entered_label,
             "grams": grams, "macros": macros}
        )
        self._refresh_table()
        self.amount_var.set("0")
        self._reset_selection()

    def _show_context_menu(self, event):
        row_id = self.tree.identify_row(event.y)
        if not row_id:
            return
        # If you right-click a row that isn't already selected, select just that one
        # — same behavior Excel uses before showing its right-click menu.
        if row_id not in self.tree.selection():
            self.tree.selection_set(row_id)
        self.context_menu.tk_popup(event.x_root, event.y_root)

    def _remove_selected(self):
        selected = self.tree.selection()
        if not selected:
            return
        indices = sorted((self.tree.index(item) for item in selected), reverse=True)
        for idx in indices:
            del self.recipe_rows[idx]
        self._refresh_table()

    def _refresh_table(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        for row in self.recipe_rows:
            values = [row["name"], row["entered"], f"{row['grams']:.1f}"]
            values += [f"{row['macros'][key]:.1f}" for key, _ in MACRO_FIELDS]
            self.tree.insert("", "end", values=values)
        self._refresh_totals()

    def _refresh_totals(self):
        totals = {key: 0.0 for key, _ in MACRO_FIELDS}
        for row in self.recipe_rows:
            for key, _ in MACRO_FIELDS:
                totals[key] += row["macros"][key]
        try:
            servings = float(self.servings_var.get())
            if servings <= 0:
                servings = 1
        except ValueError:
            servings = 1

        for key, _ in MACRO_FIELDS:
            self.total_vars[key].set(f"{totals[key]:.1f}")
            self.per_serving_vars[key].set(f"{totals[key] / servings:.2f}")

    # ------------------------------------------------------ save to excel
    def _choose_existing_workbook(self):
        path = filedialog.askopenfilename(
            title="Choose your Excel workbook",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if path:
            self.workbook_path = path
            self.workbook_label_var.set(f"Workbook: {path}")

    def _create_new_workbook(self):
        path = filedialog.asksaveasfilename(
            title="Name your new Excel workbook",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not path:
            return
        # Don't write the file yet — an Excel workbook can't have zero sheets,
        # so we create it for real once the first recipe sheet is ready to go in.
        self.workbook_path = path
        self.workbook_label_var.set(f"Workbook: {path} (will be created on save)")

    def _ask_new_or_existing(self):
        """Returns 'new', 'existing', or None (cancelled)."""
        choice = {"value": None}
        win = ctk.CTkToplevel(self)
        win.title("Choose a workbook")
        win.geometry("420x180")
        win.transient(self)
        win.grab_set()

        ctk.CTkLabel(
            win, text="Do you want to create a brand-new Excel file,\nor save into one you already have?",
            font=self.font_body, justify="center",
        ).pack(pady=(24, 16))
        btns = ctk.CTkFrame(win, fg_color="transparent")
        btns.pack(pady=(0, 16))

        def pick(value):
            choice["value"] = value
            win.destroy()

        ctk.CTkButton(btns, text="Create New…", command=lambda: pick("new"),
                      fg_color=ACCENT, hover_color=ACCENT_HOVER, corner_radius=8).pack(side="left", padx=6)
        ctk.CTkButton(btns, text="Choose Existing…", command=lambda: pick("existing"),
                      fg_color="transparent", border_width=1, text_color=("gray10", "gray90"),
                      hover_color=("gray85", "gray25"), corner_radius=8).pack(side="left", padx=6)
        ctk.CTkButton(btns, text="Cancel", command=lambda: pick(None),
                      fg_color="transparent", border_width=1, text_color=("gray10", "gray90"),
                      hover_color=("gray85", "gray25"), corner_radius=8).pack(side="left", padx=6)
        win.wait_window()
        return choice["value"]

    def _pick_workbook_button(self):
        choice = self._ask_new_or_existing()
        if choice == "new":
            self._create_new_workbook()
        elif choice == "existing":
            self._choose_existing_workbook()

    def _ask_pick_sheet(self, sheet_names):
        choice = {"value": None}
        win = ctk.CTkToplevel(self)
        win.title("Choose a recipe to load")
        win.geometry("380x" + str(min(430, 140 + 36 * len(sheet_names))))
        win.transient(self)
        win.grab_set()

        ctk.CTkLabel(
            win, text="Which saved recipe do you want to load?", font=self.font_body
        ).pack(pady=(18, 10))

        list_frame = ctk.CTkScrollableFrame(win, width=320, height=min(300, 36 * len(sheet_names) + 4))
        list_frame.pack(padx=16)

        def confirm(name):
            choice["value"] = name
            win.destroy()

        for name in sheet_names:
            ctk.CTkButton(
                list_frame, text=name, anchor="w", fg_color="transparent",
                text_color=("gray10", "gray90"), hover_color=("gray85", "gray25"),
                corner_radius=6, height=30, command=lambda n=name: confirm(n),
            ).pack(fill="x", padx=2, pady=2)

        ctk.CTkButton(
            win, text="Cancel", command=win.destroy, fg_color="transparent", border_width=1,
            text_color=("gray10", "gray90"), hover_color=("gray85", "gray25"), corner_radius=8,
        ).pack(pady=14)
        win.wait_window()
        return choice["value"]

    def _parse_recipe_sheet(self, ws):
        """
        Reads a saved recipe sheet back into recipe rows. Works with the
        current format (with a "You Entered" column) and older sheets
        saved before that column existed (falls back to grams).
        Returns (recipe_rows, servings_or_None, list_of_unmatched_ingredient_names).
        """
        rows_iter = ws.iter_rows(values_only=True)
        try:
            header = next(rows_iter)
        except StopIteration:
            return [], None, []

        col_index = {}
        for i, cell in enumerate(header):
            if cell is None:
                continue
            col_index[str(cell).strip().lower()] = i

        food_idx = col_index.get("food")
        entered_idx = col_index.get("you entered")
        grams_idx = col_index.get("grams total")
        macro_idx = {}
        for key, label in MACRO_FIELDS:
            idx = col_index.get(label.lower())
            if idx is not None:
                macro_idx[key] = idx

        if food_idx is None:
            return [], None, []

        loaded_rows = []
        unmatched = []
        servings_found = None

        for row in rows_iter:
            if row is None or len(row) <= food_idx:
                continue
            food_val = row[food_idx]
            if food_val is None or str(food_val).strip() == "":
                continue  # skips blank divider row and the repeated mini-header row
            food_val = str(food_val).strip()

            if food_val.lower() == "total":
                continue
            if food_val.lower() == "number of servings":
                if macro_idx:
                    first_macro_idx = next(iter(macro_idx.values()))
                    if len(row) > first_macro_idx and row[first_macro_idx] is not None:
                        try:
                            servings_found = float(row[first_macro_idx])
                        except (TypeError, ValueError):
                            pass
                continue
            if food_val.lower() == "nutrition per serving":
                continue

            grams_saved = None
            if grams_idx is not None and len(row) > grams_idx and row[grams_idx] is not None:
                try:
                    grams_saved = float(row[grams_idx])
                except (TypeError, ValueError):
                    grams_saved = None

            entered_saved = None
            if entered_idx is not None and len(row) > entered_idx and row[entered_idx] is not None:
                entered_saved = str(row[entered_idx]).strip()

            macros_saved = {}
            for key, idx in macro_idx.items():
                val = row[idx] if len(row) > idx else None
                try:
                    macros_saved[key] = float(val) if val is not None else 0.0
                except (TypeError, ValueError):
                    macros_saved[key] = 0.0

            ing = self.ingredients.get(food_val)

            if ing:
                amount, unit = None, None
                if entered_saved:
                    parts = entered_saved.rsplit(" ", 1)
                    if len(parts) == 2:
                        try:
                            amount = float(parts[0])
                            unit = parts[1]
                        except ValueError:
                            amount = None
                grams = ing.grams_for(amount, unit) if amount is not None and unit is not None else None
                if grams is None:
                    # No usable "You Entered" text (old sheet, or unit no longer known) —
                    # fall back to the saved grams total, macros recalculated fresh.
                    grams = grams_saved if grams_saved is not None else 0.0
                    amount, unit = grams, "g"
                macros = ing.macros_for_grams(grams)
                entered_label = entered_saved or f"{grams:g} g"
                loaded_rows.append(
                    {"name": food_val, "amount": amount, "unit": unit, "entered": entered_label,
                     "grams": grams, "macros": macros}
                )
            else:
                # Ingredient isn't in the current database (renamed/removed) —
                # keep the sheet's saved numbers as-is rather than guessing.
                unmatched.append(food_val)
                grams = grams_saved if grams_saved is not None else 0.0
                entered_label = entered_saved or f"{grams:g} g"
                loaded_rows.append(
                    {"name": food_val, "amount": grams, "unit": "g", "entered": entered_label,
                     "grams": grams, "macros": macros_saved}
                )

        return loaded_rows, servings_found, unmatched

    def _load_recipe_from_excel(self):
        if not self.workbook_path or not os.path.exists(self.workbook_path):
            self._choose_existing_workbook()
            if not self.workbook_path or not os.path.exists(self.workbook_path):
                return

        try:
            import openpyxl
        except ImportError:
            messagebox.showerror(
                "Missing dependency", "This feature needs openpyxl. Install it with:\n\npip install openpyxl"
            )
            return

        try:
            wb = openpyxl.load_workbook(self.workbook_path, data_only=True)
        except Exception as exc:
            messagebox.showerror("Could not open workbook", str(exc))
            return

        candidate_sheets = [s for s in wb.sheetnames if s.strip().lower() != "source"]
        if not candidate_sheets:
            messagebox.showinfo("No recipes found", "This workbook doesn't have any recipe sheets to load.")
            return

        chosen_sheet = self._ask_pick_sheet(candidate_sheets)
        if not chosen_sheet:
            return

        if self.recipe_rows:
            if not messagebox.askyesno(
                "Replace current recipe?",
                "This will replace the ingredients currently in your recipe list. Continue?",
            ):
                return

        loaded_rows, servings_found, unmatched = self._parse_recipe_sheet(wb[chosen_sheet])

        if not loaded_rows:
            messagebox.showwarning("Nothing to load", f"Couldn't find any ingredient rows in '{chosen_sheet}'.")
            return

        self.recipe_rows = loaded_rows
        self._refresh_table()
        if servings_found:
            self.servings_var.set(f"{servings_found:g}")
            self._refresh_totals()

        msg = f"Loaded {len(loaded_rows)} ingredient(s) from '{chosen_sheet}'."
        if unmatched:
            msg += (
                f"\n\n{len(unmatched)} ingredient(s) weren't found in your current ingredient list "
                "(maybe renamed or removed since this recipe was saved), so their old saved values "
                "were kept as-is instead of being recalculated:\n- " + "\n- ".join(unmatched)
            )
        messagebox.showinfo("Recipe loaded", msg)

    def _save_recipe(self):
        if not self.recipe_rows:
            messagebox.showwarning("Nothing to save", "Add at least one ingredient first.")
            return

        if not self.workbook_path:
            choice = self._ask_new_or_existing()
            if choice == "new":
                self._create_new_workbook()
            elif choice == "existing":
                self._choose_existing_workbook()
            if not self.workbook_path:
                return

        default_name = f"Recipe {datetime.now().strftime('%b %d')}"
        dialog = ctk.CTkInputDialog(
            title="Recipe name",
            text=f"Name this recipe (becomes the sheet name).\nE.g. '{default_name}':",
        )
        sheet_name = dialog.get_input()
        if not sheet_name:
            return
        sheet_name = sheet_name[:31]  # Excel sheet name limit

        try:
            import openpyxl
        except ImportError:
            messagebox.showerror(
                "Missing dependency", "This feature needs openpyxl. Install it with:\n\npip install openpyxl"
            )
            return

        try:
            if os.path.exists(self.workbook_path):
                wb = openpyxl.load_workbook(self.workbook_path)
            else:
                # First-ever save to a brand-new workbook path.
                wb = openpyxl.Workbook()
            if sheet_name in wb.sheetnames:
                if not messagebox.askyesno(
                    "Sheet exists", f"A sheet named '{sheet_name}' already exists. Overwrite it?"
                ):
                    return
                del wb[sheet_name]
            ws = wb.create_sheet(sheet_name)
            # openpyxl gives new workbooks one blank default sheet called "Sheet" —
            # drop it now that we have a real sheet, so we don't leave clutter behind.
            if "Sheet" in wb.sheetnames and wb["Sheet"].max_row == 1 and wb["Sheet"].max_column == 1 \
                    and wb["Sheet"]["A1"].value is None and len(wb.sheetnames) > 1:
                del wb["Sheet"]

            headers = ["Food", "You Entered", "Grams total"] + [label for _, label in MACRO_FIELDS]
            ws.append(headers)
            header_row_1 = ws.max_row
            macro_headers_only = ["", "", ""] + [label for _, label in MACRO_FIELDS]

            try:
                servings = float(self.servings_var.get())
                if servings <= 0:
                    servings = 1
            except ValueError:
                servings = 1

            totals = {key: 0.0 for key, _ in MACRO_FIELDS}
            for row in self.recipe_rows:
                ws.append(
                    [row["name"], row["entered"], round(row["grams"], 2)]
                    + [round(row["macros"][key], 2) for key, _ in MACRO_FIELDS]
                )
                for key, _ in MACRO_FIELDS:
                    totals[key] += row["macros"][key]

            ws.append([])
            ws.append(macro_headers_only)  # repeat headers so they're visible above the summary too
            header_row_2 = ws.max_row
            ws.append(["TOTAL", "", ""] + [round(totals[key], 2) for key, _ in MACRO_FIELDS])
            ws.append(["Number of servings", "", "", servings])
            ws.append(
                ["Nutrition per serving", "", ""]
                + [round(totals[key] / servings, 2) for key, _ in MACRO_FIELDS]
            )

            bold = openpyxl.styles.Font(bold=True)
            for header_row in (header_row_1, header_row_2):
                for col_idx in range(1, len(headers) + 1):
                    ws.cell(row=header_row, column=col_idx).font = bold

            for col_idx in range(1, len(headers) + 1):
                ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = 16

            wb.save(self.workbook_path)
            self.workbook_label_var.set(f"Workbook: {self.workbook_path}")
            messagebox.showinfo("Saved", f"Recipe saved as sheet '{sheet_name}' in your workbook.")
        except Exception as exc:
            messagebox.showerror("Save failed", str(exc))


if __name__ == "__main__":
    app = MacroCalculatorApp()
    app.mainloop()
