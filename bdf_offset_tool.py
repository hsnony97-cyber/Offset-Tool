#!/usr/bin/env python3
"""
BDF Offset Tool v3.0
====================
Standalone tool for calculating and applying geometric offsets to
Nastran BDF files for modeling assembly gaps between shell (landing)
elements and bar elements.

Workflow:
  Step 1 - Calculate Offsets:
    - Select input BDF
    - Select Excel with element IDs (sheets: Landing_Offset, Bar_Offset)
    - Click CALCULATE OFFSETS -> saves offset CSV

  Step 2 - Apply Offsets:
    - Select offset CSV from Step 1
    - Click APPLY OFFSETS -> generates modified BDF

Offset Logic:
  Landing (shell): zoffset = -thickness / 2
  Bar:             wa = wb = offset_vector
                   magnitude = landing_thickness + bar_thickness / 2
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import csv
import threading

import numpy as np
import pandas as pd
from pyNastran.bdf.bdf import BDF


def _bar_local_y(elem, bdf_nodes):
    """
    Compute the Nastran CBAR local y-axis unit vector.

    Nastran convention:
      x_b = unit vector GA -> GB
      v   = orientation vector (G0 or elem.x field)
      z_b = normalize( x_b × v )
      y_b = z_b × x_b

    Returns None if the orientation cannot be determined (degenerate case).
    """
    try:
        ga_xyz = np.array(bdf_nodes[elem.nodes[0]].xyz, dtype=float)
        gb_xyz = np.array(bdf_nodes[elem.nodes[1]].xyz, dtype=float)
        x_b = gb_xyz - ga_xyz
        xnorm = np.linalg.norm(x_b)
        if xnorm < 1e-10:
            return None
        x_b = x_b / xnorm

        # Orientation vector: G0 grid point or explicit vector
        if elem.g0 is not None:
            v = np.array(bdf_nodes[elem.g0].xyz, dtype=float) - ga_xyz
        elif elem.x is not None:
            v = np.array(elem.x, dtype=float)
        else:
            return None

        # z_b = x_b × v  (cross product removes any component along bar)
        z_b = np.cross(x_b, v)
        znorm = np.linalg.norm(z_b)
        if znorm < 1e-10:
            return None   # orientation vector is parallel to bar axis
        z_b = z_b / znorm

        # y_b = z_b × x_b
        y_b = np.cross(z_b, x_b)
        ynorm = np.linalg.norm(y_b)
        if ynorm < 1e-10:
            return None
        return y_b / ynorm
    except Exception:
        return None


    """
    Robustly extract shell thickness from a pyNastran property object.
    Handles PSHELL (prop.t), PCOMP (total_thickness()), and per-element
    thickness defined on CQUAD4/CTRIA3 T1..T4 fields.
    Returns None if thickness cannot be determined.
    """
    # PSHELL
    if hasattr(prop, "t") and prop.t is not None and prop.t != 0.0:
        return float(prop.t)

    # PCOMP / PCOMPG
    if hasattr(prop, "total_thickness"):
        try:
            t = prop.total_thickness()
            if t is not None and t > 0:
                return float(t)
        except Exception:
            pass
    # Also try summing ply thicknesses directly
    if hasattr(prop, "thicknesses"):
        try:
            t = sum(prop.thicknesses)
            if t > 0:
                return float(t)
        except Exception:
            pass

    # PSHELL where t is not set but element has per-corner thicknesses
    if elem is not None and hasattr(elem, "T1"):
        try:
            vals = [v for v in (elem.T1, elem.T2, elem.T3, getattr(elem, "T4", None))
                    if v is not None and v > 0]
            if vals:
                return float(sum(vals) / len(vals))
        except Exception:
            pass

    return None


class BDFOffsetTool:
    def __init__(self, root):
        self.root = root
        self.root.title("BDF Offset Tool v3.0")
        self.root.geometry("900x850")
        self.root.resizable(True, True)

        # State variables
        self.offset_input_bdf = tk.StringVar()
        self.offset_element_excel = tk.StringVar()
        self.offset_csv_path = tk.StringVar()
        self.offset_output_name = tk.StringVar(value="offset_applied.bdf")
        self.offset_csv_name = tk.StringVar(value="calculated_offsets.csv")

        self._build_ui()

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------

    def _build_ui(self):
        main = ttk.Frame(self.root, padding="12")
        main.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            main, text="BDF Offset Tool v3.0", font=("Helvetica", 15, "bold")
        ).pack(pady=(0, 12))

        # Input BDF
        bf = ttk.LabelFrame(main, text="Input BDF File", padding="10")
        bf.pack(fill=tk.X, pady=4)
        ttk.Label(bf, text="BDF File:", width=12).grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(bf, textvariable=self.offset_input_bdf, width=65).grid(
            row=0, column=1, padx=5
        )
        ttk.Button(bf, text="Browse…", command=self._browse_bdf).grid(
            row=0, column=2, padx=4
        )

        # Element IDs Excel
        ef = ttk.LabelFrame(main, text="Element IDs (Excel)", padding="10")
        ef.pack(fill=tk.X, pady=4)
        ttk.Label(ef, text="Excel File:", width=12).grid(
            row=0, column=0, sticky=tk.W
        )
        ttk.Entry(ef, textvariable=self.offset_element_excel, width=65).grid(
            row=0, column=1, padx=5
        )
        ttk.Button(ef, text="Browse…", command=self._browse_excel).grid(
            row=0, column=2, padx=4
        )
        ttk.Label(
            ef,
            text="Required sheets: 'Landing_Offset' (Column A)  |  'Bar_Offset' (Column A)",
            font=("Helvetica", 9, "italic"),
        ).grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=(6, 0))

        # Step 1: Calculate Offsets
        cf = ttk.LabelFrame(main, text="Step 1 — Calculate Offsets", padding="10")
        cf.pack(fill=tk.X, pady=4)
        ttk.Label(cf, text="Output CSV name:", width=16).grid(
            row=0, column=0, sticky=tk.W
        )
        ttk.Entry(cf, textvariable=self.offset_csv_name, width=40).grid(
            row=0, column=1, sticky=tk.W, padx=5
        )
        self.btn_calc = ttk.Button(
            cf,
            text=">>> CALCULATE OFFSETS <<<",
            command=self._start_calculate,
        )
        self.btn_calc.grid(row=1, column=0, columnspan=3, pady=10)

        # Step 2: Apply Offsets
        af = ttk.LabelFrame(main, text="Step 2 — Apply Offsets from CSV", padding="10")
        af.pack(fill=tk.X, pady=4)
        ttk.Label(af, text="Offset CSV:", width=16).grid(
            row=0, column=0, sticky=tk.W
        )
        ttk.Entry(af, textvariable=self.offset_csv_path, width=65).grid(
            row=0, column=1, padx=5
        )
        ttk.Button(af, text="Browse…", command=self._browse_csv).grid(
            row=0, column=2, padx=4
        )
        ttk.Label(af, text="Output BDF name:", width=16).grid(
            row=1, column=0, sticky=tk.W, pady=(6, 0)
        )
        ttk.Entry(af, textvariable=self.offset_output_name, width=40).grid(
            row=1, column=1, sticky=tk.W, padx=5, pady=(6, 0)
        )
        self.btn_apply = ttk.Button(
            af,
            text=">>> APPLY OFFSETS <<<",
            command=self._start_apply,
        )
        self.btn_apply.grid(row=2, column=0, columnspan=3, pady=10)

        # Toolbar
        tb = ttk.Frame(main)
        tb.pack(fill=tk.X, pady=4)
        ttk.Button(tb, text="Clear Log", command=self._clear_log).pack(side=tk.LEFT)

        # Progress bar
        self.progress = ttk.Progressbar(main, mode="indeterminate")
        self.progress.pack(fill=tk.X, pady=4)

        # Log area
        lf = ttk.LabelFrame(main, text="Log / Status", padding="8")
        lf.pack(fill=tk.BOTH, expand=True, pady=4)
        self.log_text = scrolledtext.ScrolledText(lf, height=16, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # Startup message
        self._print_intro()

    def _print_intro(self):
        sep = "=" * 68
        lines = [
            sep,
            "BDF Offset Tool v3.0  —  PyNastran Method",
            sep,
            "",
            "STEP 1: Calculate Offsets",
            "  1. Select the input BDF file.",
            "  2. Select the Excel file containing element IDs.",
            "  3. Click CALCULATE OFFSETS  ->  saves a CSV with offsets.",
            "",
            "STEP 2: Apply Offsets",
            "  1. Select the offset CSV produced in Step 1.",
            "  2. Click APPLY OFFSETS  ->  writes a modified BDF.",
            "",
            "Offset logic:",
            "  Landing (shell):  ZOFFS = -thickness / 2",
            "  Bar:              WA = WB = offset_vector",
            "                    |offset| = landing_t + bar_t / 2",
            "",
            sep,
            "",
        ]
        for ln in lines:
            self._log(ln)

    # ------------------------------------------------------------------
    # Browse callbacks
    # ------------------------------------------------------------------

    def _browse_bdf(self):
        path = filedialog.askopenfilename(
            title="Select BDF File",
            filetypes=[("BDF / Nastran files", "*.bdf *.dat *.nas"), ("All files", "*.*")],
        )
        if path:
            self.offset_input_bdf.set(path)
            self._log(f"BDF selected: {os.path.basename(path)}")

    def _browse_excel(self):
        path = filedialog.askopenfilename(
            title="Select Excel File with Element IDs",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if path:
            self.offset_element_excel.set(path)
            self._log(f"Excel selected: {os.path.basename(path)}")

    def _browse_csv(self):
        path = filedialog.askopenfilename(
            title="Select Offset CSV File",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if path:
            self.offset_csv_path.set(path)
            self._log(f"CSV selected: {os.path.basename(path)}")

    # ------------------------------------------------------------------
    # Logging helpers
    # ------------------------------------------------------------------

    def _log(self, msg: str):
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def _clear_log(self):
        self.log_text.delete(1.0, tk.END)

    # ------------------------------------------------------------------
    # Step 1 — Calculate offsets
    # ------------------------------------------------------------------

    def _start_calculate(self):
        if not self.offset_input_bdf.get():
            messagebox.showerror("Missing input", "Please select a BDF file.")
            return
        if not self.offset_element_excel.get():
            messagebox.showerror("Missing input", "Please select the Excel file with element IDs.")
            return

        self.btn_calc.config(state=tk.DISABLED)
        self.progress.start()
        threading.Thread(target=self._calculate_offsets, daemon=True).start()

    def _calculate_offsets(self):
        """Background worker: parse BDF + Excel, compute offsets, write CSV."""
        try:
            sep = "=" * 68
            self._log(f"\n{sep}")
            self._log("STEP 1: CALCULATING OFFSETS")
            self._log(sep)

            # --- Read element IDs from Excel ---
            self._log("\nReading element IDs from Excel…")
            xl = pd.ExcelFile(self.offset_element_excel.get())
            self._log(f"  Sheets found: {', '.join(xl.sheet_names)}")

            # Sheet detection: primary (requires both keywords), then fallback (any match)
            landing_sheet = bar_sheet = None
            landing_sheet_fb = bar_sheet_fb = None  # fallback candidates

            for s in xl.sheet_names:
                key = s.lower().replace("_", "").replace(" ", "")
                if "landing" in key and "offset" in key:
                    landing_sheet = s
                elif "bar" in key and "offset" in key:
                    bar_sheet = s
                # fallbacks
                if landing_sheet_fb is None and "landing" in key:
                    landing_sheet_fb = s
                if bar_sheet_fb is None and "bar" in key and "landing" not in key:
                    bar_sheet_fb = s

            if landing_sheet is None and landing_sheet_fb is not None:
                landing_sheet = landing_sheet_fb
                self._log(f"  [!] No 'Landing_Offset' sheet — using fallback: '{landing_sheet}'")
            if bar_sheet is None and bar_sheet_fb is not None:
                bar_sheet = bar_sheet_fb
                self._log(f"  [!] No 'Bar_Offset' sheet — using fallback: '{bar_sheet}'")

            if landing_sheet is None:
                self._log("  [!] WARNING: No landing sheet detected! "
                          "Sheet must contain 'landing' in its name.")
            if bar_sheet is None:
                self._log("  [!] WARNING: No bar sheet detected! "
                          "Sheet must contain 'bar' in its name.")

            landing_ids = []
            bar_ids = []

            if landing_sheet:
                self._log(f"\n  Reading landing sheet: '{landing_sheet}'…")
                df = pd.read_excel(xl, sheet_name=landing_sheet)
                landing_ids = df.iloc[:, 0].dropna().astype(int).tolist()
                self._log(f"    {len(landing_ids)} landing element IDs read from Excel")
                if landing_ids:
                    sample = landing_ids[:5]
                    self._log(f"    First IDs: {sample}")

            bar_sections = {}   # {eid: 'C' or 'I'}

            if bar_sheet:
                self._log(f"\n  Reading bar sheet: '{bar_sheet}'…")
                df = pd.read_excel(xl, sheet_name=bar_sheet)
                df_bar = df.dropna(subset=[df.columns[0]])
                bar_ids = df_bar.iloc[:, 0].astype(int).tolist()

                # Column B = Section type ('C' or 'I'). Default to 'I' if absent.
                has_section_col = df_bar.shape[1] > 1
                for _, row in df_bar.iterrows():
                    eid = int(row.iloc[0])
                    if has_section_col and pd.notna(row.iloc[1]):
                        sec = str(row.iloc[1]).strip().upper()
                    else:
                        sec = "I"
                    bar_sections[eid] = sec

                c_count = sum(1 for s in bar_sections.values() if s == "C")
                i_count = sum(1 for s in bar_sections.values() if s == "I")
                self._log(f"    {len(bar_ids)} bar element IDs read from Excel")
                self._log(f"    Section types — I: {i_count}, C: {c_count}")
                if not has_section_col:
                    self._log("    [!] No Section column found — all treated as I-section")
                if bar_ids:
                    self._log(f"    First IDs: {bar_ids[:5]}")

            # --- Read BDF ---
            self._log(f"\n{sep}")
            self._log("Reading BDF with pyNastran…")
            bdf_path = self.offset_input_bdf.get()

            def _try_read(punch: bool) -> BDF:
                b = BDF(debug=False)
                b.read_bdf(
                    bdf_path, validate=False, xref=False,
                    read_includes=True, encoding="latin-1",
                    punch=punch,
                )
                return b

            bdf = BDF(debug=False)
            try:
                bdf = _try_read(punch=False)
                self._log("  BDF read OK (standard mode)")
            except Exception as e:
                self._log(f"  Standard read raised: {e}")
                # Keep whatever was partially read; if empty, retry with punch=True
                if len(bdf.elements) == 0:
                    self._log("  Retrying with punch=True (skips EXEC/CASE CONTROL)…")
                    try:
                        bdf = _try_read(punch=True)
                        self._log("  BDF read OK (punch mode)")
                    except Exception as e2:
                        self._log(f"  punch=True also raised: {e2} — using partial data")

            self._log(f"  Nodes:      {len(bdf.nodes)}")
            self._log(f"  Elements:   {len(bdf.elements)}")
            self._log(f"  Properties: {len(bdf.properties)}")

            # Sanity-check: show element types present in BDF
            from collections import Counter
            type_counts = Counter(e.type for e in bdf.elements.values())
            self._log(f"  Element types: " +
                      ", ".join(f"{t}={n}" for t, n in
                                sorted(type_counts.items(), key=lambda x: -x[1])[:8]))

            # --- Landing offsets ---
            self._log(f"\n{sep}")
            self._log("Calculating landing offsets…")

            landing_results = []
            landing_thickness = {}
            landing_normals = {}

            # Diagnostic counters
            dbg_not_in_bdf = 0
            dbg_no_pid = 0
            dbg_pid_not_in_props = 0
            dbg_no_thickness = 0
            dbg_ok = 0
            _first_missing = []       # up to 5 IDs not found in BDF
            _first_no_thick = []      # up to 5 (eid, prop_type, t_val)

            for eid in landing_ids:
                if eid not in bdf.elements:
                    dbg_not_in_bdf += 1
                    if len(_first_missing) < 5:
                        _first_missing.append(eid)
                    continue
                elem = bdf.elements[eid]
                if not hasattr(elem, "pid"):
                    dbg_no_pid += 1
                    continue
                if elem.pid not in bdf.properties:
                    dbg_pid_not_in_props += 1
                    continue
                prop = bdf.properties[elem.pid]

                thickness = _get_thickness(prop, elem)

                if not thickness:
                    dbg_no_thickness += 1
                    if len(_first_no_thick) < 5:
                        raw = getattr(prop, "t", "n/a")
                        _first_no_thick.append(
                            f"eid={eid} prop_type={prop.type} prop.t={raw}"
                        )
                    continue

                dbg_ok += 1
                zoffset = -thickness / 2.0
                landing_thickness[eid] = thickness

                landing_results.append(
                    {
                        "Element_ID": eid,
                        "Element_Type": elem.type,
                        "Property_ID": elem.pid,
                        "Property_Type": prop.type,
                        "Thickness": thickness,
                        "Zoffset": zoffset,
                    }
                )

                # Surface normal (for bar offset direction)
                if elem.type in ("CQUAD4", "CTRIA3", "CQUAD8", "CTRIA6"):
                    n_ids = (
                        elem.node_ids[:4]
                        if elem.type.startswith("CQUAD")
                        else elem.node_ids[:3]
                    )
                    nodes = [bdf.nodes[nid] for nid in n_ids if nid in bdf.nodes]
                    if len(nodes) >= 3:
                        p1 = np.array(nodes[0].xyz)
                        p2 = np.array(nodes[1].xyz)
                        p3 = np.array(nodes[2].xyz)
                        normal = np.cross(p2 - p1, p3 - p1)
                        norm_len = np.linalg.norm(normal)
                        if norm_len > 1e-10:
                            landing_normals[eid] = normal / norm_len

            # --- Diagnostic report ---
            self._log(f"\n  Landing element diagnostics:")
            self._log(f"    IDs from Excel           : {len(landing_ids)}")
            self._log(f"    Not found in BDF         : {dbg_not_in_bdf}")
            if _first_missing:
                self._log(f"      First missing IDs      : {_first_missing}")
            self._log(f"    No pid attribute         : {dbg_no_pid}")
            self._log(f"    PID not in properties    : {dbg_pid_not_in_props}")
            self._log(f"    Zero/None thickness      : {dbg_no_thickness}")
            if _first_no_thick:
                for info in _first_no_thick:
                    self._log(f"      {info}")
            self._log(f"    Successfully processed   : {dbg_ok}")
            self._log(f"  {len(landing_results)} landing elements processed")

            # --- Node-to-shell mapping ---
            self._log(f"\n{sep}")
            self._log("Building node-to-shell map…")
            node_to_shells = {}
            for eid, elem in bdf.elements.items():
                if elem.type in ("CQUAD4", "CTRIA3", "CQUAD8", "CTRIA6"):
                    for nid in elem.node_ids:
                        node_to_shells.setdefault(nid, []).append(eid)
            self._log(f"  {len(node_to_shells)} nodes mapped")

            # --- Bar offsets ---
            self._log(f"\n{sep}")
            self._log("Calculating bar offsets…")

            bar_results = []
            bar_not_in_bdf = 0
            bar_no_prop = 0
            bar_no_thickness = 0
            bar_no_landing = 0

            for eid in bar_ids:
                if eid not in bdf.elements:
                    bar_not_in_bdf += 1
                    continue
                elem = bdf.elements[eid]
                if elem.type != "CBAR" or not (
                    hasattr(elem, "pid") and elem.pid in bdf.properties
                ):
                    bar_no_prop += 1
                    continue
                prop = bdf.properties[elem.pid]

                bar_t = None
                if prop.type == "PBARL":
                    # DIM2 (index 1) = section height used for Y-offset
                    if hasattr(prop, "dim") and len(prop.dim) > 1:
                        bar_t = float(prop.dim[1])
                    elif hasattr(prop, "dim") and prop.dim:
                        bar_t = float(prop.dim[0])  # fallback to DIM1
                elif prop.type == "PBAR":
                    if hasattr(prop, "A") and prop.A > 0:
                        bar_t = float(np.sqrt(prop.A))

                if not bar_t:
                    bar_no_thickness += 1
                    continue

                n1, n2 = elem.node_ids[:2]
                shells_n1 = set(node_to_shells.get(n1, []))
                shells_n2 = set(node_to_shells.get(n2, []))
                connected = shells_n1 & shells_n2

                # Pick connected landing shell with greatest thickness
                best_thick = 0.0
                best_normal = None
                best_landing_id = None
                for sid in connected:
                    if sid in landing_thickness:
                        t = landing_thickness[sid]
                        if t > best_thick:
                            best_thick = t
                            best_landing_id = sid
                            best_normal = landing_normals.get(sid)

                if best_normal is None or best_thick == 0:
                    bar_no_landing += 1
                    continue

                section = bar_sections.get(eid, "I")
                y_local = _bar_local_y(elem, bdf.nodes)

                if y_local is None:
                    # No orientation defined — fall back to landing normal for both types
                    magnitude = best_thick / 2.0 + bar_t / 2.0
                    offset_vec = -best_normal * magnitude
                    self._log(
                        f"  [!] eid={eid} ({section}): no bar orientation "
                        f"— using landing normal fallback"
                    )
                elif section == "I":
                    # I-beam: bottom flange (cap) at shell outer surface,
                    # neutral axis at bar_height/2 above it.
                    # offset = landing_t/2 (to shell outer surface)
                    #        + bar_dim2/2 (to neutral axis / web centre)
                    magnitude = best_thick / 2.0 + bar_t / 2.0
                    offset_vec = y_local * magnitude
                else:
                    # C-section: web face sits at shell outer surface.
                    # offset = landing_t/2 only (neutral axis at web junction).
                    magnitude = best_thick / 2.0
                    offset_vec = y_local * magnitude

                bar_results.append(
                    {
                        "Element_ID": eid,
                        "Element_Type": elem.type,
                        "Property_ID": elem.pid,
                        "Property_Type": prop.type,
                        "Section": section,
                        "Bar_Thickness": bar_t,
                        "Connected_Landing_ID": best_landing_id,
                        "Landing_Thickness": best_thick,
                        "Offset_Magnitude": magnitude,
                        "Offset_X": offset_vec[0],
                        "Offset_Y": offset_vec[1],
                        "Offset_Z": offset_vec[2],
                    }
                )

            self._log(f"\n  Bar element diagnostics:")
            self._log(f"    IDs from Excel           : {len(bar_ids)}")
            self._log(f"    Not found in BDF         : {bar_not_in_bdf}")
            self._log(f"    Not CBAR / no property   : {bar_no_prop}")
            self._log(f"    Zero/None bar thickness  : {bar_no_thickness}")
            self._log(f"    No landing connection    : {bar_no_landing}")
            self._log(f"    Successfully processed   : {len(bar_results)}")
            self._log(f"  {len(bar_results)} bar elements processed")

            # --- Write CSV ---
            self._log(f"\n{sep}")
            self._log("Writing CSV…")

            output_dir = os.path.dirname(bdf_path)
            csv_path = os.path.join(output_dir, self.offset_csv_name.get())

            with open(csv_path, "w", newline="") as fh:
                writer = csv.writer(fh)

                writer.writerow(["LANDING OFFSETS"])
                writer.writerow(
                    ["Element_ID", "Element_Type", "Property_ID",
                     "Property_Type", "Thickness", "Zoffset"]
                )
                for r in landing_results:
                    writer.writerow(
                        [r["Element_ID"], r["Element_Type"], r["Property_ID"],
                         r["Property_Type"], r["Thickness"], r["Zoffset"]]
                    )

                writer.writerow([])

                writer.writerow(["BAR OFFSETS"])
                writer.writerow(
                    ["Element_ID", "Element_Type", "Property_ID", "Property_Type",
                     "Section", "Bar_Thickness", "Connected_Landing_ID",
                     "Landing_Thickness", "Offset_Magnitude",
                     "Offset_X", "Offset_Y", "Offset_Z"]
                )
                for r in bar_results:
                    writer.writerow(
                        [r["Element_ID"], r["Element_Type"], r["Property_ID"],
                         r["Property_Type"], r["Section"], r["Bar_Thickness"],
                         r["Connected_Landing_ID"], r["Landing_Thickness"],
                         r["Offset_Magnitude"], r["Offset_X"],
                         r["Offset_Y"], r["Offset_Z"]]
                    )

            size_kb = os.path.getsize(csv_path) / 1024
            self._log(f"  Saved: {csv_path}  ({size_kb:.1f} KB)")

            self._log(f"\n{sep}")
            self._log("CALCULATION COMPLETE")
            self._log(f"  Landing elements : {len(landing_results)}")
            self._log(f"  Bar elements     : {len(bar_results)}")
            self._log(f"  Output CSV       : {os.path.basename(csv_path)}")
            self._log(sep)

            # Pre-fill Step 2 CSV path
            self.offset_csv_path.set(csv_path)

            self.root.after(
                0,
                lambda: messagebox.showinfo(
                    "Step 1 complete",
                    f"Offsets calculated.\n\n"
                    f"Landing : {len(landing_results)}\n"
                    f"Bar     : {len(bar_results)}\n\n"
                    f"CSV: {os.path.basename(csv_path)}",
                ),
            )

        except Exception as exc:
            import traceback
            self._log(f"\nERROR: {exc}")
            self._log(traceback.format_exc())
            self.root.after(0, lambda: messagebox.showerror("Error", str(exc)))
        finally:
            self.root.after(
                0,
                lambda: [self.progress.stop(), self.btn_calc.config(state=tk.NORMAL)],
            )

    # ------------------------------------------------------------------
    # Step 2 — Apply offsets
    # ------------------------------------------------------------------

    def _start_apply(self):
        if not self.offset_input_bdf.get():
            messagebox.showerror("Missing input", "Please select a BDF file.")
            return
        if not self.offset_csv_path.get():
            messagebox.showerror("Missing input", "Please select the offset CSV file.")
            return

        self.btn_apply.config(state=tk.DISABLED)
        self.progress.start()
        threading.Thread(target=self._apply_offsets, daemon=True).start()

    def _apply_offsets(self):
        """Background worker: read CSV offsets, patch BDF text, write output."""
        try:
            sep = "=" * 68
            self._log(f"\n{sep}")
            self._log("STEP 2: APPLYING OFFSETS (text-based, preserves INCLUDEs)")
            self._log(sep)

            # --- Read CSV ---
            self._log("\nReading offset CSV…")
            landing_offsets = {}
            bar_offsets = {}

            with open(self.offset_csv_path.get(), "r") as fh:
                reader = csv.reader(fh)
                section = None
                for row in reader:
                    if not row or not row[0]:
                        continue
                    if "LANDING OFFSETS" in row[0]:
                        section = "landing"
                        next(reader)   # skip header
                        continue
                    elif "BAR OFFSETS" in row[0]:
                        section = "bar"
                        next(reader)   # skip header
                        continue

                    if section == "landing":
                        try:
                            landing_offsets[int(row[0])] = float(row[5])
                        except (ValueError, IndexError):
                            pass
                    elif section == "bar":
                        try:
                            # Columns: ID, Type, PID, PropType, Section, BarT,
                            #          LandingID, LandingT, Magnitude, X, Y, Z
                            bar_offsets[int(row[0])] = (
                                float(row[9]), float(row[10]), float(row[11])
                            )
                        except (ValueError, IndexError):
                            pass

            self._log(f"  {len(landing_offsets)} landing offsets loaded")
            self._log(f"  {len(bar_offsets)} bar offsets loaded")

            # --- Field formatter (8-char fixed width) ---
            def fmt(value, width=8) -> str:
                if isinstance(value, float):
                    s = f"{value:.4f}"
                    if len(s) > width:
                        s = f"{value:.2E}"
                    return s[:width].ljust(width)
                return str(value)[:width].ljust(width)

            # --- Read BDF as raw text ---
            self._log(f"\n{sep}")
            self._log("Reading BDF as text…")
            bdf_path = self.offset_input_bdf.get()
            with open(bdf_path, "r", encoding="latin-1") as fh:
                lines = fh.readlines()
            self._log(f"  {len(lines)} lines read")

            # --- Patch lines ---
            new_lines = []
            i = 0
            land_mod = bar_mod = 0

            while i < len(lines):
                line = lines[i]

                if line.startswith("CQUAD4"):
                    try:
                        eid = int(line[8:16].strip())
                        if eid in landing_offsets:
                            zoff = landing_offsets[eid]
                            if len(line) >= 64:
                                tail = line[72:] if len(line) > 72 else "\n"
                                new_lines.append(line[:64] + fmt(zoff) + tail)
                            else:
                                new_lines.append(line.rstrip().ljust(64) + fmt(zoff) + "\n")
                            land_mod += 1
                            i += 1
                            continue
                    except (ValueError, IndexError):
                        pass
                    new_lines.append(line)
                    i += 1

                elif line.startswith("CBAR"):
                    try:
                        eid = int(line[8:16].strip())
                        if eid in bar_offsets:
                            ox, oy, oz = bar_offsets[eid]
                            has_cont = (
                                i + 1 < len(lines)
                                and lines[i + 1][:1] in ("+", "*", " ")
                            )
                            new_lines.append(line)
                            if has_cont:
                                c = lines[i + 1]
                                new_cont = (
                                    c[:24]
                                    + fmt(ox) + fmt(oy) + fmt(oz)
                                    + fmt(ox) + fmt(oy) + fmt(oz)
                                    + "\n"
                                )
                                new_lines.append(new_cont)
                                bar_mod += 1
                                i += 2
                            else:
                                tag = "+CB" + str(eid)[-4:]
                                # Append continuation marker to CBAR line
                                new_lines[-1] = line.rstrip() + tag + "\n"
                                new_cont = (
                                    tag.ljust(8)
                                    + "        "
                                    + "        "
                                    + fmt(ox) + fmt(oy) + fmt(oz)
                                    + fmt(ox) + fmt(oy) + fmt(oz)
                                    + "\n"
                                )
                                new_lines.append(new_cont)
                                bar_mod += 1
                                i += 1
                            continue
                    except (ValueError, IndexError):
                        pass
                    new_lines.append(line)
                    i += 1

                else:
                    new_lines.append(line)
                    i += 1

            self._log(f"\n  CQUAD4 ZOFFS applied : {land_mod}")
            self._log(f"  CBAR  WA/WB applied  : {bar_mod}")

            # --- Write output BDF ---
            output_dir = os.path.dirname(bdf_path)
            output_path = os.path.join(output_dir, self.offset_output_name.get())
            self._log(f"\nWriting output BDF…")
            with open(output_path, "w", encoding="latin-1") as fh:
                fh.writelines(new_lines)

            size_mb = os.path.getsize(output_path) / (1024 * 1024)
            self._log(f"  Saved: {output_path}  ({size_mb:.2f} MB)")

            self._log(f"\n{sep}")
            self._log("APPLICATION COMPLETE")
            self._log(f"  Landing (ZOFFS) : {land_mod}")
            self._log(f"  Bar (WA/WB)     : {bar_mod}")
            self._log(f"  Output BDF      : {os.path.basename(output_path)}")
            self._log("  INCLUDEs preserved (text-based method)")
            self._log(sep)

            self.root.after(
                0,
                lambda: messagebox.showinfo(
                    "Step 2 complete",
                    f"Offsets applied.\n\n"
                    f"Landing (ZOFFS) : {land_mod}\n"
                    f"Bar (WA/WB)     : {bar_mod}\n\n"
                    f"Output: {os.path.basename(output_path)}\n"
                    f"INCLUDEs preserved.",
                ),
            )

        except Exception as exc:
            import traceback
            self._log(f"\nERROR: {exc}")
            self._log(traceback.format_exc())
            self.root.after(0, lambda: messagebox.showerror("Error", str(exc)))
        finally:
            self.root.after(
                0,
                lambda: [self.progress.stop(), self.btn_apply.config(state=tk.NORMAL)],
            )


# ----------------------------------------------------------------------

def main():
    root = tk.Tk()
    BDFOffsetTool(root)
    root.mainloop()


if __name__ == "__main__":
    main()
