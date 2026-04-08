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
            self._log(f"  Sheets: {', '.join(xl.sheet_names)}")

            landing_sheet = bar_sheet = None
            for s in xl.sheet_names:
                key = s.lower().replace("_", "").replace(" ", "")
                if "landing" in key and "offset" in key:
                    landing_sheet = s
                elif "bar" in key and "offset" in key:
                    bar_sheet = s

            landing_ids: list[int] = []
            bar_ids: list[int] = []

            if landing_sheet:
                self._log(f"\n  Reading '{landing_sheet}'…")
                df = pd.read_excel(xl, sheet_name=landing_sheet)
                landing_ids = df.iloc[:, 0].dropna().astype(int).tolist()
                self._log(f"    {len(landing_ids)} landing element IDs found")

            if bar_sheet:
                self._log(f"\n  Reading '{bar_sheet}'…")
                df = pd.read_excel(xl, sheet_name=bar_sheet)
                bar_ids = df.iloc[:, 0].dropna().astype(int).tolist()
                self._log(f"    {len(bar_ids)} bar element IDs found")

            # --- Read BDF ---
            self._log(f"\n{sep}")
            self._log("Reading BDF with pyNastran…")
            bdf = BDF(debug=False)
            bdf_path = self.offset_input_bdf.get()
            try:
                bdf.read_bdf(
                    bdf_path, validate=False, xref=False,
                    read_includes=True, encoding="latin-1"
                )
            except Exception:
                self._log("  Standard read failed — retrying with punch=True…")
                bdf = BDF(debug=False)
                bdf.read_bdf(
                    bdf_path, validate=False, xref=False,
                    read_includes=True, encoding="latin-1", punch=True
                )

            self._log(f"  Nodes:      {len(bdf.nodes)}")
            self._log(f"  Elements:   {len(bdf.elements)}")
            self._log(f"  Properties: {len(bdf.properties)}")

            # --- Landing offsets ---
            self._log(f"\n{sep}")
            self._log("Calculating landing offsets…")

            landing_results = []
            landing_thickness: dict[int, float] = {}
            landing_normals: dict[int, np.ndarray] = {}

            for eid in landing_ids:
                if eid not in bdf.elements:
                    continue
                elem = bdf.elements[eid]
                if not (hasattr(elem, "pid") and elem.pid in bdf.properties):
                    continue
                prop = bdf.properties[elem.pid]

                thickness = None
                if hasattr(prop, "t"):                   # PSHELL
                    thickness = prop.t
                elif hasattr(prop, "total_thickness"):   # PCOMP
                    thickness = prop.total_thickness()

                if not thickness:
                    continue

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
                    n_ids = elem.node_ids[:4] if elem.type.startswith("CQUAD") else elem.node_ids[:3]
                    nodes = [bdf.nodes[nid] for nid in n_ids if nid in bdf.nodes]
                    if len(nodes) >= 3:
                        p1 = np.array(nodes[0].xyz)
                        p2 = np.array(nodes[1].xyz)
                        p3 = np.array(nodes[2].xyz)
                        normal = np.cross(p2 - p1, p3 - p1)
                        norm_len = np.linalg.norm(normal)
                        if norm_len > 1e-10:
                            landing_normals[eid] = normal / norm_len

            self._log(f"  {len(landing_results)} landing elements processed")

            # --- Node-to-shell mapping ---
            self._log(f"\n{sep}")
            self._log("Building node-to-shell map…")
            node_to_shells: dict[int, list[int]] = {}
            for eid, elem in bdf.elements.items():
                if elem.type in ("CQUAD4", "CTRIA3", "CQUAD8", "CTRIA6"):
                    for nid in elem.node_ids:
                        node_to_shells.setdefault(nid, []).append(eid)
            self._log(f"  {len(node_to_shells)} nodes mapped")

            # --- Bar offsets ---
            self._log(f"\n{sep}")
            self._log("Calculating bar offsets…")

            bar_results = []
            skipped = 0

            for eid in bar_ids:
                if eid not in bdf.elements:
                    continue
                elem = bdf.elements[eid]
                if elem.type != "CBAR" or not (
                    hasattr(elem, "pid") and elem.pid in bdf.properties
                ):
                    continue
                prop = bdf.properties[elem.pid]

                bar_t = None
                if prop.type == "PBARL":
                    if hasattr(prop, "dim") and prop.dim:
                        bar_t = prop.dim[0]
                elif prop.type == "PBAR":
                    if hasattr(prop, "A") and prop.A > 0:
                        bar_t = float(np.sqrt(prop.A))

                if not bar_t:
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
                    skipped += 1
                    continue

                magnitude = best_thick + bar_t / 2.0
                offset_vec = -best_normal * magnitude

                bar_results.append(
                    {
                        "Element_ID": eid,
                        "Element_Type": elem.type,
                        "Property_ID": elem.pid,
                        "Property_Type": prop.type,
                        "Bar_Thickness": bar_t,
                        "Connected_Landing_ID": best_landing_id,
                        "Landing_Thickness": best_thick,
                        "Offset_Magnitude": magnitude,
                        "Offset_X": offset_vec[0],
                        "Offset_Y": offset_vec[1],
                        "Offset_Z": offset_vec[2],
                    }
                )

            self._log(f"  {len(bar_results)} bar elements processed")
            if skipped:
                self._log(f"  {skipped} bars skipped (no landing connection found)")

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
                     "Bar_Thickness", "Connected_Landing_ID", "Landing_Thickness",
                     "Offset_Magnitude", "Offset_X", "Offset_Y", "Offset_Z"]
                )
                for r in bar_results:
                    writer.writerow(
                        [r["Element_ID"], r["Element_Type"], r["Property_ID"],
                         r["Property_Type"], r["Bar_Thickness"],
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
            landing_offsets: dict[int, float] = {}
            bar_offsets: dict[int, tuple] = {}

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
                            bar_offsets[int(row[0])] = (
                                float(row[8]), float(row[9]), float(row[10])
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
