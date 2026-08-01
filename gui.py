import os
import tkinter as tk
from tkinter import filedialog, messagebox

from logic import (
    build_comment_dict_from_source,
    build_floor_tag_index,
    plan_placements,
    write_missing_tags_json,
    DuplicateTagError,
    TagColumnNotFoundError,
)

from xlwings_comment import insert_comments_batch


class STAFLayoutAnnotatorApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("STAF Layout Annotator (Tag-based, xlwings-safe)")
        # Wider GUI so logs are readable
        self.root.geometry("980x640")

        self.source_path = None
        self.target_path = None

        self.comment_dict = None
        self.floor_index = None
        self.plan_result = None
        self.placements = None
        self.missing_json_path = None

        self._build_ui()

    def _build_ui(self):
        # Source file
        tk.Label(self.root, text="Machine_Details.xls (source):").grid(row=0, column=0, sticky="w", padx=10, pady=(12, 6))
        tk.Button(self.root, text="Browse...", command=self._pick_source, width=18).grid(row=0, column=1, sticky="w", padx=10, pady=(12, 6))

        # Target file
        tk.Label(self.root, text="STAF.xlsm (target):").grid(row=1, column=0, sticky="w", padx=10, pady=6)
        tk.Button(self.root, text="Browse...", command=self._pick_target, width=18).grid(row=1, column=1, sticky="w", padx=10, pady=6)

        # Buttons
        tk.Button(self.root, text="1) Run Logic (build + scan + plan)", command=self._run_logic, width=34)\
            .grid(row=2, column=0, columnspan=2, pady=(14, 6))

        tk.Button(self.root, text="2) Insert Notes (xlwings, preserve drawings)", command=self._insert_notes, width=34)\
            .grid(row=3, column=0, columnspan=2, pady=6)

        # Log box (bigger)
        self.log_box = tk.Text(self.root, height=24, width=112)
        self.log_box.grid(row=4, column=0, columnspan=2, padx=10, pady=10)

        # Status
        self.status_label = tk.Label(self.root, text="🚀 Ready.", anchor="w", relief="sunken")
        self.status_label.grid(row=5, column=0, columnspan=2, sticky="we", padx=10, pady=(0, 10))

        # Stretch
        self.root.grid_columnconfigure(0, weight=0)
        self.root.grid_columnconfigure(1, weight=1)

    def _set_status(self, msg: str):
        self.status_label.config(text=msg)
        self.status_label.update_idletasks()

    def _log(self, msg: str):
        self.log_box.insert(tk.END, msg + "\n")
        self.log_box.see(tk.END)
        self.log_box.update_idletasks()

    def _pick_source(self):
        p = filedialog.askopenfilename(
            title="Select Machine_Details source file",
            filetypes=[("Excel Files", "*.xls *.xlsx *.xlsm")]
        )
        if p:
            self.source_path = p
            self._log(f"✔ Source: {p}")

    def _pick_target(self):
        p = filedialog.askopenfilename(
            title="Select STAF target file",
            filetypes=[("Excel Macro-Enabled Workbook", "*.xlsm"), ("Excel Files", "*.xlsx *.xlsm")]
        )
        if p:
            self.target_path = p
            self._log(f"✔ Target: {p}")

    def _run_logic(self):
        try:
            self._set_status("⏳ Running logic: build source dict, scan FLOOR PLAN, plan placements...")

            if not self.source_path or not self.target_path:
                raise ValueError("Please select both Source (Machine_Details) and Target (STAF.xlsm).")

            # 1) Build source dict
            self._log("— Building comment dictionary from source (Tag as key)...")
            self.comment_dict = build_comment_dict_from_source(self.source_path, log_callback=self._log)

            # 2) Scan FLOOR PLAN once
            self._log("— Scanning FLOOR PLAN once to index all tags (merged-aware)...")
            #
            allowed = set(self.comment_dict.keys())
            self.floor_index = build_floor_tag_index(
                self.target_path,
                allowed_tags=allowed,
                sheet_name="FLOOR PLAN",
                log_callback=self._log
            )

            # 3) Plan placements
            self._log("— Planning placements...")
            self.plan_result = plan_placements(self.comment_dict, self.floor_index, log_callback=self._log)
            self.placements = self.plan_result.placements

            # 4) Write missing JSON (same folder as STAF)
            if self.plan_result.missing_tags:
                self.missing_json_path = write_missing_tags_json(self.plan_result.missing_tags, self.target_path)
                self._log(f"🧾 Missing tags JSON written: {self.missing_json_path}")
            else:
                self.missing_json_path = None
                self._log("✅ No missing tags. (All source tags found on FLOOR PLAN)")

            self._log("=== LOGIC SUMMARY ===")
            self._log(f"Source unique tags: {self.plan_result.total_source_tags}")
            self._log(f"Target indexed tags: {self.plan_result.total_target_tags}")
            self._log(f"Planned placements : {len(self.placements)}")
            self._log(f"Missing tags       : {len(self.plan_result.missing_tags)}")

            self._set_status("✅ Logic complete. Ready to insert notes via xlwings.")
        except (DuplicateTagError, TagColumnNotFoundError) as e:
            messagebox.showerror("Tag Error", str(e))
            self._log(str(e))
            self._set_status("❌ Stopped due to Tag rule violation.")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self._log(str(e))
            self._set_status("❌ Error while running logic.")

    def _insert_notes(self):
        try:
            self._set_status("⏳ Inserting notes via xlwings (always replace, preserve drawings)...")

            if not self.target_path:
                raise ValueError("Please select a STAF target file first.")
            if not self.placements:
                raise ValueError("Run Logic first. No placements are available to write.")

            # Save output in same folder as STAF (default out path does this already)
            summary = insert_comments_batch(
                in_path=self.target_path,
                sheet_name="FLOOR PLAN",
                placements=self.placements,
                out_path=None,           # creates *_with_Note.xlsm in same folder
                make_visible=False,
                autosize=True,
            )

            self._log("=== INSERT SUMMARY ===")
            for k, v in summary.items():
                self._log(f"{k}: {v}")

            ok = "OK" if summary.get("shapes_intact") else "WARNING"
            self._set_status(f"✅ Done. Notes written={summary.get('written')}, errors={summary.get('errors')}, Shapes={ok}")

        except Exception as e:
            messagebox.showerror("Error", str(e))
            self._log(str(e))
            self._set_status("❌ Error during insertion.")


if __name__ == "__main__":
    root = tk.Tk()
    app = STAFLayoutAnnotatorApp(root)
    root.mainloop()
