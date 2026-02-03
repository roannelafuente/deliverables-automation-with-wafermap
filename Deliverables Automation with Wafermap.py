import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import openpyxl
import csv
import os
import xlwings as xw
from datetime import datetime
import random

# Deliverables Automation Tool with Wafermap
# Author: Rose Anne Lafuente
# Licensed Electronics Engineer | Product Engineer II | Python Automation
# Description: Automates CSV-to-Excel workflows with pivot tables, custom formatting, End Test validation, 
# and wafermap visualization for yield and defect tracking. Built with Python, Tkinter, OpenPyXL, and xlwings.

class AutomatingDeliverables:
    def __init__(self, root):
        self.root = root
        self.root.title("Automating Deliverables")
        self.root.geometry("800x500")

        # Professional Neutral Theme
        self.bg_color = "#f5f5f5"
        self.fg_color = "#222222"
        self.entry_bg = "#ffffff"
        self.btn_bg = "#e0e0e0"
        self.btn_active = "#BEE395"

        self.root.configure(bg=self.bg_color)

        self.path_var = tk.StringVar()

        self.create_file_selection_frame()

        # Show filter selector immediately (empty at first)
        self.create_filter_selector([])

        # Then status box below it
        self.create_status_box()
        self.create_exit_button()


    def create_file_selection_frame(self):
        # File Selection frame with subtle border and spacing
        input_frame = tk.LabelFrame(
            self.root,
            text="File Selection",
            padx=10, pady=10,
            bd=2,
            relief="groove",
            font=("Segoe UI", 10, "bold")
        )
        input_frame.pack(fill="x", padx=15, pady=10)

        # Label inside the frame
        label = tk.Label(input_frame, text="Select CSV File:")
        label.pack(side="left", padx=(0, 10), pady=5)

        # Entry box inside the frame
        path_entry = tk.Entry(input_frame, textvariable=self.path_var,
                              bg="white", fg="black", insertbackground="black")
        path_entry.pack(side="left", padx=10, pady=5, fill="x", expand=True)

        # ✅ Convert button inside the same frame
        convert_btn = tk.Button(
            input_frame,
            text="Convert to Excel",
            width=18,
            command=self.convert_to_excel,
            bg=self.btn_bg,
            fg=self.fg_color,
            activebackground=self.btn_active
        )
        convert_btn.pack(side="right", padx=10, pady=5)
        # Browse button inside the frame
        browse_btn = tk.Button(input_frame, text="Browse", width=12, command=self.browse_file)
        browse_btn.pack(side="right", pady=5)


    def get_unique_c1_mark_values(raw_items):
        flat = []
        for item in raw_items:
            if isinstance(item, list):   # flatten nested lists
                flat.extend(item)
            elif item is not None:
                flat.append(item)

        # Strip whitespace but keep case
        cleaned = [str(i).strip() for i in flat if i]

        # Deduplicate while preserving order (case-sensitive)
        unique = list(dict.fromkeys(cleaned))
        return unique
    
    def create_filter_selector(self, items):
        # Pivot Filter Selection frame with subtle border and spacing
        filter_frame = tk.LabelFrame(
            self.root,
            text="Pivot Filter Selection",
            padx=10, pady=10,
            bd=2,
            relief="groove",
            font=("Segoe UI", 10, "bold")
        )

        filter_frame.pack(fill="x", padx=15, pady=10)

        tk.Label(filter_frame, text="Select C1_MARK:").pack(side="left", padx=5)

        # ✅ Clean and deduplicate items
        clean_items = [str(i) for i in items if i is not None]
        unique_items = list(dict.fromkeys(clean_items))

        self.filter_var = tk.StringVar()
        self.filter_dropdown = ttk.Combobox(
            filter_frame,
            textvariable=self.filter_var,
            values=unique_items,
            state="readonly",
            width=25
        )
        self.filter_dropdown.pack(side="left", padx=10)

        gen_pivot_btn = tk.Button(
            filter_frame,
            text="Generate Pivot Table",
            width=18,
            command=self.generate_pivot,
            bg=self.btn_bg,
            fg=self.fg_color,
            activebackground=self.btn_active
        )
        gen_pivot_btn.pack(side="left", padx=10)

        check_test_btn = tk.Button(
            filter_frame,
            text="Check End Test No",
            width=18,
            command=self.check_end_test,
            bg=self.btn_bg,
            fg=self.fg_color,
            activebackground=self.btn_active
        )
        check_test_btn.pack(side="left", padx=10)

        gen_wafermap_btn = tk.Button(
            filter_frame,
            text="Generate Wafermap",
            width=18,
            command=self.generate_wafermap,
            bg=self.btn_bg,
            fg=self.fg_color,
            activebackground=self.btn_active
        )
        gen_wafermap_btn.pack(side="left", padx=10)

    def create_status_box(self):
        status_frame = tk.LabelFrame(self.root, text="", padx=10, pady=10)
        status_frame.pack(fill="both", expand=True, padx=15, pady=10)

        self.status_box = tk.Text(
            status_frame,
            height=10,
            wrap="word",
            bg="white",
            fg="black",
            state="disabled"
        )
        self.status_box.pack(fill="both", expand=True)

    def create_exit_button(self):
        # Create an "invisible" frame with same background as root
        exit_frame = tk.Frame(self.root, bg=self.bg_color)
        exit_frame.pack(fill="x", side="bottom", padx=15, pady=5)

        # Place Exit button aligned right
        exit_btn = tk.Button(exit_frame, text="EXIT", width=12,
                             bg="#d32f2f", fg="white", command=self.root.destroy)
        exit_btn.pack(side="right", pady=10)

        clear_btn = tk.Button(exit_frame, text="Clear All", width=12,
                      command=self.clear_all,
                      bg=self.btn_bg, fg=self.fg_color, activebackground=self.btn_active)
        clear_btn.pack(side="right", padx=10)

    def show_status(self, message, color=None, clear=False):
        # Default to black unless explicitly set to red
        if color is None:
            color = "#000000"  # black

        self.status_box.config(state="normal")

        if clear:
            self.status_box.delete("1.0", "end")

        if message:
            self.status_box.insert("end", message + "\n")

            # Unique tag per line so colors don't overwrite
            line_tag = f"status_{self.status_box.index('end-2l')}"
            start_index = self.status_box.index("end-2l")
            end_index = self.status_box.index("end-1c")
            self.status_box.tag_add(line_tag, start_index, end_index)
            self.status_box.tag_config(line_tag, foreground=color)

        self.status_box.config(state="disabled")
        
    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select CSV File",
            filetypes=[("CSV files", "*.csv")]
        )
        if file_path:
            self.path_var.set(file_path)
            # Just show status that file is selected
            self.show_status(f"📂 Selected file:{file_path}", color="black")
            
    def convert_to_excel(self):
        file_path = self.path_var.get()
        if not file_path:
            self.show_status("⚠️ No file selected. Please browse for a CSV first.", color="#d32f2f")
            return

        try:
            # --- Convert CSV to Excel (vectorized) ---
            wb = openpyxl.Workbook()
            ws = wb.active

            sheet_name = os.path.splitext(os.path.basename(file_path))[0]
            ws.title = sheet_name[:31].replace(":", "_").replace("/", "_").replace("\\", "_")

            # Read CSV into list of lists
            with open(file_path, newline='', encoding='utf-8') as f:
                reader = csv.reader(f)
                rows = []
                for row in reader:
                    parsed = []
                    for value in row:
                        try:
                            if value.isdigit():
                                parsed.append(int(value))
                            else:
                                parsed.append(float(value))
                        except ValueError:
                            parsed.append(value)
                    rows.append(parsed)

            # Vectorized write: append all rows
            for r in rows:
                ws.append(r)

            out_file = os.path.splitext(file_path)[0] + ".xlsx"
            wb.save(out_file)
            wb.close()

            # --- Open with xlwings to read filter items ---
            app = xw.App(visible=False)
            wb_xlw = app.books.open(out_file)
            sht = wb_xlw.sheets[0]

            # Find header row in Column G
            first_row = sht.range("G1").end("down").row
            header_value = sht.range(f"G{first_row}").value
            if header_value != "C1_MARK":
                self.show_status("❌ First non-empty cell in Column G is not 'C1_MARK'.", color="#d32f2f")
                wb_xlw.close()
                app.quit()
                return

            # Collect filter items from C1_MARK column
            last_row = sht.range((first_row, 7)).end("down").row
            raw_items = sht.range((first_row+1, 7), (last_row, 7)).value

            # Deduplicate case-sensitive
            flat = [str(i).strip() for i in raw_items if i]
            unique_items = list(dict.fromkeys(flat))

            self.filter_dropdown['values'] = unique_items
            self.out_file = out_file
            self.base_name = sheet_name

            wb_xlw.close()
            app.quit()

            self.show_status(f"\n✅ Conversion complete: CSV → .xlsx\nFile saved at: {out_file}\n\nFilter options loaded.")

        except Exception as e:
            self.show_status(f"❌ Error: {e}", color="#d32f2f")

    def generate_pivot(self):
        selected = self.filter_var.get()
        if not selected:
            self.show_status("⚠️ Please select a C1_MARK value first.", color="#d32f2f")
            return
        
        self.show_status(f"\nℹ️ Generating pivot table...")

        app = None
        wb_xlw = None
        try:
            app = xw.App(visible=False)
            wb_xlw = app.books.open(self.out_file)
            sht = wb_xlw.sheets[self.base_name]

            # --- Find header row in Column G ---
            first_row = sht.range("G1").end("down").row
            header_value = sht.range(f"G{first_row}").value
            if str(header_value).strip().upper() != "C1_MARK":
                raise ValueError("First non-empty cell in Column G is not 'C1_MARK'")

            # --- Find ET header ---
            row_values = sht.range((first_row, 7), (first_row, sht.range((first_row, 7)).end("right").column)).value
            et_col = None
            for idx, val in enumerate(row_values, start=7):
                if str(val).strip().upper() == "ET":
                    et_col = idx
                    break
            if not et_col:
                raise ValueError("'ET' column not found to the right of C1_MARK")

            # --- Define pivot source range ---
            last_row = sht.range((first_row, 7)).end("down").row
            pivot_range = sht.range((first_row, 7), (last_row, et_col))

            # --- Create Pivot sheet ---
            pivot_sheet = wb_xlw.sheets.add("Pivot", after=sht)

            # --- Create pivot cache and table ---
            pivot_cache = wb_xlw.api.PivotCaches().Create(SourceType=1, SourceData=pivot_range.api)
            table_name = f"PivotTable_{datetime.now().strftime('%Y%m%d%H%M%S')}"
            pivot_table = pivot_cache.CreatePivotTable(TableDestination=pivot_sheet.range("A3").api, TableName=table_name)

            # --- Filter: C1_MARK ---
            pf = pivot_table.PivotFields("C1_MARK")
            pf.Orientation = 3
            valid_items = [item.Name for item in pf.PivotItems()]
            if selected in valid_items:
                pf.CurrentPage = selected
                self.show_status(f"\nApplied filter: {selected}")
            else:
                self.show_status(f"⚠️ Selected '{selected}' not found in C1_MARK items {valid_items}", color="#d32f2f")
                return

            # --- Rows: ET ---
            pivot_table.PivotFields("ET").Orientation = 1

            # --- Values: Count of FT ---
            pivot_table.AddDataField(pivot_table.PivotFields("FT"), "Count of FT", -4112)

            # --- Fallout Table Logic ---
            data = pivot_sheet.range("A4").expand().value
            sheet = wb_xlw.sheets[self.base_name]
            theoretical_num = None
            for i, val in enumerate(sheet.range("A:A").value, start=1):
                if str(val).strip().upper() == "THEORETICAL_NUM":
                    theoretical_num = sheet.range((i, 1)).offset(0, 2).value
                    break

            fallout_table = []
            for row in data:
                if not row or not row[0] or str(row[0]).strip().lower() == "grand total":
                    continue
                et_val = str(int(row[0])) if isinstance(row[0], (int, float)) and float(row[0]).is_integer() else str(row[0])
                count_val = int(row[1]) if isinstance(row[1], (int, float)) and float(row[1]).is_integer() else row[1]
                fallout = (float(row[1]) / theoretical_num * 100) if theoretical_num else 0
                fallout_table.append([et_val, count_val, f"{fallout:.2f}%"])

            fallout_table.sort(key=lambda x: int(x[1]), reverse=True)
            grand_total_val = str(int(theoretical_num)) if isinstance(theoretical_num, (int, float)) and float(theoretical_num).is_integer() else str(theoretical_num)
            fallout_table.insert(0, ["End Test No.", "Count", "Fallout%"])  # header row
            fallout_table.append(["Grand Total", grand_total_val, ""])

            # --- Vectorized write fallout table ---
            pivot_sheet.range("D3").value = fallout_table

            # --- Apply formatting ---
            last_row_ft = 3 + len(fallout_table) - 1
            fallout_range = pivot_sheet.range(f"D3:F{last_row_ft}")
            fallout_range.api.HorizontalAlignment = -4108
            fallout_range.api.VerticalAlignment = -4108
            fallout_range.api.IndentLevel = 0

            # Header row
            pivot_sheet.range("D3:F3").color = (192, 230, 245)
            pivot_sheet.range("D3:F3").api.Font.Bold = True
            # First data row
            pivot_sheet.range("D4:F4").color = (255, 159, 159)
            pivot_sheet.range("D4:F4").api.Font.Bold = True
            # Grand Total row
            pivot_sheet.range(f"D{last_row_ft}:F{last_row_ft}").color = (192, 230, 245)
            pivot_sheet.range(f"D{last_row_ft}:F{last_row_ft}").api.Font.Bold = True

            fallout_range.api.Borders.Weight = 2

            wb_xlw.save()

            # --- Show fallout table in status box ---
            self.status_box.config(state="normal")
            self.status_box.insert(tk.END, "\nPreview Table:\n")
            for et_val, count_val, fallout_val in fallout_table:
                self.status_box.insert(tk.END, f"{str(et_val):<15}{str(count_val):<10}{str(fallout_val)}\n")
            self.status_box.config(state="disabled")

            self.show_status(f"\n✅ Succesfully generated table for C1_MARK:{selected}")

        except Exception as e:
            self.show_status(f"❌ Error generating pivot/fallout: {e}", color="#d32f2f")

        finally:
            if wb_xlw:
                try: wb_xlw.close()
                except: pass
            if app:
                try: app.quit()
                except: pass


    def check_end_test(self):
        app = None
        wb_xlw = None
        try:
            app = xw.App(visible=False)
            wb_xlw = app.books.open(self.out_file)

            # --- Ensure Pivot sheet exists ---
            try:
                pivot_sheet = wb_xlw.sheets["Pivot"]
            except:
                pivot_sheet = wb_xlw.sheets.add("Pivot")

            data_sheet = wb_xlw.sheets[self.base_name]

            # --- Get highest fails End Test No from D4 ---
            raw_val = pivot_sheet.range("D4").value
            if raw_val is None:
                end_test_no = ""
            elif isinstance(raw_val, float) and raw_val.is_integer():
                end_test_no = str(int(raw_val))
            else:
                end_test_no = str(raw_val).strip()

            self.show_status(f"\n🔍Checking End Test No.: {end_test_no}")

            # --- Find LOLIMIT row in Column F ---
            lolimit_row = data_sheet.range("F1").end("down").row
            lolimit_val = data_sheet.range(f"F{lolimit_row}").value
            if str(lolimit_val).strip().upper() != "LOLIMIT":
                raise ValueError("LOLIMIT not found in Column F")

            # --- Expand reference table ---
            ref_range = data_sheet.range((lolimit_row, 1)).expand("table")

            # --- Locate TESTNO column (Column B) ---
            testno_values = data_sheet.range(
                (lolimit_row + 1, 2),
                (lolimit_row + ref_range.rows.count - 1, 2)
            ).value

            # Normalize TESTNO values to strings
            testno_values = ["" if v is None else str(int(v)) if isinstance(v, float) and v.is_integer() else str(v).strip() for v in testno_values]

            found_row = None
            if end_test_no in testno_values:
                idx = testno_values.index(end_test_no) + lolimit_row + 1
                found_row = idx

            # --- Vectorized write of header + data ---
            start_cell = pivot_sheet.range("H3")
            header = ["TSNO", "TESTNO", "COMMENT", "MODE", "HILIMIT", "LOLIMIT"]

            if found_row:
                row_values = data_sheet.range((found_row, 1), (found_row, 6)).value
                row_values = ["" if v is None else str(v).strip() for v in row_values]

                # Write header + data in one call
                pivot_sheet.range("H3").value = [header, row_values]

                # --- Apply formatting in bulk ---
                ref_range_excel = pivot_sheet.range("H3:M4")
                header_range = pivot_sheet.range("H3:M3")
                data_range = pivot_sheet.range("H4:M4")

                header_range.color = (192, 230, 245)   # light blue
                header_range.api.Font.Bold = True
                data_range.color = (255, 255, 255)     # white
                data_range.api.Font.Bold = True        # bold data row

                # Borders + alignment
                ref_range_excel.api.Borders.Weight = 2
                ref_range_excel.api.HorizontalAlignment = -4108
                ref_range_excel.api.VerticalAlignment = -4108
                ref_range_excel.api.IndentLevel = 0

                wb_xlw.save()

                # --- Show End Test No. table in status box ---
                self.status_box.config(state="normal")
                self.status_box.insert(tk.END, "\nEnd Test No. Reference:\n")
                self.status_box.insert(
                    tk.END,
                    f"{'TSNO':<10}{'TESTNO':<10}{'COMMENT':<15}{'MODE':<10}{'HILIMIT':<10}{'LOLIMIT'}\n"
                )
                self.status_box.insert(tk.END, "-" * 70 + "\n")
                tsno, testno, comment, mode, hilimit, lolimit = row_values
                self.status_box.insert(
                    tk.END,
                    f"{tsno:<10}{testno:<10}{comment:<15}{mode:<10}{hilimit:<10}{lolimit}\n"
                )
                self.status_box.config(state="disabled")

                # --- Status message depending on limits ---
                if lolimit != "":
                    self.show_status("\n✅ Found with Limits")
                else:
                    self.show_status("\n⚠️ Found with no Limit", color="#FFBF00")
            else:
                self.show_status("\n❌ No End Test No. found in the TESTNO Column", color="#d32f2f")

        except Exception as e:
            self.show_status(f"\n❌ Error checking End Test No: {e}", color="#d32f2f")

        finally:
            if wb_xlw:
                try: wb_xlw.close()
                except: pass
            if app:
                try: app.quit()
                except: pass


    def generate_wafermap(self):
        app = None
        wb_xlw = None
        try:
            app = xw.App(visible=False)
            wb_xlw = app.books.open(self.out_file)
            data_sheet = wb_xlw.sheets[self.base_name]

            # --- SLOT handling ---
            slot_row = None
            for i, val in enumerate(data_sheet.range("A:A").value, start=1):
                if str(val).strip().upper() == "SLOT":
                    slot_row = i
                    break

            if not slot_row:
                self.show_status("\n⚠️ SLOT header not found in Column A", color="#d32f2f")
                return

            slot_val = data_sheet.range((slot_row+1, 1)).value
            if slot_val is None:
                self.show_status("\n⚠️ SLOT value below header is empty", color="#d32f2f")
                return

            slot_str = str(int(slot_val)).zfill(2)
            self.show_status(f"\n🔍 Generating wafermap for W #{slot_str}...")
            sheet_name = f"W#{slot_str}_wafermap_by_End_Test_No"

            # --- Disable gridlines ---
            #data_sheet.api.Parent.Windows(1).DisplayGridlines = False
            wb_xlw.save()

            # --- Create or reuse Wafermap Pivot Table sheet ---
            try:
                pivot_sheet = wb_xlw.sheets["Wafermap Pivot Table"]
                pivot_sheet.clear()
            except:
                pivot_sheet = wb_xlw.sheets.add("Wafermap Pivot Table", after=data_sheet)

            # --- Create or reuse slot-specific wafermap sheet ---
            try:
                wafermap_sheet = wb_xlw.sheets[sheet_name]
                wafermap_sheet.clear()
            except:
                wafermap_sheet = wb_xlw.sheets.add(sheet_name, after=pivot_sheet)

            # --- Find header row in Column G ---
            first_row = data_sheet.range("G1").end("down").row

            # --- Read header row ---
            row_values = data_sheet.range(
                (first_row, 1),
                (first_row, data_sheet.range((first_row, 1)).end("right").column)
            ).value

            # --- Locate X, Y, ET columns ---
            x_col = y_col = et_col = None
            for idx, val in enumerate(row_values, start=1):
                if str(val).strip().upper() == "X":
                    x_col = idx
                elif str(val).strip().upper() == "Y":
                    y_col = idx
                elif str(val).strip().upper() in ["ET", "END TEST NO."]:
                    et_col = idx

            if not (x_col and y_col and et_col):
                raise ValueError("Required columns 'X', 'Y', 'ET' not found in header row")

            # --- Define pivot source range ---
            last_row = data_sheet.range((first_row+1, et_col)).end("down").row
            pivot_range = data_sheet.range((first_row, x_col), (last_row, et_col))

            # --- Create pivot cache and table ---
            pivot_cache = wb_xlw.api.PivotCaches().Create(SourceType=1, SourceData=pivot_range.api)
            table_name = f"PivotTable_{datetime.now().strftime('%Y%m%d%H%M%S')}"
            pivot_table = pivot_cache.CreatePivotTable(
                TableDestination=pivot_sheet.range("A1").api,
                TableName=table_name
            )

            # --- Configure pivot ---
            pivot_table.PivotFields("Y").Orientation = 1
            pivot_table.PivotFields("X").Orientation = 2
            pivot_table.AddDataField(pivot_table.PivotFields("ET"), "Min of ET", -4139)
            pivot_table.ColumnGrand = False
            pivot_table.RowGrand = False
            pivot_sheet.range("A2").value = "No."

            # --- Copy pivot output ---
            pivot_block = pivot_sheet.range("A2").expand()
            data_block = pivot_block.value

            # --- Paste values into wafermap sheet ---
            rows = len(data_block)
            cols = len(data_block[0])
            wafermap_sheet.range((1,1), (rows,cols)).value = data_block

            # --- Alignment ---
            wafermap_sheet.range((1,1), (rows,cols)).api.HorizontalAlignment = -4108
            wafermap_sheet.range((1,1), (rows,cols)).api.VerticalAlignment = -4108

            # --- Find last used row/col ---
            last_col = wafermap_sheet.range("1:1").end("right").column
            last_row = wafermap_sheet.range("A:A").end("down").row

            # --- Header formatting ---
            dark_blue = xw.utils.rgb_to_int((46, 110, 158))
            wafermap_sheet.range((1,1),(1,last_col)).color = (228, 241, 253)
            wafermap_sheet.range((1,1),(1,last_col)).api.Font.Color = dark_blue
            wafermap_sheet.range((1,1),(last_row,1)).color = (228, 241, 253)
            wafermap_sheet.range((1,1),(last_row,1)).api.Font.Color = dark_blue

            # --- Random colors for wafermap grid ---
            for r in range(2, last_row+1):
                for c in range(2, last_col+1):
                    cell = wafermap_sheet.range((r,c))
                    val = cell.value
                    if val is None or str(val).strip() == "":
                        continue
                    elif val == 0:
                        cell.color = (0,255,0)
                    else:
                        cell.color = (
                            random.randint(150,255),
                            random.randint(150,255),
                            random.randint(150,255)
                        )

            # --- Copy Row 1 (Ctrl+Shift+Right) and paste it after last used row ---
            row1_vals = wafermap_sheet.range((1,1),(1,last_col)).value
            wafermap_sheet.range((last_row+1,1),(last_row+1,last_col)).value = row1_vals
            wafermap_sheet.range((last_row+1,1),(last_row+1,last_col)).color = (228,241,253)
            wafermap_sheet.range((last_row+1,1),(last_row+1,last_col)).api.Font.Color = dark_blue
            wafermap_sheet.range((last_row+1,1),(last_row+1,last_col)).api.Font.Bold = True  # bold copy of Row 1

            # --- Copy Column A (Ctrl+Shift+Down) and paste it after last used column ---
            colA_vals = wafermap_sheet.range((1,1),(last_row,1)).value

            # Ensure values are shaped as a column (list of lists)
            if isinstance(colA_vals, list) and not isinstance(colA_vals[0], list):
                colA_vals = [[v] for v in colA_vals]

            # Paste Column A into the new rightmost column
            wafermap_sheet.range((1,last_col+1),(last_row,last_col+1)).value = colA_vals
            wafermap_sheet.range((1,last_col+1),(last_row,last_col+1)).color = (228,241,253)
            wafermap_sheet.range((1,last_col+1),(last_row,last_col+1)).api.Font.Color = dark_blue
            wafermap_sheet.range((1,last_col+1),(last_row,last_col+1)).api.Font.Bold = True  # bold copy of Column A

            # --- Add "No." at the very last row of that new column ---
            wafermap_sheet.range((last_row+1, last_col+1)).value = "No."
            wafermap_sheet.range((last_row+1, last_col+1)).color = (228,241,253)
            wafermap_sheet.range((last_row+1, last_col+1)).api.Font.Color = dark_blue
            wafermap_sheet.range((last_row+1, last_col+1)).api.Font.Bold = True  # bold "No." cell

            # --- Also bold the original Row 1 and Column A ---
            wafermap_sheet.range((1,1),(1,last_col)).api.Font.Bold = True
            wafermap_sheet.range((1,1),(last_row,1)).api.Font.Bold = True
            
            # --- Remove gridlines from wafermap sheet ---
            wafermap_sheet.api.Parent.Windows(1).DisplayGridlines = False

            # --- Alignment (center everything including mirrored row/col) ---
            used_range = wafermap_sheet.range((1,1),(last_row+1,last_col+1))
            used_range.api.HorizontalAlignment = -4108  # xlCenter
            used_range.api.VerticalAlignment = -4108    # xlCenter

            
            # --- Borders ---
            used_range = wafermap_sheet.range((1,1),(last_row+1,last_col+1))
            used_range.api.Borders.Weight = 2

            wb_xlw.save()
            wb_xlw.close()
            app.quit()

            self.show_status(f"\n✅ Wafermap created on {sheet_name} sheet.")

            # --- Reopen workbook to safely delete pivot sheet ---
            app = xw.App(visible=False)
            wb_xlw = app.books.open(self.out_file)

            try:
                pivot_sheet = wb_xlw.sheets["Wafermap Pivot Table"]
                # Activate another sheet first
                wb_xlw.sheets[0].activate()
                pivot_sheet.delete()
                #self.show_status("\n🗑️ Wafermap Pivot Table sheet deleted after reopen.")
            except Exception as e:
                #self.show_status(f"\n⚠️ Could not delete Wafermap Pivot Table: {e}", color="#d32f2f")
                pass
            
            wb_xlw.save()
            wb_xlw.close()
            app.quit()
            
        except Exception as e:
            self.show_status(f"\n❌ Error generating wafermap: {e}", color="#d32f2f")

        finally:
            if wb_xlw:
                try: wb_xlw.close()
                except: pass
            if app:
                try: app.quit()
                except: pass

                                
    def clear_all(self):
        # Reset file path
        self.path_var.set("")

        # Clear status box
        self.show_status("", clear=True)

        # Reset combobox selection and values
        if hasattr(self, "filter_dropdown"):
            self.filter_var.set("")                 # clear current selection
            self.filter_dropdown['values'] = []     # empty the dropdown list

# --- Run the App ---
if __name__ == "__main__":
    root = tk.Tk()
    app = AutomatingDeliverables(root)
    root.mainloop()
