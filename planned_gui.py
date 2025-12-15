
import tkinter as tk
from tkinter import filedialog, messagebox
import threading

try:
    import planned_parser
except Exception:
    planned_parser = None

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Planned Start + Patient Parser")
        self.input_path = tk.StringVar()
        self.status = tk.StringVar()
        self.count_var = tk.StringVar()
        self.months_var = tk.StringVar()
        self.total_rows_var = tk.StringVar()
        self.pre_rows_var = tk.StringVar()
        self.dedupe_var = tk.BooleanVar()
        self.dedupe_var.set(True)
        self.status.set("Select an Excel file (.xlsx)")
        self.count_var.set("")
        self.months_var.set("")
        self.total_rows_var.set("")
        self.pre_rows_var.set("")
        frm = tk.Frame(root, padx=12, pady=12)
        frm.pack(fill='both', expand=True)
        lbl = tk.Label(frm, text="Input file (.xlsx):")
        lbl.grid(row=0, column=0, sticky='w')
        ent = tk.Entry(frm, textvariable=self.input_path, width=50)
        ent.grid(row=1, column=0, sticky='we')
        btn_browse = tk.Button(frm, text="Browse", command=self.browse)
        btn_browse.grid(row=1, column=1, padx=6)
        chk = tk.Checkbutton(frm, text="Deduplicate by Planned Start (date-only) + Patient", variable=self.dedupe_var)
        chk.grid(row=2, column=0, sticky='w', pady=6)
        btn_run = tk.Button(frm, text="Run", command=self.run)
        btn_run.grid(row=3, column=0, pady=10, sticky='w')
        lbl_status = tk.Label(frm, textvariable=self.status, fg='blue')
        lbl_status.grid(row=4, column=0, columnspan=2, sticky='w')
        lbl_count = tk.Label(frm, textvariable=self.count_var, fg='green')
        lbl_count.grid(row=5, column=0, columnspan=2, sticky='w')
        lbl_months_title = tk.Label(frm, text="Month counts:")
        lbl_months_title.grid(row=6, column=0, sticky='w', pady=(8,0))
        lbl_months = tk.Label(frm, textvariable=self.months_var, justify='left')
        lbl_months.grid(row=7, column=0, columnspan=2, sticky='w')
        lbl_pre_rows = tk.Label(frm, textvariable=self.pre_rows_var, fg='brown')
        lbl_pre_rows.grid(row=8, column=0, columnspan=2, sticky='w', pady=(8,0))
        lbl_total_rows = tk.Label(frm, textvariable=self.total_rows_var, fg='purple')
        lbl_total_rows.grid(row=9, column=0, columnspan=2, sticky='w')
        frm.columnconfigure(0, weight=1)

    def browse(self):
        path = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.input_path.set(path)
            self.status.set("Ready to run")
            self.count_var.set("")
            self.months_var.set("")
            self.total_rows_var.set("")
            self.pre_rows_var.set("")

    def run(self):
        path = self.input_path.get().strip()
        if not path:
            messagebox.showwarning("Missing file", "Please select an input .xlsx file.")
            return
        do_dedupe = self.dedupe_var.get()
        def worker():
            try:
                self.status.set("Processing...")
                out, unique_count, month_counts, total_rows, pre_dedupe_rows = planned_parser.process_with_months(path, do_dedupe)
                self.status.set(f"Done. Output: {out}")
                if unique_count is not None:
                    self.count_var.set(f"Total unique pairs: {unique_count}")
                    self.pre_rows_var.set("Rows before dedupe: " + str(pre_dedupe_rows))
                else:
                    self.count_var.set("")
                    self.pre_rows_var.set("")
                lines = []
                for label, v in month_counts:
                    lines.append(label + ": " + str(v))
                self.months_var.set("\n".join(lines))
                self.total_rows_var.set("Total rows: " + str(total_rows))
            except Exception as e:
                self.status.set("Error")
                self.count_var.set("")
                self.months_var.set("")
                self.total_rows_var.set("")
                self.pre_rows_var.set("")
                messagebox.showerror("Error", str(e))
        t = threading.Thread(target=worker)
        t.start()

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
