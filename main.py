import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import os
import sys

# pandas is optional at runtime; show a friendly message if missing
try:
    import pandas as pd
except ImportError:
    # tkinter is already available; show an error dialog with install instructions
    try:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "Missing dependency",
            "The 'pandas' package is required but not installed.\n\nInstall it using:\n    pip install pandas openpyxl\n\nAfter installing, re-run this application.")
        root.destroy()
    except Exception:
        # fallback to stderr
        sys.stderr.write("Missing dependency: pandas. Install with: pip install pandas openpyxl\n")
    sys.exit(1)

# Color scheme (Continental: dark + orange accents)
DARK_BG = "#23272a"
DARK_FG = "#f4f6f8"
ORANGE = "#ff6600"
HEADER_BG = "#2c2f33"
SELECT_BG = "#ff6600"
SELECT_FG = "#23272a"


class PlannerViewer(tk.Toplevel):
    def __init__(self, master, df, plan_name="Planner Export", on_close=None, source_files=None):
        super().__init__(master)
        self.master = master
        self.title(plan_name)
        self.geometry("1000x600")
        self.configure(bg=DARK_BG)
        self.df = df.copy()
        self.filtered_df = self.df.copy()
        self.active_column = None
        # sorting state
        self.sort_col = None
        self.sort_asc = True
        # track original file paths used to create/merge this viewer (for session persistence)
        self.source_files = list(source_files) if source_files else []
        self.on_close_callback = on_close
        self.create_widgets()
        self.center_window()
        # Notify main app when closed so it can remove from its list
        self.protocol("WM_DELETE_WINDOW", self._handle_close)

    def _handle_close(self):
        if callable(self.on_close_callback):
            try:
                self.on_close_callback(self)
            except Exception:
                pass
        self.destroy()

    def create_widgets(self):
        # Search / toolbar frame
        search_frame = tk.Frame(self, bg=DARK_BG)
        search_frame.pack(fill=tk.X, padx=10, pady=8)

        tk.Label(search_frame, text="Filter:", font=("Segoe UI", 11), bg=DARK_BG, fg=ORANGE).pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        search_entry = tk.Entry(search_frame, textvariable=self.search_var, font=("Segoe UI", 11), width=40, bg=HEADER_BG, fg=DARK_FG, insertbackground=ORANGE)
        search_entry.pack(side=tk.LEFT, padx=8)
        search_entry.bind("<KeyRelease>", self.on_search)

        # control buttons: Add Task, Export
        btn_frame_right = tk.Frame(search_frame, bg=DARK_BG)
        btn_frame_right.pack(side=tk.RIGHT)

        add_task_btn = tk.Button(btn_frame_right, text="Add Task", font=("Segoe UI", 10, "bold"),
                     bg=ORANGE, fg=SELECT_FG, activebackground="#ff8a33",
                     command=self.add_task_dialog)
        add_task_btn.pack(side=tk.LEFT, padx=4)

        mark_done_btn = tk.Button(btn_frame_right, text="Mark Done", font=("Segoe UI", 10),
                     bg="#2e7d32", fg=SELECT_FG, activebackground="#4caf50",
                     command=self.mark_selected_done)
        mark_done_btn.pack(side=tk.LEFT, padx=4)

        delete_btn = tk.Button(btn_frame_right, text="Delete", font=("Segoe UI", 10),
                     bg="#8b0000", fg=DARK_FG, activebackground="#b22222",
                     command=self.delete_selected)
        delete_btn.pack(side=tk.LEFT, padx=4)

        export_btn = tk.Button(btn_frame_right, text="Export...", font=("Segoe UI", 10),
                       bg=HEADER_BG, fg=DARK_FG, command=self.export_to_excel)
        export_btn.pack(side=tk.LEFT, padx=4)

        # Treeview style
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Treeview",
                        font=("Segoe UI", 10),
                        rowheight=28,
                        background=DARK_BG,
                        fieldbackground=DARK_BG,
                        foreground=DARK_FG)
        style.configure("Treeview.Heading",
                        font=("Segoe UI", 11, "bold"),
                        background=HEADER_BG,
                        foreground=ORANGE)
        style.map("Treeview",
                  background=[("selected", SELECT_BG)],
                  foreground=[("selected", SELECT_FG)])

        cols = list(self.df.columns)
        self.tree = ttk.Treeview(self, columns=cols, show="headings", selectmode="browse")
        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 5))

        for col in cols:
            # on header click, filter by that column
            self.tree.heading(col, text=col, command=lambda c=col: self.on_column_click(c))
            self.tree.column(col, width=140, anchor=tk.W)

        # bind double-click on heading to sort
        self.tree.bind('<Double-1>', self._on_tree_double_click)

        # Column values combobox (shows unique values for active column)
        self.value_combo = ttk.Combobox(search_frame, values=[], width=30)
        self.value_combo.pack(side=tk.RIGHT, padx=(8, 0))
        self.value_combo.bind('<<ComboboxSelected>>', self._on_value_selected)

        clear_btn = tk.Button(search_frame, text="Clear", font=("Segoe UI", 9), bg=HEADER_BG, fg=DARK_FG, command=self._clear_filters)
        clear_btn.pack(side=tk.RIGHT, padx=4)

        self.populate_tree(self.df)

        # Status bar
        self.status_var = tk.StringVar()
        status_bar = tk.Label(self, textvariable=self.status_var, font=("Segoe UI", 10), bg=HEADER_BG, fg=ORANGE, anchor=tk.W)
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)
        self.update_status()

    def populate_tree(self, df):
        self.tree.delete(*self.tree.get_children())
        # apply sorting if set
        if self.sort_col and self.sort_col in df.columns:
            try:
                df_to_show = df.sort_values(by=self.sort_col, ascending=self.sort_asc, kind='mergesort')
            except Exception:
                df_to_show = df
        else:
            df_to_show = df

        for idx, row in df_to_show.iterrows():
            # convert row values to strings safely
            vals = ["" if pd.isna(v) else str(v) for v in row.tolist()]
            # use DataFrame index as tree item iid (item ID)
            self.tree.insert("", "end", iid=str(idx), values=vals)

    def on_search(self, event=None):
        query = self.search_var.get().lower().strip()
        if query == "":
            self.filtered_df = self.df.copy()
        elif self.active_column:
            mask = self.df[self.active_column].astype(str).str.lower().str.contains(query, na=False)
            self.filtered_df = self.df[mask]
        else:
            mask = self.df.apply(lambda row: row.astype(str).str.lower().str.contains(query, na=False).any(), axis=1)
            self.filtered_df = self.df[mask]
        self.populate_tree(self.filtered_df)
        self.update_status()
        self.highlight_active_column()
        # update values combobox for active column
        self._update_column_values_dropdown()

    def on_column_click(self, col):
        # Toggle active column filter: clicking the same column again clears filter
        if self.active_column == col:
            self.active_column = None
        else:
            self.active_column = col
        # Re-run search which will use active_column when present
        self.on_search()
        # when a column is activated, update values dropdown
        self._update_column_values_dropdown()

    def highlight_active_column(self):
        # Update header text slightly to show active column with an accent
        for c in self.df.columns:
            text = c
            if c == self.active_column:
                # show a simple [filter] marker (avoid null characters)
                text = f"{c} [filter]"
            # set heading text (can't change color easily per-heading cross-platform)
            self.tree.heading(c, text=text, anchor=tk.W)

    def _on_tree_double_click(self, event):
        # detect header double-click for sorting
        region = self.tree.identify_region(event.x, event.y)
        if region != 'heading':
            return
        col_id = self.tree.identify_column(event.x)  # like '#1'
        try:
            idx = int(col_id.replace('#', '')) - 1
            col = self.tree['columns'][idx]
        except Exception:
            return
        # toggle sorting: if same column, flip direction; else set asc
        if self.sort_col == col:
            self.sort_asc = not self.sort_asc
        else:
            self.sort_col = col
            self.sort_asc = True
        # re-populate
        self.populate_tree(self.filtered_df)
        self.update_status()

    def update_status(self):
        total = len(self.df)
        filtered = len(self.filtered_df)
        if self.active_column:
            self.status_var.set(f"Filtering by '{self.active_column}' | Showing {filtered} of {total} tasks")
        else:
            self.status_var.set(f"Showing {filtered} of {total} tasks")

    def center_window(self):
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        if w <= 1:  # not yet rendered; use default geometry
            w, h = 1000, 600
        ws = self.winfo_screenwidth()
        hs = self.winfo_screenheight()
        x = (ws // 2) - (w // 2)
        y = (hs // 2) - (h // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")

    # ---------- New methods for adding tasks ----------
    def add_task_dialog(self):
        """Open a dialog to manually add a new task (one input per column).

        Uses simple heuristics to pick input widgets:
        - date-like column names get a text entry with YYYY-MM-DD placeholder
        - status/priority get comboboxes with existing or common choices
        - columns with small sets of unique values get a combobox
        """
        dlg = tk.Toplevel(self)
        dlg.title("Add New Task")
        dlg.configure(bg=DARK_BG)
        # dialog height depends on number of columns
        dlg.geometry("520x{}+0+0".format(90 + 34 * max(3, len(self.df.columns))))
        dlg.transient(self)
        dlg.grab_set()

        frm = tk.Frame(dlg, bg=DARK_BG)
        frm.pack(fill=tk.BOTH, expand=True, padx=12, pady=10)

        entries = {}
        likely_status = ["status", "progress", "state"]
        likely_priority = ["priority"]
        likely_date = ["due", "date", "start"]

        for idx, col in enumerate(self.df.columns):
            lbl = tk.Label(frm, text=col + ":", font=("Segoe UI", 10), bg=DARK_BG, fg=DARK_FG)
            lbl.grid(row=idx, column=0, sticky="w", pady=6, padx=(0,6))

            col_lower = col.lower()
            widget = None
            # date-like
            if any(d in col_lower for d in likely_date):
                widget = tk.Entry(frm, font=("Segoe UI", 10), bg=HEADER_BG, fg=DARK_FG, insertbackground=ORANGE, width=48)
                widget.insert(0, "YYYY-MM-DD")
            elif any(p in col_lower for p in likely_priority) or any(s in col_lower for s in likely_status):
                existing = sorted(set(self.df[col].dropna().astype(str).unique().tolist()))
                choices = existing if existing else ["High", "Medium", "Low"]
                widget = ttk.Combobox(frm, values=choices, width=46)
            else:
                unique_vals = self.df[col].dropna().astype(str).unique()
                if 1 < len(unique_vals) <= 50:
                    widget = ttk.Combobox(frm, values=sorted(unique_vals), width=46)
                else:
                    widget = tk.Entry(frm, font=("Segoe UI", 10), bg=HEADER_BG, fg=DARK_FG, insertbackground=ORANGE, width=48)

            widget.grid(row=idx, column=1, pady=6, sticky="w")
            entries[col] = widget

        btn_frame = tk.Frame(dlg, bg=DARK_BG)
        btn_frame.pack(fill=tk.X, pady=(6,12))

        ok_btn = tk.Button(btn_frame, text="Add", font=("Segoe UI", 10, "bold"),
                           bg=ORANGE, fg=SELECT_FG, width=10,
                           command=lambda: self._on_add_task_confirm(entries, dlg))
        ok_btn.pack(side=tk.RIGHT, padx=8)
        cancel_btn = tk.Button(btn_frame, text="Cancel", font=("Segoe UI", 10),
                               bg=HEADER_BG, fg=DARK_FG, width=10, command=dlg.destroy)
        cancel_btn.pack(side=tk.RIGHT, padx=8)

        # center dialog over parent
        dlg.update_idletasks()
        px = self.winfo_rootx(); py = self.winfo_rooty()
        pw = self.winfo_width(); ph = self.winfo_height()
        dw = dlg.winfo_width(); dh = dlg.winfo_height()
        dlg.geometry(f"+{px + (pw-dw)//2}+{py + (ph-dh)//2}")

    def _on_add_task_confirm(self, entries, dlg):
        """Collect inputs, append a row to the DataFrame, refresh view and close dialog."""
        new_row = {}
        for col, ent in entries.items():
            new_row[col] = ent.get()
        # append preserving columns order
        new_df = pd.DataFrame([new_row], columns=self.df.columns)
        self.df = pd.concat([self.df, new_df], ignore_index=True)
        # normalize index to be simple integer range for stable iids
        self.df = self.df.reset_index(drop=True)

        # re-run current filter (keeps active_column behavior)
        self.on_search()
        # update column-values for the combobox to include new values
        self._update_column_values_dropdown()

        # focus / select the added row in the treeview
        children = self.tree.get_children()
        if children:
            last = children[-1]
            self.tree.see(last)
            self.tree.selection_set(last)

        dlg.destroy()
    # --------------------------------------------------

    # ----------------- helper methods -----------------
    def _update_column_values_dropdown(self):
        """Populate the values combobox with unique values from the active column (or clear it)."""
        try:
            if not self.active_column or self.active_column not in self.df.columns:
                self.value_combo['values'] = []
                self.value_combo.set("")
                return
            # use filtered_df to show relevant values
            vals = self.filtered_df[self.active_column].dropna().astype(str).unique().tolist()
            vals = sorted(vals)
            # keep at most 200 values to avoid performance issues
            if len(vals) > 200:
                vals = vals[:200]
            self.value_combo['values'] = vals
            self.value_combo.set("")
        except Exception:
            self.value_combo['values'] = []
            self.value_combo.set("")

    def _on_value_selected(self, event):
        val = self.value_combo.get()
        if val:
            self.search_var.set(val)
            self.on_search()

    def _clear_filters(self):
        self.search_var.set("")
        self.active_column = None
        self.sort_col = None
        self.sort_asc = True
        self.filtered_df = self.df.copy()
        self.populate_tree(self.filtered_df)
        self.update_status()
        self._update_column_values_dropdown()

    def export_to_excel(self):
        """Export current (filtered) view to an Excel file."""
        try:
            path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel files', '*.xlsx')], title='Export to Excel')
            if not path:
                return
            # export filtered view so user can save what they're seeing
            self.filtered_df.to_excel(path, index=False)
            messagebox.showinfo('Exported', f'Exported {len(self.filtered_df)} rows to:\n{path}')
        except Exception as e:
            messagebox.showerror('Error', f'Failed to export to Excel:\n{e}')
    # --------------------------------------------------

    # ---------- New methods for task actions ----------
    def mark_selected_done(self):
        """Mark the selected task(s) as done (set 'Status' to 'Completed')."""
        try:
            selected_items = self.tree.selection()
            if not selected_items:
                messagebox.showinfo("Info", "No tasks selected")
                return
            # update status to 'Completed' in the underlying DataFrame
            # find a status-like column
            status_col = None
            for col in self.df.columns:
                if col.lower() in ('status', 'state', 'progress'):
                    status_col = col
                    break
            if status_col is None:
                # create a Status column
                status_col = 'Status'
                self.df[status_col] = ''

            for item in selected_items:
                idx_label = item
                try:
                    idx = int(idx_label)
                except Exception:
                    continue
                if idx in self.df.index:
                    self.df.at[idx, status_col] = 'Completed'

            # re-run current filter and refresh
            self.on_search()
            self.update_status()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to mark task as done:\n{e}")

    def delete_selected(self):
        """Delete the selected task(s) from the DataFrame and update the view."""
        try:
            selected_items = self.tree.selection()
            if not selected_items:
                messagebox.showinfo("Info", "No tasks selected")
                return
            # confirm deletion
            if not messagebox.askyesno("Confirm Delete", "Are you sure you want to delete the selected task(s)?"):
                return
            # collect indices to delete (from tree iids which are DataFrame indices)
            indices_to_delete = []
            for item in selected_items:
                try:
                    idx = int(item)
                    indices_to_delete.append(idx)
                except Exception:
                    continue
            if not indices_to_delete:
                return
            # drop from underlying DataFrame
            self.df = self.df.drop(index=indices_to_delete)
            # reset index so we have simple integer indexes
            self.df = self.df.reset_index(drop=True)
            # re-run current filter and refresh
            self.on_search()
            self.update_status()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete task:\n{e}")
    # --------------------------------------------------


def load_excel(file_path):
    df = pd.read_excel(file_path)
    return df


class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Planner Multi-Plan Viewer")
        self.geometry("420x260")
        self.configure(bg=DARK_BG)
        self.viewers = []  # track open PlannerViewer instances
        self.create_widgets()
        self.center_window()
        # load previous session silently (if present)
        try:
            self.load_session()
        except Exception:
            pass
        # auto-save session on close
        self.protocol("WM_DELETE_WINDOW", self._on_exit)

    def create_widgets(self):
        tk.Label(self, text="Planner Multi-Plan Viewer", font=("Segoe UI", 16, "bold"), bg=DARK_BG, fg=ORANGE).pack(pady=14)
        add_btn = tk.Button(self, text="Add New Plan", font=("Segoe UI", 12), bg=ORANGE, fg=SELECT_FG, command=self.add_new_plan)
        add_btn.pack(pady=10)
        tk.Label(self, text="Each plan opens in a new window. You can also merge plans.", font=("Segoe UI", 10), bg=DARK_BG, fg=DARK_FG, wraplength=360, justify=tk.CENTER).pack(pady=6)

        # listbox showing opened plans
        self.lb = tk.Listbox(self, bg=HEADER_BG, fg=DARK_FG, width=50, height=4)
        self.lb.pack(pady=6)

        close_btn = tk.Button(self, text="Close Selected Viewer", font=("Segoe UI", 10), bg=HEADER_BG, fg=DARK_FG, command=self.close_selected_viewer)
        close_btn.pack(pady=(0,10))

        # Menu: File -> Save/Load Session, Exit
        menubar = tk.Menu(self)
        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label="Save Session...", command=self._save_session_dialog)
        filemenu.add_command(label="Load Session...", command=self._load_session_dialog)
        filemenu.add_separator()
        filemenu.add_command(label="Exit", command=self._on_exit)
        menubar.add_cascade(label="File", menu=filemenu)
        try:
            self.config(menu=menubar)
        except Exception:
            pass

    def add_new_plan(self):
        file_path = filedialog.askopenfilename(
            title="Select Microsoft Planner Exported Excel",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not file_path:
            return

        try:
            df = load_excel(file_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file:\n{e}")
            return

        # If there are existing viewers, ask whether to merge or open new
        if self.viewers:
            choice = self._ask_merge_or_new()
            if choice == "merge":
                # choose which viewer to merge into
                target = self._pick_viewer_dialog("Select Viewer to Merge Into")
                if target:
                    self._merge_into_viewer(target, df, file_path)
                return
            elif choice == "new":
                pass  # open new below
            else:
                return

        # open new PlannerViewer; track source file
        self._open_new_viewer(df, plan_name=file_path.split("\\")[-1], source_files=[file_path])

    def _open_new_viewer(self, df, plan_name="Planner Export", source_files=None):
        viewer = PlannerViewer(self, df, plan_name=plan_name, on_close=self._on_viewer_closed, source_files=source_files)
        self.viewers.append(viewer)
        self._refresh_listbox()

    # simple session persistence (save list of source files for each open viewer)
    def save_session(self, path=None):
        session = []
        for v in self.viewers:
            session.append({
                'title': v.title(),
                'sources': getattr(v, 'source_files', [])
            })
        path = path or os.path.join(os.path.expanduser('~'), '.planner_viewer_session.json')
        try:
            with open(path, 'w', encoding='utf8') as f:
                json.dump(session, f, indent=2)
        except Exception:
            pass

    def load_session(self, path=None):
        path = path or os.path.join(os.path.expanduser('~'), '.planner_viewer_session.json')
        if not os.path.exists(path):
            return
        try:
            with open(path, 'r', encoding='utf8') as f:
                session = json.load(f)
        except Exception:
            return
        for item in session:
            sources = item.get('sources') or []
            for s in sources:
                try:
                    df = load_excel(s)
                    self._open_new_viewer(df, plan_name=os.path.basename(s))
                    break
                except Exception:
                    continue

    def _on_viewer_closed(self, viewer):
        try:
            self.viewers.remove(viewer)
        except ValueError:
            pass
        self._refresh_listbox()

    def _refresh_listbox(self):
        self.lb.delete(0, tk.END)
        for v in self.viewers:
            title = v.title() if v.title() else "(untitled)"
            self.lb.insert(tk.END, title)

    def close_selected_viewer(self):
        sel = self.lb.curselection()
        if not sel:
            messagebox.showinfo("Info", "No viewer selected")
            return
        idx = sel[0]
        viewer = self.viewers[idx]
        viewer._handle_close()

    def _ask_merge_or_new(self):
        dlg = tk.Toplevel(self)
        dlg.title("Add Plan: Merge or New")
        dlg.configure(bg=DARK_BG)
        dlg.transient(self)
        dlg.grab_set()

        tk.Label(dlg, text="Import options", font=("Segoe UI", 12, "bold"), bg=DARK_BG, fg=ORANGE).pack(padx=12, pady=(12,6))

        choice_var = tk.StringVar(value="merge")
        rb1 = tk.Radiobutton(dlg, text="Merge into an existing window", variable=choice_var, value="merge", bg=DARK_BG, fg=DARK_FG, selectcolor=HEADER_BG)
        rb1.pack(anchor="w", padx=18, pady=4)
        rb2 = tk.Radiobutton(dlg, text="Open in a new window", variable=choice_var, value="new", bg=DARK_BG, fg=DARK_FG, selectcolor=HEADER_BG)
        rb2.pack(anchor="w", padx=18, pady=4)

        btn_frame = tk.Frame(dlg, bg=DARK_BG)
        btn_frame.pack(fill=tk.X, pady=12)
        ok_btn = tk.Button(btn_frame, text="OK", bg=ORANGE, fg=SELECT_FG, width=10, command=lambda: (dlg.destroy()))
        ok_btn.pack(side=tk.RIGHT, padx=8)
        cancel_btn = tk.Button(btn_frame, text="Cancel", bg=HEADER_BG, fg=DARK_FG, width=10, command=lambda: (choice_var.set("cancel"), dlg.destroy()))
        cancel_btn.pack(side=tk.RIGHT, padx=8)

        # center and wait
        dlg.update_idletasks()
        px = self.winfo_rootx(); py = self.winfo_rooty()
        pw = self.winfo_width(); ph = self.winfo_height()
        dw = dlg.winfo_width(); dh = dlg.winfo_height()
        dlg.geometry(f"+{px + (pw-dw)//2}+{py + (ph-dh)//2}")
        self.wait_window(dlg)
        return choice_var.get()

    def _pick_viewer_dialog(self, title="Select Viewer"):
        dlg = tk.Toplevel(self)
        dlg.title(title)
        dlg.configure(bg=DARK_BG)
        dlg.transient(self)
        dlg.grab_set()

        tk.Label(dlg, text=title, font=("Segoe UI", 12, "bold"), bg=DARK_BG, fg=ORANGE).pack(padx=12, pady=(12,6))

        lb = tk.Listbox(dlg, bg=HEADER_BG, fg=DARK_FG, width=60, height=6)
        lb.pack(padx=12, pady=6)
        for v in self.viewers:
            lb.insert(tk.END, v.title())

        result = {"viewer": None}

        def on_ok():
            sel = lb.curselection()
            if not sel:
                messagebox.showinfo("Info", "No viewer selected")
                return
            idx = sel[0]
            result["viewer"] = self.viewers[idx]
            dlg.destroy()

        btn_frame = tk.Frame(dlg, bg=DARK_BG)
        btn_frame.pack(fill=tk.X, pady=10)
        ok_btn = tk.Button(btn_frame, text="OK", bg=ORANGE, fg=SELECT_FG, width=10, command=on_ok)
        ok_btn.pack(side=tk.RIGHT, padx=8)
        cancel_btn = tk.Button(btn_frame, text="Cancel", bg=HEADER_BG, fg=DARK_FG, width=10, command=dlg.destroy)
        cancel_btn.pack(side=tk.RIGHT, padx=8)

        dlg.update_idletasks()
        px = self.winfo_rootx(); py = self.winfo_rooty()
        pw = self.winfo_width(); ph = self.winfo_height()
        dw = dlg.winfo_width(); dh = dlg.winfo_height()
        dlg.geometry(f"+{px + (pw-dw)//2}+{py + (ph-dh)//2}")
        self.wait_window(dlg)
        return result["viewer"]

    def _merge_into_viewer(self, viewer, new_df, file_path):
        # merge by union of columns, preserving order: existing columns first then new
        union_cols = list(dict.fromkeys(list(viewer.df.columns) + list(new_df.columns)))
        viewer.df = viewer.df.reindex(columns=union_cols, fill_value="")
        new_df = new_df.reindex(columns=union_cols, fill_value="")
        viewer.df = pd.concat([viewer.df, new_df], ignore_index=True)
        viewer.on_search()  # refresh with current filter state
        # record the source file for session persistence
        if not hasattr(viewer, 'source_files'):
            viewer.source_files = []
        try:
            if file_path not in viewer.source_files:
                viewer.source_files.append(file_path)
        except Exception:
            pass
        messagebox.showinfo("Merged", f"Imported and merged '{file_path.split('\\')[-1]}' into '{viewer.title()}'.")

    def center_window(self):
        """Center the main app window on screen."""
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        if w <= 1:
            w, h = 420, 260
        ws = self.winfo_screenwidth()
        hs = self.winfo_screenheight()
        x = (ws // 2) - (w // 2)
        y = (hs // 2) - (h // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _save_session_dialog(self):
        path = filedialog.asksaveasfilename(defaultextension='.json', filetypes=[('JSON', '*.json')], title='Save Session As')
        if not path:
            return
        try:
            self.save_session(path=path)
            messagebox.showinfo('Saved', f'Session saved to:\n{path}')
        except Exception as e:
            messagebox.showerror('Error', f'Failed to save session:\n{e}')

    def _load_session_dialog(self):
        path = filedialog.askopenfilename(title='Load Session', filetypes=[('JSON', '*.json')])
        if not path:
            return
        try:
            self.load_session(path=path)
            messagebox.showinfo('Loaded', f'Session loaded from:\n{path}')
        except Exception as e:
            messagebox.showerror('Error', f'Failed to load session:\n{e}')

    def _on_exit(self):
        # try to auto-save session
        try:
            self.save_session()
        except Exception:
            pass
        # close all child viewers
        for v in list(self.viewers):
            try:
                v._handle_close()
            except Exception:
                pass
        try:
            self.destroy()
        except Exception:
            os._exit(0)


# Add application entrypoint if missing
if __name__ == "__main__":
    app = MainApp()
    app.mainloop()
