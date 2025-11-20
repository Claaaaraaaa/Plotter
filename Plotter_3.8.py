# -*- coding: utf-8 -*-
"""
Created on Tue Jul 22 2025
Update version 3 on Wed Oct 01 2025


@author: Clara & Thomas
"""

import numpy as np
import matplotlib as mpl

# --- Keep only Ctrl+S in Matplotlib keymaps ---
mpl.rcParams["keymap.save"] = ["ctrl+s"]
for _k in (
    "keymap.grid", "keymap.fullscreen", "keymap.home", "keymap.back",
    "keymap.forward", "keymap.pan", "keymap.zoom", "keymap.quit",
    "keymap.yscale", "keymap.xscale", "keymap.grid_minor", "keymap.help"
):
    mpl.rcParams[_k] = []

import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os, pandas as pd, random, json
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure
from matplotlib import ticker as mticker


def apply_style(root):
    style = ttk.Style(root)
    style.theme_use('clam')
    style.configure('.', padding=4)
    # Optional: subtle gray for old paths
    style.configure("OldPath.TLabel", foreground="#666666", font=("Segoe UI", 9))

class RelinkDialog(tk.Toplevel):
    """Modal dialog to relink many missing files in one go."""
    def __init__(self, master, missing_rows, title="Relink missing files", filetype_by_kind=None, initialdir=None):
        super().__init__(master)
        self.title(title)
        self.resizable(True, True)
        self.result = None
        self.filetype_by_kind = filetype_by_kind or {}
        self._rows = []
        self._lastdir = initialdir or os.path.expanduser("~")

        # Make it modal
        self.transient(master)
        self.grab_set()

        # Main layout
        info = ttk.Label(self, text="Select a new location for each missing file, then press Apply.",
                         anchor="w", justify="left")
        info.pack(fill="x", padx=10, pady=(10, 6))

        container = ttk.Frame(self)
        container.pack(fill="both", expand=True, padx=10, pady=6)

        # Header
        hdr = ttk.Frame(container)
        hdr.grid(row=0, column=0, sticky="ew")
        ttk.Label(hdr, text="Type", width=6).grid(row=0, column=0, sticky="w")
        ttk.Label(hdr, text="Original path").grid(row=0, column=1, sticky="w", padx=(10,0))
        ttk.Label(hdr, text="New path").grid(row=0, column=2, sticky="w", padx=(10,0))
        container.columnconfigure(1, weight=1)
        container.columnconfigure(2, weight=1)

        # Rows
        for i, row in enumerate(missing_rows, start=1):
            kind = row["kind"]
            oldp = row["old_path"]

            ttk.Label(container, text=("DATA" if kind=="data" else "REF")).grid(row=i, column=0, sticky="w")
            ttk.Label(container, text=oldp, wraplength=520, style="OldPath.TLabel").grid(row=i, column=1, sticky="we", padx=(10,6))

            var = tk.StringVar(value="")
            ent = ttk.Entry(container, textvariable=var)
            ent.grid(row=i, column=2, sticky="we", padx=(10,6))

            def _browse(_kind=kind, _var=var, _old=oldp):
                """Ask user for a replacement path with proper filters."""
                types = self.filetype_by_kind.get(_kind, [("All files", "*.*")])
                start_dir = self._lastdir
                if not (start_dir and os.path.isdir(start_dir)):
                    parent = os.path.dirname(_old)
                    start_dir = parent if os.path.isdir(parent) else os.path.expanduser("~")

                newp = filedialog.askopenfilename(
                    title=f"Relink ({_kind}) — original: {os.path.basename(_old)}",
                    initialdir=start_dir,
                    filetypes=types
                )
                if newp:
                    _var.set(newp)
                    try:
                        self._lastdir = os.path.dirname(newp)
                    except Exception:
                        pass

            ttk.Button(container, text="Browse…", command=_browse).grid(row=i, column=3, sticky="w")

            self._rows.append((kind, oldp, var))

        # Buttons
        btns = ttk.Frame(self)
        btns.pack(fill="x", padx=10, pady=(6,10))
        ttk.Button(btns, text="Cancel", command=self._on_cancel).pack(side="right")
        ttk.Button(btns, text="Apply", command=self._on_apply).pack(side="right", padx=(0,6))

        # Keyboard bindings
        self.bind("<Escape>", lambda e: self._on_cancel())
        self.bind("<Return>", lambda e: self._on_apply())

        # Center on parent
        self.update_idletasks()
        try:
            x = master.winfo_rootx() + (master.winfo_width() - self.winfo_width()) // 2
            y = master.winfo_rooty() + (master.winfo_height() - self.winfo_height()) // 2
            self.geometry(f"+{x}+{y}")
        except Exception:
            pass

    def _on_cancel(self):
        self.result = None
        self.destroy()

    def _on_apply(self):
        """Collect mapping old->new (only rows with a new path)."""
        mapping = []
        for kind, oldp, var in self._rows:
            newp = var.get().strip()
            if newp:
                mapping.append({"kind": kind, "old_path": oldp, "new_path": newp})
        self.result = mapping
        self.destroy()

    def get_lastdir(self):
        return self._lastdir


class PrefixChangeDialog(tk.Toplevel):
    """
    Dialog to handle prefix-based path correction.
    - Displays the old common prefix (read-only)
    - Lets user choose a new base folder
    - Shows a live preview table of all affected files
    - Highlights found (green) vs missing (red)
    - Returns a mapping [{'kind','old_path','new_path'}]
    """
    def __init__(self, master, missing_rows, old_base_guess, initial_new_base=None, title="Prefix change (moved folder)"):
        super().__init__(master)
        self.title(title)
        self.resizable(True, True)
        self.result = None
        self._rows = missing_rows
        self._old_base = old_base_guess
        self._new_base_var = tk.StringVar(value=initial_new_base or "")

        # Modal behavior
        self.transient(master)
        self.grab_set()

        # --- Header: old/new base
        head = ttk.Frame(self)
        head.pack(fill="x", padx=10, pady=8)
        ttk.Label(head, text="Old prefix (detected):").grid(row=0, column=0, sticky="w")
        e_old = ttk.Entry(head)
        e_old.insert(0, self._old_base)
        e_old.configure(state="readonly")
        e_old.grid(row=0, column=1, sticky="we", padx=(6, 0))
        ttk.Label(head, text="New prefix:").grid(row=1, column=0, sticky="w", pady=(6, 0))
        e_new = ttk.Entry(head, textvariable=self._new_base_var)
        e_new.grid(row=1, column=1, sticky="we", padx=(6, 0), pady=(6, 0))

        def _browse_dir():
            d = filedialog.askdirectory(
                title="Select new prefix folder",
                initialdir=self._new_base_var.get() or os.path.expanduser("~")
            )
            if d:
                self._new_base_var.set(d)
                self._refresh_preview()

        ttk.Button(head, text="Browse…", command=_browse_dir).grid(row=1, column=2, padx=(6, 0), pady=(6, 0))
        head.columnconfigure(1, weight=1)

        # --- Preview table
        body = ttk.Frame(self)
        body.pack(fill="both", expand=True, padx=10, pady=(0, 8))
        cols = ("Type", "Relative to old", "New path", "Exists?")
        self.tv = ttk.Treeview(body, columns=cols, show="headings", height=12, selectmode="browse")
        for c, w in zip(cols, (60, 320, 420, 80)):
            self.tv.heading(c, text=c)
            self.tv.column(c, width=w, anchor="w", stretch=(c != "Type"))
        self.tv.pack(fill="both", expand=True)

        # Color tags
        self.tv.tag_configure("ok", foreground="#0b7a00")
        self.tv.tag_configure("ko", foreground="#b00020")

        # --- Buttons
        btns = ttk.Frame(self)
        btns.pack(fill="x", padx=10, pady=(0, 10))
        ttk.Button(btns, text="Cancel", command=self._cancel).pack(side="right")
        ttk.Button(btns, text="Apply", command=self._apply).pack(side="right", padx=(0, 6))

        # Live update when new prefix changes
        self._new_base_var.trace_add("write", lambda *_: self._refresh_preview())

        # Fill first time
        self._refresh_preview()

        # Center on parent
        self.update_idletasks()
        try:
            x = master.winfo_rootx() + (master.winfo_width() - self.winfo_width()) // 2
            y = master.winfo_rooty() + (master.winfo_height() - self.winfo_height()) // 2
            self.geometry(f"+{x}+{y}")
        except Exception:
            pass

    def _relpath_safe(self, path, start):
        """Compute relative path; fallback to filename on failure."""
        try:
            return os.path.relpath(path, start=start)
        except Exception:
            return os.path.basename(path)

    def _refresh_preview(self):
        """Rebuild the table according to the new prefix entered."""
        self.tv.delete(*self.tv.get_children())
        new_base = self._new_base_var.get().strip()
        for row in self._rows:
            kind = row["kind"]
            oldp = row["old_path"]
            rel = self._relpath_safe(oldp, self._old_base)
            newp = os.path.normpath(os.path.join(new_base, rel)) if new_base else ""
            exists = os.path.exists(newp) if newp else False
            tag = "ok" if exists else "ko"
            self.tv.insert("", "end",
                           values=(("DATA" if kind == "data" else "REF"), rel, newp, ("Yes" if exists else "No")),
                           tags=(tag,))

    def _apply(self):
        """Return mapping for existing files only."""
        new_base = self._new_base_var.get().strip()
        if not new_base:
            messagebox.showwarning("Missing prefix", "Please choose a new prefix.")
            return
        mapping = []
        for row in self._rows:
            kind = row["kind"]
            oldp = row["old_path"]
            rel = self._relpath_safe(oldp, self._old_base)
            newp = os.path.normpath(os.path.join(new_base, rel))
            if os.path.exists(newp):
                mapping.append({"kind": kind, "old_path": oldp, "new_path": newp})
        self.result = {"mapping": mapping, "new_base": new_base}
        self.destroy()

    def _cancel(self):
        """Abort without applying any mapping."""
        self.result = None
        self.destroy()

class ReplaceSegmentDialog(tk.Toplevel):
    """
    Replace an intermediate path segment across many missing files.
    - User enters 'Old segment' (text) and 'New segment' (text or chosen folder)
    - Live preview table shows the result and whether files exist
    - Returns mapping [{'kind','old_path','new_path'}]
    """
    def __init__(self, master, missing_rows, initial_dir=None, title="Replace path segment"):
        super().__init__(master)
        self.title(title); self.resizable(True, True)
        self._rows = missing_rows
        self.result = None
        self._initial_dir = initial_dir or os.path.expanduser("~")

        # --- Inputs
        head = ttk.Frame(self); head.pack(fill="x", padx=10, pady=8)
        ttk.Label(head, text="Old segment:").grid(row=0, column=0, sticky="w")
        self.old_seg = tk.StringVar()
        ttk.Entry(head, textvariable=self.old_seg).grid(row=0, column=1, sticky="we", padx=(6,0))

        ttk.Label(head, text="New segment:").grid(row=1, column=0, sticky="w", pady=(6,0))
        self.new_seg = tk.StringVar()
        ttk.Entry(head, textvariable=self.new_seg).grid(row=1, column=1, sticky="we", padx=(6,0), pady=(6,0))

        def _browse_dir():
            d = filedialog.askdirectory(title="Pick new segment folder", initialdir=self._initial_dir)
            if d:
                # Use normalized path separator to embed as a segment
                self.new_seg.set(d)
        ttk.Button(head, text="Browse…", command=_browse_dir).grid(row=1, column=2, padx=(6,0), pady=(6,0))
        head.columnconfigure(1, weight=1)

        ttk.Label(head, text="Tip: Match the exact folder part to replace (e.g., '\\\\Figures\\\\T°\\\\').",
                  foreground="#666").grid(row=2, column=1, columnspan=2, sticky="w", pady=(4,0))

        # --- Preview
        body = ttk.Frame(self); body.pack(fill="both", expand=True, padx=10, pady=(0,8))
        cols = ("Type", "Old path", "New path", "Exists?")
        self.tv = ttk.Treeview(body, columns=cols, show="headings", height=12, selectmode="browse")
        for c, w in zip(cols, (60, 360, 420, 80)):
            self.tv.heading(c, text=c)
            self.tv.column(c, width=w, anchor="w", stretch=(c != "Type"))
        ys = ttk.Scrollbar(body, orient="vertical", command=self.tv.yview)
        self.tv.configure(yscrollcommand=ys.set)
        self.tv.pack(side="left", fill="both", expand=True)
        ys.pack(side="right", fill="y")

        self.tv.tag_configure("ok", foreground="#0b7a00")
        self.tv.tag_configure("ko", foreground="#b00020")

        # --- Buttons
        btns = ttk.Frame(self); btns.pack(fill="x", padx=10, pady=(0,10))
        ttk.Button(btns, text="Cancel", command=self._cancel).pack(side="right")
        ttk.Button(btns, text="Apply", command=self._apply).pack(side="right", padx=(0,6))

        # live refresh
        self.old_seg.trace_add("write", lambda *_: self._refresh())
        self.new_seg.trace_add("write", lambda *_: self._refresh())
        self._refresh()

        # modal + centering
        self.transient(master); self.grab_set()
        self.update_idletasks()
        try:
            x = master.winfo_rootx() + (master.winfo_width() - self.winfo_width()) // 2
            y = master.winfo_rooty() + (master.winfo_height() - self.winfo_height()) // 2
            self.geometry(f"+{x}+{y}")
        except Exception:
            pass

    # --- Helpers
    def _refresh(self):
        self.tv.delete(*self.tv.get_children())
        old_raw = self.old_seg.get()
        new_raw = self.new_seg.get()
        found = 0; total = 0

        # Normalize separators and case (Windows-friendly)
        def norm(p): return os.path.normcase(os.path.normpath(p))
        old_norm = norm(old_raw) if old_raw else ""
        new_norm = new_raw  # keep user text; we build an abs or joined path below

        for r in self._rows:
            total += 1
            kind = "DATA" if r["kind"] == "data" else "REF"
            oldp = r["old_path"]
            newp = self._replace_once(oldp, old_norm, new_norm)
            exists = os.path.exists(newp) if newp else False
            if exists: found += 1
            tag = "ok" if exists else "ko"
            self.tv.insert("", "end", values=(kind, oldp, newp or "", "Yes" if exists else "No"), tags=(tag,))
        # optional: summary row or external label (kept simple here)

    def _replace_once(self, original_path, old_seg_norm, new_seg_text):
        if not old_seg_norm or not new_seg_text:
            return ""
        # Normalize both for matching; but keep original separators when rebuilding.
        original_norm = os.path.normcase(os.path.normpath(original_path))
        old_idx = original_norm.find(old_seg_norm)
        if old_idx < 0:
            return ""  # segment not found
        # Compute boundaries in the original string by mapping back via split/join for safety
        # Simpler: rebuild using os.sep-consistent strings
        # Split original into components, replace the first matching subpath sequence
        parts = original_norm.split(os.sep)
        old_parts = [p for p in old_seg_norm.split(os.sep) if p]
        if not old_parts:
            return ""

        # Sliding window search for first occurrence
        def find_window(parts, sub):
            for i in range(len(parts) - len(sub) + 1):
                if parts[i:i+len(sub)] == sub:
                    return i
            return -1

        idx = find_window(parts, old_parts)
        if idx < 0:
            return ""

        # Now rebuild using the *original* (non-lowercased) path to preserve accents etc.
        orig_parts = os.path.normpath(original_path).split(os.sep)
        left = orig_parts[:idx]
        right = orig_parts[idx+len(old_parts):]

        # If new_seg_text is an absolute path, we take its components; else treat it as components
        new_parts = os.path.normpath(new_seg_text).split(os.sep)
        # Avoid empty items (root split)
        new_parts = [p for p in new_parts if p]

        new_all = left + new_parts + right
        return os.path.normpath(os.sep.join(new_all))

    def _apply(self):
        mapping = []
        for iid in self.tv.get_children():
            tp, oldp, newp, ok = self.tv.item(iid, "values")
            if newp and os.path.exists(newp):
                mapping.append({
                    "kind": "data" if tp == "DATA" else "ref",
                    "old_path": oldp,
                    "new_path": newp
                })
        self.result = mapping
        self.destroy()

    def _cancel(self):
        self.result = None
        self.destroy()
        
class ReplaceAnyDialog(tk.Toplevel):
    """
    Let user select (by text selection) the exact substring to replace in an 'Old path' entry,
    then provide a replacement string (typed or picked via directory chooser).
    Live preview applies replacement once per missing path (case-insensitive on Windows).
    Returns mapping: [{'kind','old_path','new_path'}]
    """
    def __init__(self, master, missing_rows, last_dir=None, title="Replace prefix or segment"):
        super().__init__(master)
        self.title(title); self.resizable(True, True)
        self.result = None
        self._rows = missing_rows
        self._last_dir = last_dir or os.path.expanduser("~")

        # --- Header / sample old path (editable + selectable)
        head = ttk.LabelFrame(self, text="Selection of the part to replace")
        head.pack(fill="x", padx=10, pady=(10,6))
        ttk.Label(head, text="Old path (editable):").grid(row=0, column=0, sticky="w")
        self.sample_var = tk.StringVar(value=(self._rows[0]["old_path"] if self._rows else ""))
        self.ent_old = ttk.Entry(head, textvariable=self.sample_var)
        self.ent_old.grid(row=0, column=1, sticky="we", padx=(6,0))
        ttk.Button(head, text="Use from list ↓", command=self._load_from_selection).grid(row=0, column=2, padx=(6,0))
        head.columnconfigure(1, weight=1)

        # Tip
        ttk.Label(head, text="Select with mouse the exact part to replace.",
                  foreground="#666").grid(row=1, column=1, columnspan=2, sticky="w", pady=(4,2))

        # --- Replacement area
        repl = ttk.LabelFrame(self, text="Replacement")
        repl.pack(fill="x", padx=10, pady=6)
        ttk.Label(repl, text="Replace selection with:").grid(row=0, column=0, sticky="w")
        self.new_var = tk.StringVar(value="")
        ttk.Entry(repl, textvariable=self.new_var).grid(row=0, column=1, sticky="we", padx=(6,0))
        ttk.Button(repl, text="Browse…", command=self._browse_dir).grid(row=0, column=2, padx=(6,0))
        repl.columnconfigure(1, weight=1)

        # Options
        self.case_sensitive = tk.BooleanVar(value=(os.name != "nt"))  # Windows -> insensitive by default
        ttk.Checkbutton(repl, text="Case sensitive match", variable=self.case_sensitive).grid(row=1, column=1, sticky="w", pady=(4,0))

        # --- Preview table
        body = ttk.LabelFrame(self, text="Preview")
        body.pack(fill="both", expand=True, padx=10, pady=6)
        cols = ("Type", "Old path", "New path", "Exists?")
        self.tv = ttk.Treeview(body, columns=cols, show="headings", height=12, selectmode="browse")
        for c, w in zip(cols, (60, 360, 420, 80)):
            self.tv.heading(c, text=c)
            self.tv.column(c, width=w, anchor="w", stretch=(c != "Type"))
        ys = ttk.Scrollbar(body, orient="vertical", command=self.tv.yview)
        self.tv.configure(yscrollcommand=ys.set)  # <- plus de xscrollcommand
        self.tv.pack(side="left", fill="both", expand=True)
        ys.pack(side="right", fill="y")
        self.tv.tag_configure("ok", foreground="#0b7a00")
        self.tv.tag_configure("ko", foreground="#b00020")

        # --- List of missing paths (to pick a sample line)
        side = ttk.LabelFrame(self, text="Missing files")
        side.pack(fill="both", expand=False, padx=10, pady=(0,6))
        self.listbox = tk.Listbox(side, height=6, exportselection=False)
        for r in self._rows:
            self.listbox.insert(tk.END, r["old_path"])
        self.listbox.pack(fill="both", expand=True, padx=4, pady=4)

        # --- Buttons
        btns = ttk.Frame(self); btns.pack(fill="x", padx=10, pady=(0,10))
        ttk.Button(btns, text="Cancel", command=self._cancel).pack(side="right")
        ttk.Button(btns, text="Apply", command=self._apply).pack(side="right", padx=(0,6))

        # Live refresh
        self.sample_var.trace_add("write", lambda *_: self._refresh())
        self.new_var.trace_add("write",     lambda *_: self._refresh())
        self.case_sensitive.trace_add("write", lambda *_: self._refresh())
        # Also refresh on selection changes in Entry (needs a tiny delay)
        self.ent_old.bind("<<Selection>>", lambda e: self.after(10, self._refresh))
        self.ent_old.bind("<ButtonRelease-1>", lambda e: self.after(10, self._refresh))
        self.ent_old.bind("<KeyRelease>", lambda e: self.after(10, self._refresh))

        # Modal + center
        self.transient(master); self.grab_set()
        self.update_idletasks()
        try:
            x = master.winfo_rootx() + (master.winfo_width() - self.winfo_width()) // 2
            y = master.winfo_rooty() + (master.winfo_height() - self.winfo_height()) // 2
            self.geometry(f"+{x}+{y}")
        except Exception:
            pass

        self._refresh()

    # ---------- UI helpers ----------
    def _browse_dir(self):
        d = filedialog.askdirectory(title="Pick replacement folder", initialdir=self._last_dir)
        if d:
            self.new_var.set(d)
            try: self._last_dir = d
            except: pass

    def _load_from_selection(self):
        """Load the selected 'Old path' from Preview (Treeview) if any,
        otherwise from the Missing files Listbox, into the 'Old path' Entry."""
        # 1) Try the Treeview (Preview)
        try:
            sel = self.tv.selection()
            if sel:
                vals = self.tv.item(sel[0], "values")
                if vals and len(vals) >= 2:
                    oldp = vals[1]  # column order: ("Type", "Old path", "New path", "Exists?")
                    self.sample_var.set(oldp)
                    self.ent_old.focus_set()
                    self.ent_old.selection_range(0, tk.END)
                    return
        except Exception:
            pass
    
        # 2) Fallback: Listbox (Missing files)
        try:
            idxs = self.listbox.curselection()
            if idxs:
                oldp = self.listbox.get(idxs[0])
                self.sample_var.set(oldp)
                self.ent_old.focus_set()
                self.ent_old.selection_range(0, tk.END)
                return
        except Exception:
            pass

    # ---------- Core logic ----------
    def _get_selection(self):
        """Return currently selected substring in the Old path Entry (may be empty)."""
        try:
            i1 = self.ent_old.index("sel.first")
            i2 = self.ent_old.index("sel.last")
            txt = self.sample_var.get()
            return txt[int(i1):int(i2)]
        except tk.TclError:
            return ""

    def _replace_once(self, s, old_sub, new_sub, case_sensitive):
        """Replace first occurrence of old_sub in s (path-normalized match if insensitive)."""
        if not old_sub:
            return ""
        if case_sensitive:
            pos = s.find(old_sub)
            if pos < 0: return ""
            return s[:pos] + new_sub + s[pos+len(old_sub):]
        # Case-insensitive: find on lowered version but rebuild on original
        s_low = s.lower()
        old_low = old_sub.lower()
        pos = s_low.find(old_low)
        if pos < 0: return ""
        return s[:pos] + new_sub + s[pos+len(old_sub):]

    def _refresh(self):
        self.tv.delete(*self.tv.get_children())
        old_sel = self._get_selection()
        new_txt = self.new_var.get().strip()
        case_sensitive = self.case_sensitive.get()
        for r in self._rows:
            kind = "DATA" if r["kind"] == "data" else "REF"
            oldp = r["old_path"]
            newp = self._replace_once(oldp, old_sel, new_txt, case_sensitive) if old_sel and new_txt else ""
            ok = bool(newp and os.path.exists(newp))
            tag = "ok" if ok else "ko"
            self.tv.insert("", "end", values=(kind, oldp, newp, "Yes" if ok else "No"), tags=(tag,))

    def _apply(self):
        mapping = []
        for iid in self.tv.get_children():
            typ, oldp, newp, ok = self.tv.item(iid, "values")
            if newp and os.path.exists(newp):
                mapping.append({
                    "kind": "data" if typ == "DATA" else "ref",
                    "old_path": oldp,
                    "new_path": newp
                })
        self.result = mapping
        self.destroy()

    def _cancel(self):
        self.result = None
        self.destroy()



    
class Plotter:
    def __init__(self, master):
        self.master = master
        self.master.title("Plotter with Command Box")
        apply_style(self.master)
        self.master.geometry("980x640")
        self.files = []
        self.references = []
        self.commands = {}
        self.offset_between = 2.0
        self.default_color = 'black'
        self.custom_names = {}      # for data files
        self.custom_ref_names = {}  # for reference files
        self.build_gui()
        self._bind_shortcuts() 
        self._last_relink_dir = None   # remember last folder used for relinking
        self._error_buffer = []

    # ---- Centralized error accumulator ----
    def _add_error(self, kind: str, path: str, exc: Exception):
        """
        Collect a formatted error message without showing multiple popups.
        kind: 'DATA' or 'REF'
        path: file path that failed
        exc:  exception object
        """
        try:
            base = os.path.basename(path)
        except Exception:
            base = str(path)
        msg = f"[{kind}] {base} — {exc}"
        if not hasattr(self, "_error_buffer"):
            self._error_buffer = []
        self._error_buffer.append(msg)

    def _flush_errors(self, title="Import errors"):
        """
        Display a single error box containing all accumulated errors.
        Clears the buffer afterward.
        """
        if not getattr(self, "_error_buffer", []):
            return
        body = "Some files could not be loaded:\n\n" + "\n".join(self._error_buffer)
        messagebox.showerror(title, body)
        self._error_buffer.clear()

   # -------------------- Project save/load --------------------
    def save_project(self):
        filename = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")],
            title="Save Project As"
        )
        if not filename:
            return  # user cancelled

        project_data = {
            "data_files": self.files,
            "ref_files": self.references,
            "commands": self.cmd_entry.get("1.0", tk.END),
            "plot_settings": self.get_plot_settings_json(),
            "custom_names": self.custom_names,
            "custom_ref_names": self.custom_ref_names
        }

        try:
            with open(filename, "w") as f:
                json.dump(project_data, f, indent=2)
            #messagebox.showinfo("Project Saved", f"Project saved to:\n{filename}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save project:\n{e}")            
            
    def prune_custom_names(self):
        existing = set(self.files)
        self.custom_names = {k: v for k, v in self.custom_names.items() if k in existing}
        existing_ref = set(self.references)
        self.custom_ref_names = {k: v for k, v in self.custom_ref_names.items() if k in existing_ref}
            
     # -------------------- Plot settings helpers --------------------
    def get_plot_settings_json(self):
        """Return JSON-compatible dict of current plot settings."""
        # Replace these with your real axes, colors, log flags, etc.
        return {
            "xlim": [0, 10],
            "ylim": [0, 100],
            "colors": ["red", "blue"],
            "log_scale": False,
            "line_styles": self.parse_line_styles()   # <--- added for saving line styles

        }

    def apply_plot_settings_json(self, settings):
        """Apply plot settings from loaded project."""
        self.plot_settings = settings
        # Example: set axes limits/colors/log here if you use matplotlib

    def _shortcut_guard(self):
        """ True if global shortcuts are allowed (not typing in Entry/Text)."""
        w = self.master.focus_get()
        return not isinstance(w, (tk.Entry, tk.Text))
    
    def _bind_shortcuts(self):
        """ Bind Ctrl/Command shortcuts but ignore them while editing text."""
        def guard(fn):
            def _wrapped(event=None):
                if self._shortcut_guard():
                    fn()
                return "break"
            return _wrapped
    
        # Windows/Linux
        self.master.bind("<Control-Shift-d>", guard(self.load_files))
        self.master.bind("<Control-Shift-r>", guard(self.load_refs))
        self.master.bind("<Control-l>",       guard(self.load_project))
        self.master.bind("<Control-s>",       guard(self.save_project))
        self.master.bind("<Control-p>",       guard(self.apply_commands_and_plot))
    
        # macOS (Command key)
        self.master.bind("<Command-Shift-d>", guard(self.load_files))
        self.master.bind("<Command-Shift-r>", guard(self.load_refs))
        self.master.bind("<Command-l>",       guard(self.load_project))
        self.master.bind("<Command-s>",       guard(self.save_project))
        self.master.bind("<Command-p>",       guard(self.apply_commands_and_plot))


    # -------------------- GUI --------------------
    def build_gui(self):
        """Build the UI with a single 'Manage' tab (left: Data, right: References), plus Plot and Commands."""
        root = self.master
    
        # --- Notebook (tabs)
        nb = ttk.Notebook(root)
        nb.pack(fill='both', expand=True)
    
        tab_manage = ttk.Frame(nb)
        tab_plot   = ttk.Frame(nb)
    
        nb.add(tab_manage, text="Manage")
        nb.add(tab_plot,   text="Plot")
    
        # ===== Manage tab: single centered actions bar + split left/right =====
        
        # --- One centered actions bar (Load/Save Project) above both columns
        actions_bar = ttk.LabelFrame(tab_manage, text="Project Actions")
        actions_bar.pack(fill='x', padx=8, pady=(8, 4))
        
        # Use an inner frame packed with anchor='center' so buttons stay centered
        actions_row = ttk.Frame(actions_bar)
        actions_row.pack(pady=4)
        ttk.Button(actions_row, text="Load Project", command=self.load_project).pack(side='left', padx=6)
        ttk.Button(actions_row, text="Save Project", command=self.save_project).pack(side='left', padx=6)
        
        # --- Split area: left (Data) | right (References)
        pw = ttk.Panedwindow(tab_manage, orient='horizontal')
        pw.pack(fill='both', expand=True, padx=8, pady=(4, 8))
        
        left_pane  = ttk.Frame(pw)   # Data side
        right_pane = ttk.Frame(pw)   # References side
        pw.add(left_pane,  weight=1)
        pw.add(right_pane, weight=1)
        
        # ---- Left: Data files
        data_frame = ttk.LabelFrame(left_pane, text="Data files")
        data_frame.pack(fill='both', expand=True, padx=4, pady=4)
        
        row_data = ttk.Frame(data_frame)
        row_data.pack(fill='x', pady=4)
        ttk.Button(row_data, text="Load",       command=self.load_files).pack(side='left', padx=3)
        ttk.Button(row_data, text="Up",   command=lambda: self._tree_move_selected(self.data_list, self.files, -1)).pack(side='left', padx=3)
        ttk.Button(row_data, text="Down", command=lambda: self._tree_move_selected(self.data_list, self.files,  1)).pack(side='left', padx=3)
        ttk.Button(row_data, text="Reverse", command=lambda: self._tree_reverse(self.data_list, self.files)).pack(side='left', padx=3)
        ttk.Button(row_data, text="Remove All", command=self.remove_all_data).pack(side='left', padx=3)
        
        self.data_list = ttk.Treeview(data_frame, columns=("Name","Path"), show="headings", height=10, selectmode="extended")
        self.data_list.heading("Name", text="Name")
        self.data_list.heading("Path", text="Path")
        self.data_list.column("Name", width=320, anchor='w', stretch=True)
        self.data_list.column("Path", width=220, anchor='w', stretch=True)
        self.data_list.pack(fill='both', expand=True, padx=2, pady=2)
        self.data_list.bind("<Delete>",      self._delete_selected_data)
        self.data_list.bind("<Double-1>",    lambda e: self._start_inline_rename(self.data_list, self.custom_names, e))
        self.data_list.bind("<Control-r>",   lambda e: self._start_inline_rename(self.data_list, self.custom_names, e))
        self.data_list.bind("<Up>",          lambda e: self._move_selection(self.data_list, -1))
        self.data_list.bind("<Down>",        lambda e: self._move_selection(self.data_list,  1))
        self.data_list.bind("<Tab>",         lambda e: self._move_selection(self.data_list,  1))   # Tab = Down
        self.data_list.bind("<Shift-Tab>",   lambda e: self._move_selection(self.data_list, -1))   # Shift+Tab = Up
        self.data_list.bind("<Alt-Up>",   lambda e: self._tree_move_selected(self.data_list, self.files, -1))
        self.data_list.bind("<Alt-Down>", lambda e: self._tree_move_selected(self.data_list, self.files,  1))
        self.data_list.bind("<Alt-r>",  lambda e: self._tree_reverse(self.data_list, self.files))

        
        # ---- Right: Reference files
        ref_frame = ttk.LabelFrame(right_pane, text="Reference files")
        ref_frame.pack(fill='both', expand=True, padx=4, pady=4)
        
        row_ref = ttk.Frame(ref_frame)
        row_ref.pack(fill='x', pady=4)
        ttk.Button(row_ref, text="Load",       command=self.load_refs).pack(side='left', padx=3)
        ttk.Button(row_ref, text="Up",   command=lambda: self._tree_move_selected(self.ref_list, self.references, -1)).pack(side='left', padx=3)
        ttk.Button(row_ref, text="Down", command=lambda: self._tree_move_selected(self.ref_list, self.references,  1)).pack(side='left', padx=3)
        ttk.Button(row_ref, text="Reverse", command=lambda: self._tree_reverse(self.ref_list, self.references)).pack(side='left', padx=3)
        ttk.Button(row_ref, text="Remove All", command=self.remove_all_references).pack(side='left', padx=3)
        
        self.ref_list = ttk.Treeview(ref_frame, columns=("Name","Path"), show="headings", height=10, selectmode="extended")
        self.ref_list.heading("Name", text="Name")
        self.ref_list.heading("Path", text="Path")
        self.ref_list.column("Name", width=320, anchor='w', stretch=True)
        self.ref_list.column("Path", width=220, anchor='w', stretch=True)
        self.ref_list.pack(fill='both', expand=True, padx=2, pady=2)
        self.ref_list.bind("<Delete>",       self._delete_selected_refs)  # <-- fix: call the correct handler
        self.ref_list.bind("<Double-1>",     lambda e: self._start_inline_rename(self.ref_list, self.custom_ref_names, e))
        self.ref_list.bind("<Control-r>",    lambda e: self._start_inline_rename(self.ref_list, self.custom_ref_names, e))
        self.ref_list.bind("<Up>",           lambda e: self._move_selection(self.ref_list, -1))
        self.ref_list.bind("<Down>",         lambda e: self._move_selection(self.ref_list,  1))
        self.ref_list.bind("<Tab>",          lambda e: self._move_selection(self.ref_list,  1))
        self.ref_list.bind("<Shift-Tab>",    lambda e: self._move_selection(self.ref_list, -1))
        self.ref_list.bind("<Alt-Up>",   lambda e: self._tree_move_selected(self.ref_list, self.references, -1))
        self.ref_list.bind("<Alt-Down>", lambda e: self._tree_move_selected(self.ref_list, self.references,  1))
        self.ref_list.bind("<Alt-r>",   lambda e: self._tree_reverse(self.ref_list, self.references))
        
    
        # ===== Plot tab (simple, like Manage) =====
        # Split left (preview) | right (controls) with a basic Panedwindow
        plot_split = ttk.Panedwindow(tab_plot, orient='horizontal')
        plot_split.pack(fill='both', expand=True, padx=8, pady=8)
        
        left_plot  = ttk.Frame(plot_split)   # preview area
        right_ctrl = ttk.Frame(plot_split)   # controls (actions + commands)
        plot_split.add(left_plot,  weight=3)
        plot_split.add(right_ctrl, weight=1)
        
        # --- LEFT: Preview area (Matplotlib embedded) ---
        preview_frame = ttk.LabelFrame(left_plot, text="Preview")
        preview_frame.pack(fill='both', expand=True, padx=4, pady=4)
        
        # Keep a reference to the frame that hosts the Matplotlib canvas (must be set before binding)
        self.preview_frame = preview_frame
        
        # Auto-fit is ON by default: the figure follows the frame size
        self._cfg_binding = None
        self._autosize = True
        self._set_autosize(True)
        
        # Default rendering settings
        self.current_dpi = 100
        self.current_figsize = (6, 4)  # will be overridden by commands 'figsize'
        self._mpl_cids = []  # to track and (later) disconnect event callbacks 
        
        # Build a fresh Figure/Canvas/Toolbar using current settings
        self.fig = Figure(figsize=self.current_figsize, dpi=self.current_dpi)
        self.ax  = self.fig.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.preview_frame)
        self.canvas_widget = self.canvas.get_tk_widget()
        self.canvas_widget.pack(fill='both', expand=True)
        
        self.toolbar = NavigationToolbar2Tk(self.canvas, self.preview_frame, pack_toolbar=False)
        self.toolbar.update()
        self.toolbar.pack(side='bottom', fill='x')
        
        self.canvas_widget.configure(takefocus=0)
        self.canvas_widget.bind("<Key>", lambda e: "break")
        self._kill_mpl_keys()     
        
        # --- X cursor controls under Preview ---
        # Controls panel for enabling a vertical cursor and moving it along X with a slider
        cursor_panel = ttk.LabelFrame(left_plot, text="X cursor")
        cursor_panel.pack(fill='x', padx=4, pady=(0, 6))
        
        self._cursor_enabled = False            # cursor activation state
        self._cursor_x = None                   # current x position of the cursor
        self._cursor_vline = None               # matplotlib Line2D object for the vertical line
        self._cursor_cid_click = None           # mpl connection id for click callback
        
        # Enable/disable button
        self.cursor_btn = ttk.Button(cursor_panel, text="Enable cursor", command=self._toggle_cursor)
        self.cursor_btn.pack(side='left', padx=4, pady=4)
        
        # Slider for X; initial dummy range, updated after plotting
        self.cursor_scale = tk.Scale(
            cursor_panel, from_=0, to=1, resolution=0.01,
            orient='horizontal', length=300,
            showvalue=True,
            command=self._on_slider_change
        )
        self.cursor_scale.configure(state='disabled')
        self.cursor_scale.pack(side='left', padx=4, pady=2)
        
        # --- RIGHT: Controls (Actions on top, Commands below) ---
        actions_frame = ttk.LabelFrame(right_ctrl, text="Actions")
        actions_frame.pack(fill='x', padx=4, pady=(4,2))
        row = ttk.Frame(actions_frame); row.pack(fill='x', pady=2)
        ttk.Button(row, text="Apply & Plot",  command=self.apply_commands_and_plot).pack(side='left', padx=3)
        ttk.Button(row, text="Save Image",    command=self.save_plot).pack(side='left', padx=3)
        ttk.Button(row, text="Save Project",  command=self.save_project).pack(side='left', padx=3)
        
        cmd_frame = ttk.LabelFrame(right_ctrl, text="Commands")
        cmd_frame.pack(fill='both', expand=True, padx=4, pady=(2,4))
        
        # Create the text area ONLY here
        self.cmd_entry = tk.Text(cmd_frame, height=20, width=42, font=("Consolas", 10))
        self.cmd_entry.pack(fill='both', expand=True)
        self.cmd_entry.delete("1.0", tk.END)
        self.cmd_entry.insert("1.0", self.default_commands_text())
        
    def _kill_mpl_keys(self):
        """
        Disable keyboard shortcuts without breaking the toolbar.
        - Keep Matplotlib keymaps cleared elsewhere if you wish.
        - Swallow keypresses on the Tk canvas so plain letters do nothing.
        - Do NOT disconnect toolbar handlers or override toolbar.key_press.
        """
        try:
            self.canvas_widget.configure(takefocus=0)
            self.canvas_widget.bind("<Key>", lambda e: "break")
        except Exception:
            pass
        
    def default_commands_text(self):
        """Return the default text block for the Commands tab."""
        return (
            "--- Normalize & stacking ---\n"
            "normalize = off\n"
            "offset = 0\n"
            "refbase = -1\n"
            "refoffset = 0\n\n"
            
            "--- Title & labels ---\n"
            "title = \n"
            "xlabel = (°, Cu Kα)\n"
            "#xlim = 10,80\n"
            "ylabel = Intensity (a.u.)\n"
            "#ylim = 0,10\n"
            "#figsize_cm = 9,7\n"
            "axes_size_cm = 7.2,7.2\n"  
            "margins_cm = 0.6,0.1,0.1,1.2\n"
            "# Left, Right, Top, Bottom\n\n"

            "--- Legend & Colors ---\n"
            "legend = on\n"
            "#legendpos = outside\n"
            "colormap = rainbow\n"
            "legend_labelspacing = 0.1\n\n"
            
            "--- Font and linewidth Settings ---\n"
            "font = Times New Roman\n"
            "textcolor = black\n"
            "linewidth = 0.5\n"
            "reflinewidth = 2\n"
            "label_size = 10\n"
            "tick_size = 10\n"
            "title_size = 12\n"
            "legend_size = 9\n\n"
            "square_color = black\n"
            "data_bg = white\n"
            "legendlinewidth = 1\n"
            "legendlinewidthref = 1\n\n"
            
            "--- Ticks ---\n"
            "xtick_major = auto\n"
            "ytick_major = off\n"
            "xtick_minor = off\n"
            "ytick_minor = off\n"
        )

    def _axes_size_is_square(self, tol_cm: float = 1e-2) -> bool:
        """
        Return True if axes_size_cm is defined and width ≈ height (within tol_cm, in cm).
        """
        pair = self._parse_pair_cm("axes_size_cm")
        if not pair:
            return False
        w_cm, h_cm = pair
        return abs(w_cm - h_cm) <= tol_cm

    def _center_canvas_for_fixed_size(self):
        """Place the canvas widget at the center of the preview frame (fixed cm size)."""
        dpi = float(getattr(self, "current_dpi", 100))
        w_px = int(round(self.fig.get_figwidth()  * dpi))
        h_px = int(round(self.fig.get_figheight() * dpi))
    
        # Remove previous packing
        try: self.canvas_widget.pack_forget()
        except: pass
    
        self.canvas_widget.configure(width=w_px, height=h_px)
        self.canvas_widget.place(relx=0.5, rely=0.5, anchor="center")
        
    def _toggle_cursor(self):
        """Enable/disable the X cursor (vertical line) and click capture."""
        if not self._cursor_enabled:
            self._enable_cursor()
        else:
            self._disable_cursor()
    
    def _enable_cursor(self):
        # Enable UI and create/mount the vertical line on current axes
        self._cursor_enabled = True
        self.cursor_btn.config(text="Disable cursor")
        self.cursor_scale.configure(state='normal')
    
        # Determine initial x (mid of current xlim if None)
        try:
            xmin, xmax = self.ax.get_xlim()
        except Exception:
            xmin, xmax = (0.0, 1.0)
        if self._cursor_x is None:
            self._cursor_x = 0.5 * (xmin + xmax)
    
        # Create the vline on current axes
        if self._cursor_vline is None:
            self._cursor_vline = self.ax.axvline(self._cursor_x, color='black', linewidth=1)
    
        # Allow clicking inside axes to move the line & sync slider
        if self._cursor_cid_click is None:
            self._cursor_cid_click = self.canvas.mpl_connect('button_press_event', self._on_click_move)
    
        # Update slider to current axes bounds
        self._update_cursor_slider_from_axes()
        self.canvas.draw_idle()
    
    def _disable_cursor(self):
        # Disable UI and remove line and callbacks
        self._cursor_enabled = False
        self.cursor_btn.config(text="Enable cursor")
        self.cursor_scale.configure(state='disabled')
        if self._cursor_cid_click is not None:
            try:
                self.canvas.mpl_disconnect(self._cursor_cid_click)
            except Exception:
                pass
            self._cursor_cid_click = None
        if self._cursor_vline is not None:
            try:
                self._cursor_vline.remove()
            except Exception:
                pass
            self._cursor_vline = None
        self.canvas.draw_idle()
    
    def _on_slider_change(self, val):
        """Slider movement → update the line and redraw."""
        if not self._cursor_enabled:
            return
        try:
            x = float(val)
        except Exception:
            return
        self._cursor_x = x
        # Recreate vline if axes were cleared
        if self._cursor_vline is None:
            self._cursor_vline = self.ax.axvline(self._cursor_x, color='black', linewidth=1)
        else:
            self._cursor_vline.set_xdata([self._cursor_x, self._cursor_x])
        self.canvas.draw_idle()
    
    def _on_click_move(self, event):
        """Click inside the axes → place the line and synchronize the slider."""
        if (not self._cursor_enabled) or (event.inaxes != self.ax):
            return
        if event.xdata is None:
            return
        self._cursor_x = float(event.xdata)
        if self._cursor_vline is None:
            self._cursor_vline = self.ax.axvline(self._cursor_x, color='black', linewidth=1)
        else:
            self._cursor_vline.set_xdata([self._cursor_x, self._cursor_x])
        # Sync slider if x within range
        try:
            self.cursor_scale.set(self._cursor_x)
        except Exception:
            pass
        self.canvas.draw_idle()
    
    def _update_cursor_slider_from_axes(self):
        """Update the slider range from the current x-limits and center if needed."""
        try:
            xmin, xmax = self.ax.get_xlim()
        except Exception:
            xmin, xmax = (0.0, 1.0)
        span = max(1e-12, (xmax - xmin))  # avoid zero-span
        res = max(span / 1000.0, 1e-6)    # 1/1000 of span as resolution
    
        self.cursor_scale.configure(from_=xmin, to=xmax, resolution=res)
        if self._cursor_x is None or not (xmin <= self._cursor_x <= xmax):
            self._cursor_x = 0.5 * (xmin + xmax)
        self.cursor_scale.set(self._cursor_x)
    
    def _remount_cursor_after_clear(self):
        """Reattach the vline if the plot was cleared/recreated and resynchronize the slider."""
        if not self._cursor_enabled:
            return
        if self._cursor_vline is None:
            self._cursor_vline = self.ax.axvline(self._cursor_x if self._cursor_x is not None else 0, color='black', linewidth=1)
        self._update_cursor_slider_from_axes()

    def refresh_file_lists(self):
        """Refresh the Treeviews in the Data/References tabs, if present."""
        # Data
        if hasattr(self, 'data_list'):
            self.data_list.delete(*self.data_list.get_children())
            for f in self.files:
                base = os.path.basename(f)
                shown = self.custom_names.get(f, os.path.splitext(base)[0])
                self.data_list.insert("", "end", values=(shown, f))
        # References
        if hasattr(self, 'ref_list'):
            self.ref_list.delete(*self.ref_list.get_children())
            for f in self.references:
                base = os.path.basename(f)
                shown = self.custom_ref_names.get(f, os.path.splitext(base)[0])
                self.ref_list.insert("", "end", values=(shown, f))
                
    def _move_selection(self, tv: ttk.Treeview, delta: int):
        """Move selection/focus up or down by `delta` rows and keep item visible."""
        items = tv.get_children()
        if not items:
            return "break"
        current = tv.focus() or (tv.selection()[0] if tv.selection() else items[0])
        try:
            idx = items.index(current)
        except ValueError:
            idx = 0
        new_idx = max(0, min(len(items) - 1, idx + delta))
        new_id = items[new_idx]
        tv.selection_set(new_id)
        tv.focus(new_id)
        tv.see(new_id)
        return "break"
                
    def _start_inline_rename(self, tv: ttk.Treeview, name_dict: dict, event=None):
        """Start inline rename on the Name column (#1). Uses click position if provided."""
        # 1) Determine target item
        item_id = None
        if event is not None:
            # Use pointer position to get the exact row and ensure it's the Name column
            row = tv.identify_row(event.y)
            col = tv.identify_column(event.x)  # '#1' = Name
            region = tv.identify("region", event.x, event.y)
            if region == "cell" and col == "#1" and row:
                item_id = row
    
        if not item_id:
            # Fallback: use focus/selection
            item_id = tv.focus() or (tv.selection()[0] if tv.selection() else None)
        if not item_id:
            return "break"
    
        # 2) BBox of Name cell
        bbox = tv.bbox(item_id, "#1")
        if not bbox:
            tv.see(item_id)
            bbox = tv.bbox(item_id, "#1")
            if not bbox:
                return "break"
        x, y, w, h = bbox
    
        # 3) Current values (Name, Path)
        values = list(tv.item(item_id, "values"))
        if not values:
            return "break"
        old_name = values[0]
        path = values[1] if len(values) > 1 else None
    
        # 4) Create inline Entry overlay
        entry = tk.Entry(tv, borderwidth=1)
        entry.insert(0, old_name)
        entry.select_range(0, tk.END)
        entry.place(x=x, y=y, width=w, height=h)
        entry.focus_set()
        for seq in ("<Control-s>", "<Control-l>", "<Control-p>",
                    "<Control-Shift-d>", "<Control-Shift-r>"):
            entry.bind(seq, lambda e: "break")
    
        def _commit(move=None, reopen=False):
            new_name = entry.get().strip() or old_name
            values[0] = new_name
            tv.item(item_id, values=values)
        
            if path:
                name_dict[path] = new_name
        
            # Close current editor
            entry.destroy()
        
            # Keep selection on current item first
            tv.focus(item_id)
            tv.selection_set(item_id)
            tv.see(item_id)
        
            # Keyboard-driven navigation?
            if move is not None:
                # Move selection up/down
                self._move_selection(tv, move)
                # Reopen editor on the newly focused row if requested
                if reopen:
                    tv.after(10, lambda: self._start_inline_rename(tv, name_dict))
            else:
                # Stay on the same row
                if reopen:
                    tv.after(10, lambda: self._start_inline_rename(tv, name_dict))
            return "break"
        
        def _cancel(move=None):
            # Close editor and optionally move selection; do NOT reopen
            entry.destroy()
            tv.focus(item_id)
            tv.selection_set(item_id)
            tv.see(item_id)
            if move is not None:
                self._move_selection(tv, move)
            return "break"
    
        # 5) Key bindings inside editor
        entry.bind("<Return>",     lambda e: _commit(None, reopen=True))
        entry.bind("<KP_Enter>",   lambda e: _commit(None, reopen=True))
        entry.bind("<Escape>",     lambda e: _cancel(None))
        entry.bind("<FocusOut>",   lambda e: _commit(None, reopen=False))
        entry.bind("<Tab>",        lambda e: _commit(+1, reopen=True))
        entry.bind("<Shift-Tab>",  lambda e: _commit(-1, reopen=True))
        entry.bind("<Down>",       lambda e: _commit(+1, reopen=True))
        entry.bind("<Up>",         lambda e: _commit(-1, reopen=True))
    
        return "break"
    
    def _sync_backing_from_tv(self, tv: ttk.Treeview, backing_list: list):
        """Rebuild the Python list (files/references) from the Treeview's visual order."""
        new_order = []
        for iid in tv.get_children():
            vals = tv.item(iid, "values")
            if len(vals) >= 2:
                new_order.append(vals[1])  # Path column
        backing_list[:] = new_order
        
    def _tree_reverse(self, tv: ttk.Treeview, backing_list: list):
        """Reverse the Treeview visual order and synchronize the corresponding Python list.
        Keep the focus on a valid row after reversal."""
        items = list(tv.get_children())
        if not items:
            return "break"
    
        # Option: remember current index to restore focus
        cur = tv.focus()
        cur_idx = items.index(cur) if cur in items else 0
        new_idx = len(items) - 1 - cur_idx
    
        # Move items in reversed order
        for i, iid in enumerate(reversed(items)):
            tv.move(iid, "", i)
    
        # Reselect/refocus
        new_items = tv.get_children()
        target = new_items[new_idx] if new_items else ""
        if target:
            tv.selection_set(target)
            tv.focus(target)
            tv.see(target)
    
        # Sync the Python list
        self._sync_backing_from_tv(tv, backing_list)
        return "break"
    
    def _tree_move_selected(self, tv: ttk.Treeview, backing_list: list, delta: int):
        """Move the selected item by delta (+1/-1) rows in the Treeview
        and update the corresponding Python list."""
        sel = tv.selection()
        if not sel:
            return "break"
        iid = sel[0]
        idx = tv.index(iid)
        children = tv.get_children()
        new_idx = max(0, min(len(children) - 1, idx + delta))
        if new_idx != idx:
            tv.move(iid, "", new_idx)
            tv.selection_set(iid)
            tv.focus(iid)
            tv.see(iid)
            self._sync_backing_from_tv(tv, backing_list)
        return "break"
    
    def _delete_selected_data(self, event=None):
        """Delete selected rows from the Data tree and underlying self.files."""
        # Confirm deletion (optional)
        if not self.data_list.selection():
            return
        if not messagebox.askyesno("Delete data", "Remove selected data files from the list?"):
            return
    
        # Collect selected paths from the 'Path' column (index 1)
        selected_paths = []
        for iid in self.data_list.selection():
            vals = self.data_list.item(iid, 'values')
            if len(vals) >= 2:
                selected_paths.append(vals[1])
    
        # Filter out selected paths from self.files
        self.files = [p for p in self.files if p not in selected_paths]
        self.refresh_file_lists()
        self.data_list.focus_set() 
        
    def _delete_selected_refs(self, event=None):
        """Delete selected rows from the Reference tree and underlying self.references."""
        if not self.ref_list.selection():
            return
        if not messagebox.askyesno("Delete references", "Remove selected reference files from the list?"):
            return
    
        selected_paths = []
        for iid in self.ref_list.selection():
            vals = self.ref_list.item(iid, 'values')
            if len(vals) >= 2:
                selected_paths.append(vals[1])
    
        self.references = [p for p in self.references if p not in selected_paths]
        self.refresh_file_lists()
        self.ref_list.focus_set() 
            
    # -------------------- Placeholder methods --------------------
    def load_files(self):
        """Open a file dialog and append selected data files."""
        new_files = filedialog.askopenfilenames(filetypes=[("Files","*.xy *.csv *.dat *.txt *.gr")])
        if new_files:
            self.files.extend(new_files)
            self.refresh_file_lists()
    
    def remove_all_data(self):
        """Remove all loaded data files."""
        self.files.clear()
        messagebox.showinfo("Info", "All data files have been removed.")
        self.refresh_file_lists()
    
    def load_refs(self):
        """Open a file dialog and append selected reference files."""
        new_refs = filedialog.askopenfilenames(filetypes=[("Reference Files","*.csv *.xy *.txt *.xlsx")])
        if new_refs:
            self.references.extend(new_refs)
            self.refresh_file_lists()
    
    def remove_all_references(self):
        """Remove all loaded reference files."""
        self.references.clear()
        messagebox.showinfo("Info", "All reference files have been removed.")
        self.refresh_file_lists()
    
    def _remap_dict_key(self, d: dict, old_key: str, new_key: str):
        """If a key changed (path relink), move its value to the new key."""
        if old_key in d and old_key != new_key:
            d[new_key] = d.pop(old_key)
    
    def _guess_initialdir(self, project_json_path: str, paths: list[str]) -> str:
        """Heuristic to pick a sensible initial directory for the relink dialog."""
        # 1) try the JSON folder
        if project_json_path and os.path.isdir(os.path.dirname(project_json_path)):
            return os.path.dirname(project_json_path)
        # 2) try first existing sibling of paths
        for p in paths:
            if os.path.exists(p):
                return os.path.dirname(p)
        # 3) fallback: user's home
        return os.path.expanduser("~")
    
    def _relink_missing_entries(self, paths: list[str], name_dict: dict, dlg_title: str,
                                project_json_path: str, filetypes):
        """
        For each missing path, prompt the user to select a new location.
        Keeps custom names by remapping dict keys from old path to the new one.
        """
        unresolved = []
        initialdir = self._guess_initialdir(project_json_path, paths)
    
        for idx, old_path in enumerate(list(paths)):
            if os.path.exists(old_path):
                continue
            base = os.path.basename(old_path)
            # Ask user to relink this file
            if not messagebox.askyesno(
                "Relink missing file",
                f"{dlg_title}\n\nOriginal path:\n{old_path}\n\nDo you want to locate it now?"
            ):
                unresolved.append(old_path)
                continue
    
            new_path = filedialog.askopenfilename(
                title=f"{dlg_title} — select new location\n(Original: {base})",
                initialdir=initialdir,
                filetypes=filetypes
            )
            if new_path:
                # Replace path in the backing list
                try:
                    pos = paths.index(old_path)
                    paths[pos] = new_path
                except ValueError:
                    # Might have been already modified; append as fallback
                    paths.append(new_path)
                # Remap custom name key if present
                self._remap_dict_key(name_dict, old_path, new_path)
                # Update initialdir to the folder just used (quality of life)
                initialdir = os.path.dirname(new_path)
            else:
                unresolved.append(old_path)
    
        if unresolved:
            messagebox.showwarning(
                "Unresolved files",
                "Some files could not be relinked:\n\n" + "\n".join(unresolved)
            )
    
    def _apply_relink_mapping(self, mapping):
        """Apply a list of {'kind','old_path','new_path'} mappings to lists and custom dicts."""
        for m in mapping:
            kind = m["kind"]
            oldp = m["old_path"]
            newp = m["new_path"]
            if kind == "data":
                try:
                    i = self.files.index(oldp)
                    self.files[i] = newp
                except ValueError:
                    # if not found, append as fallback
                    self.files.append(newp)
                self._remap_dict_key(self.custom_names, oldp, newp)
            else:  # "ref"
                try:
                    i = self.references.index(oldp)
                    self.references[i] = newp
                except ValueError:
                    self.references.append(newp)
                self._remap_dict_key(self.custom_ref_names, oldp, newp)
                
    def _common_base(self, paths: list[str]) -> str:
        """Return the longest common existing/virtual base path among given paths."""
        if not paths:
            return ""
        try:
            return os.path.commonpath(paths)
        except Exception:
            # Fallback: incrementally reduce until valid
            parts = paths[0].split(os.sep)
            for i in range(len(parts), 0, -1):
                candidate = os.sep.join(parts[:i])
                try:
                    if all(p.startswith(candidate) for p in paths):
                        return candidate
                except Exception:
                    break
            return ""

    def _bulk_prefix_relink(self, missing_rows, project_json_path: str):
        """
        Open the prefix-change dialog (shows old/new prefix and live preview).
        Returns mapping list [{'kind','old_path','new_path'}] for found files.
        """
        all_missing = [r["old_path"] for r in missing_rows]
        old_base_guess = self._common_base(all_missing) or (
            os.path.dirname(project_json_path) if project_json_path else os.path.expanduser("~")
        )
    
        initial_new = self._last_relink_dir or os.path.expanduser("~")
    
        dlg = PrefixChangeDialog(
            self.master,
            missing_rows=missing_rows,
            old_base_guess=old_base_guess,
            initial_new_base=initial_new,
            title="Prefix change (moved folder)"
        )
        self.master.wait_window(dlg)
        if dlg.result:
            try:
                self._last_relink_dir = dlg.result.get("new_base") or self._last_relink_dir
            except Exception:
                pass
            return dlg.result.get("mapping", [])
        return []

    def _segment_replace_relink(self, missing_rows):
        """Open the 'replace path segment' dialog and return a mapping."""
        dlg = ReplaceSegmentDialog(self.master, missing_rows, initial_dir=self._last_relink_dir)
        self.master.wait_window(dlg)
        if dlg.result:
            try:
                # remember last dir as the folder of some new path
                self._last_relink_dir = os.path.dirname(dlg.result[-1]["new_path"])
            except Exception:
                pass
            return dlg.result
        return []
    
    def _ask_relink_mode(self):
        """
        Return 'prefix', 'segment', 'manual', or None if cancelled.
        """
        choice = {"val": None}
        win = tk.Toplevel(self.master)
        win.title("Relink strategy"); win.resizable(False, False)
        ttk.Label(win, text="How do you want to relink missing files?").pack(padx=12, pady=(12,8))
        frm = ttk.Frame(win); frm.pack(padx=12, pady=(0,12), fill="x")
    
        def set_and_close(val):
            choice["val"] = val; win.destroy()
    
        ttk.Button(frm, text="Prefix changed (moved base folder)", command=lambda: set_and_close("prefix")).pack(fill="x", pady=3)
        ttk.Button(frm, text="Replace a folder name in the middle", command=lambda: set_and_close("segment")).pack(fill="x", pady=3)
        ttk.Button(frm, text="Relink one by one (manual)", command=lambda: set_and_close("manual")).pack(fill="x", pady=3)
    
        # cancel handling
        ttk.Button(frm, text="Cancel", command=lambda: set_and_close(None)).pack(fill="x", pady=(10,0))
    
        win.transient(self.master); win.grab_set()
        self.master.wait_window(win)
        return choice["val"]

    def _replace_any_relink(self, missing_rows):
        """Open the unified selection-based relink dialog and return mapping."""
        dlg = ReplaceAnyDialog(self.master, missing_rows, last_dir=self._last_relink_dir)
        self.master.wait_window(dlg)
        if dlg.result:
            try:
                self._last_relink_dir = os.path.dirname(dlg.result[-1]["new_path"])
            except Exception:
                pass
            return dlg.result
        return []

    def load_project(self):
        """Load a project JSON and restore state (with batch relink UX)."""
        filename = filedialog.askopenfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")],
            title="Open Project"
        )
        if not filename:
            return
        try:
            with open(filename) as f:
                project_data = json.load(f)
    
            self.files = project_data.get("data_files", [])
            self.references = project_data.get("ref_files", [])
            self.cmd_entry.delete("1.0", tk.END)
            self.cmd_entry.insert(tk.END, project_data.get("commands", ""))
            self.apply_plot_settings_json(project_data.get("plot_settings", {}))
            self.custom_names = project_data.get("custom_names", {})
            self.custom_ref_names = project_data.get("custom_ref_names", {})
            self.prune_custom_names()
    
            # --- Detect missing files (both kinds) ---
            missing_rows = []
            for p in self.files:
                if not os.path.exists(p):
                    missing_rows.append({"kind": "data", "old_path": p})
            for p in self.references:
                if not os.path.exists(p):
                    missing_rows.append({"kind": "ref", "old_path": p})
        
            # --- If any missing, guide the user once ---
            if missing_rows:
                do_relink = messagebox.askyesno(
                    "Missing files detected",
                    "Some files referenced in this project are missing.\n\n"
                    "Do you want to relink them now?"
                )
                if do_relink:
                    # 1) Unified selection-based dialog (prefix or segment via selection)
                    mapping = self._replace_any_relink(missing_rows)
                    if mapping:
                        self._apply_relink_mapping(mapping)
            
                    # 2) Remaining unresolved? Offer manual dialog
                    remaining = []
                    for row in missing_rows:
                        newlist = self.files if row["kind"] == "data" else self.references
                        # If the old_path is still there and still missing on disk, it's unresolved
                        if row["old_path"] in newlist and (not os.path.exists(row["old_path"])):
                            remaining.append(row)
            
                    if remaining:
                        cont = messagebox.askyesno(
                            "Some files still missing",
                            "Some files could not be found with the replacement.\n"
                            "Do you want to select them manually now?"
                        )
                        if cont:
                            dlg = RelinkDialog(
                                self.master,
                                remaining,
                                title="Relink missing files",
                                filetype_by_kind={
                                    "data": [("Data files", "*.xy *.csv *.dat *.txt *.gr"), ("All files", "*.*")],
                                    "ref":  [("Reference files", "*.csv *.xy *.txt *.xlsx"), ("All files", "*.*")]
                                },
                                initialdir=self._last_relink_dir or self._guess_initialdir(filename, [r["old_path"] for r in remaining])
                            )
                            self.master.wait_window(dlg)
                            if dlg.result:
                                self._apply_relink_mapping(dlg.result)
                                self._last_relink_dir = dlg.get_lastdir()    
            # Refresh UI and plot
            self.refresh_file_lists()
            self.apply_commands_and_plot()
    
        except Exception as e:
            messagebox.showerror("Error", f"Could not load project:\n{e}")
    
    def save_plot(self):
        """Export the figure preserving exact physical sizes (in cm)."""
        if not (self.files or self.references):
            messagebox.showwarning("Warning", "No data or reference files loaded.")
            return
    
        file_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG Image", "*.png"),
                       ("PDF", "*.pdf"),
                       ("SVG", "*.svg"),
                       ("All Files", "*.*")],
            title="Save Figure"
        )
        if not file_path:
            return
    
        # Apply cm-based sizing for its side-effects; we don't need the return value here.
        self._apply_physical_size_from_cm(self.fig)
    
        # DPI for raster formats; ignored by vector (PDF/SVG).
        dpi = 300
        try:
            dpi = int(float(self.commands.get("export_dpi", dpi)))
        except Exception:
            pass
        dpi = max(72, min(1200, dpi))
    
        # IMPORTANT: keep bbox_inches=None to preserve margins set in centimeters.
        # Using 'tight' would alter margins and break your cm layout.
        try:
            self.fig.savefig(file_path, dpi=dpi, facecolor='white', bbox_inches=None)
            #messagebox.showinfo("Image Saved", f"Saved to:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save the figure:\n{e}")

    def apply_commands_and_plot(self):
        self.commands = self.parse_commands(self.cmd_entry.get("1.0", tk.END))
        self.plot_all()

    def parse_commands(self, text):
        cmd_dict = {}
        for line in text.strip().split("\n"):
            if '=' in line:
                key, value = line.split('=', 1)
                cmd_dict[key.strip().lower()] = value.strip()
        return cmd_dict

    def normalize(self, y):
        y_min = np.min(y)
        y_max = np.max(y)
        if y_max - y_min == 0:
            return y
        return (y - y_min) / (y_max - y_min)

    def robust_read_csv(self, filepath, max_header_lines=5):
        """
        Tries to read a reference file (csv, xy, xls) with unknown delimiter and variable header lines.
        - Tries common delimiters.
        - Tries to skip up to max_header_lines lines until data parses correctly.
        - Supports CSV-like and Excel files.
        Returns a DataFrame with at least two columns (angle, intensity).
        """
        ext = os.path.splitext(filepath)[1].lower()
        
        # For Excel files
        if ext in ['.xls', '.xlsx']:
            try:
                df = pd.read_excel(filepath)
                if df.shape[1] >= 2:
                    return df
            except Exception as e:
                raise ValueError(f"Cannot read Excel file: {e}")
        
        # For text files
        delimiters = [',', '\t', ';', ' ']
        
        # Try skipping 0 to max_header_lines lines
        for skip in range(max_header_lines + 1):
            for delim in delimiters:
                try:
                    df = pd.read_csv(filepath, delimiter=delim, skiprows=skip, engine='python', header=None)
                    # Check at least 2 numeric columns in data
                    if df.shape[1] >= 2:
                        # Check if first two columns are numeric (floats or ints)
                        try:
                            #df_check = df.iloc[:, :2].apply(pd.to_numeric)
                            return df.iloc[:, :2].copy()  # Return first 2 columns only
                        except:
                            continue
                except:
                    continue
        raise ValueError(f"Cannot parse reference file {filepath} with common delimiters and header skips.")
           
    def read_gr_file(self, filepath):
        """
        Custom reader for .gr files from PDFgetX3 which contain a config header.
        It skips lines until it finds the data block (starting with #L ...).
        """
        with open(filepath, 'r') as f:
            lines = f.readlines()
    
        # Find the line starting with '#L' which defines the data columns
        for idx, line in enumerate(lines):
            if line.strip().startswith("#L"):
                data_start = idx + 1
                break
        else:
            raise ValueError(f"Could not find '#L' header in {filepath}")
    
        # Now parse the data from that point onward
        from io import StringIO
        data_str = ''.join(lines[data_start:])
        df = pd.read_csv(StringIO(data_str), delim_whitespace=True, header=None)
        
        if df.shape[1] < 2:
            raise ValueError(f"Data section in {filepath} does not have two columns.")
        
        return df.iloc[:, :2].copy()

    def get_distinct_colors(self, n):
        colors = []
        for _ in range(n):
            r, g, b = [random.uniform(0.1, 0.9) for _ in range(3)]
            colors.append((r, g, b))
        return colors
    
    def parse_line_styles(self):
        """Parse line styles from the command box."""
        lines = self.cmd_entry.get("1.0", tk.END).splitlines()
        line_styles = {}
        for line in lines:
            line = line.strip()
            if line.startswith("line"):
                try:
                    key, value = line.split("=")
                    key = key.strip()   # line1
                    value = value.strip()  # solid/dashed/etc
                    line_styles[key] = value
                except:
                    pass
        return line_styles

    def prepare_options(self):
        def get_bool(key, default):
            return self.commands.get(key, str(default)).strip().lower() in ("on", "yes", "true")
    
        def get_opt_float(key):
            """Return float(value) if possible, else None (for 'auto' / empty / invalid)."""
            try:
                return float(str(self.commands.get(key)).strip())
            except Exception:
                return None
            
        def parse_tick_value(val):
            """Return ('off'|'auto'|float)."""
            if val is None:
                return "auto"
            v = str(val).strip().lower()
            if v in ("off", "none"):
                return "off"
            if v in ("on", "auto", ""):
                return "auto"
            try:
                return float(v)
            except:
                return "auto"
      
        def get_float(key, default):
            try:
                return float(self.commands.get(key, default))
            except Exception:
                return default
    
        opts = {
            "normalize": self.commands.get("normalize", "on").strip().lower(),
            "normalizeref": self.commands.get("normalizeref", "on").strip().lower(),
            "refbase":    get_float("refbase", -1.0),
            "refoffset":  get_float("refoffset", 1.0),
            "offset": get_float("offset", self.offset_between),
            "colormap": self.commands.get("colormap", None),
            "linewidth": get_float("linewidth", 1.5),
            "legendlinewidth": get_float("legendlinewidth", 2),
            "reflinewidth": get_float("reflinewidth", 2),
            "legendlinewidthref": get_float("legendlinewidthref", 2),
            "legend": get_bool("legend", True),
            "xlabel": self.commands.get("xlabel", ""),
            "ylabel": self.commands.get("ylabel", "Intensity (a.u.)"),
            "title": self.commands.get("title", ""),
            "legendpos": self.commands.get("legendpos", "best").strip().lower(),
            "default_size": get_float("default_size", 10),
            "label_size": get_float("label_size", 12),
            "title_size": get_float("title_size", 12),
            "tick_size": get_float("tick_size", 10),
            "legend_size": get_float("legend_size", 10),
            "xticks": get_bool("xticks", True),
            "yticks": get_bool("yticks", True),            
            "font": self.commands.get("font", "serif"),
            "legend_labelspacing": get_float("legend_labelspacing", 0.5),
            "square_color": self.commands.get("square_color", "black"),
            "textcolor": self.commands.get("textcolor", "black"),
            "data_bg": self.commands.get("data_bg", "white"),  # default white
            "xtick_major": parse_tick_value(self.commands.get("xtick_major", "auto")),
            "ytick_major": parse_tick_value(self.commands.get("ytick_major", "auto")),
            "xtick_minor": parse_tick_value(self.commands.get("xtick_minor", "off")),
            "ytick_minor": parse_tick_value(self.commands.get("ytick_minor", "off")),
            "square_width": get_opt_float("square_width") or 1.0,
        }
        return opts
        
    def _cm_to_in(self, cm: float) -> float:
        """Convert centimeters to inches."""
        return float(cm) / 2.54
    
    def _parse_pair_cm(self, key: str):
        """
        Parse 'w,h' in centimeters for a given command key.
        Accepts separators: ',', ';', ' ', 'x', '×'. Returns (w_cm, h_cm) or None.
        """
        raw = self.commands.get(key, "").strip()
        if not raw:
            return None
        txt = raw.replace("cm", "").replace("×", "x").strip()
        for sep in (",", ";", " ", "x"):
            if sep in txt:
                a, b = [p.strip() for p in txt.split(sep, 1)]
                break
        else:
            return None
        a = a.replace(",", "."); b = b.replace(",", ".")
        return (float(a), float(b))
    
    def _parse_margins_cm(self):
        """
        Parse margins_cm = left,right,top,bottom in cm.
        Defaults are publication-friendly: 1.5,1.0,1.0,1.2 cm.
        """
        raw = self.commands.get("margins_cm", "").strip()
        if not raw:
            return (1.5, 1.0, 1.0, 1.2)
        txt = raw.replace("cm", " ")
        for sep in (",", ";", " "):
            if sep in txt:
                parts = [p.strip().replace(",", ".") for p in txt.split(sep) if p.strip()]
                break
        else:
            parts = [txt.replace(",", ".")]
        if len(parts) != 4:
            return (1.5, 1.0, 1.0, 1.2)
        L, R, T, B = map(float, parts)
        return (L, R, T, B)
    
    def _apply_physical_size_from_cm(self, fig):
        """
        Apply physical sizing using centimeters.
        Priority:
          1) axes_size_cm (width,height) + margins_cm (L,R,T,B)  -> exact plotting area in cm
          2) figsize_cm (W,H) total figure size in cm
          3) figsize in inches (legacy)
        Returns True if a fixed size was applied (auto-fit must be OFF), else False.
        """
        axes_cm = self._parse_pair_cm("axes_size_cm")
        if axes_cm:
            L_cm, R_cm, T_cm, B_cm = self._parse_margins_cm()
            ax_w_cm, ax_h_cm = axes_cm
            fig_w_cm = ax_w_cm + L_cm + R_cm
            fig_h_cm = ax_h_cm + T_cm + B_cm
            fig.set_size_inches(self._cm_to_in(fig_w_cm), self._cm_to_in(fig_h_cm), forward=True)
            # Convert margins in cm to figure fractions for subplots_adjust:
            left   = L_cm / fig_w_cm
            right  = 1.0 - (R_cm / fig_w_cm)
            bottom = B_cm / fig_h_cm
            top    = 1.0 - (T_cm / fig_h_cm)
            fig.subplots_adjust(left=left, right=right, top=top, bottom=bottom)
            return True
    
        fig_cm = self._parse_pair_cm("figsize_cm")
        if fig_cm:
            fig.set_size_inches(self._cm_to_in(fig_cm[0]), self._cm_to_in(fig_cm[1]), forward=True)
            return True
    
        raw_fs = self.commands.get("figsize", "").strip().lower()
        if raw_fs and raw_fs != "auto":
            try:
                txt = raw_fs.replace(";", ",").replace(" ", ",")
                w_in, h_in = map(float, txt.split(",")[:2])
                fig.set_size_inches(w_in, h_in, forward=True)
                return True
            except:
                pass
    
        return False

    def _on_preview_resize(self, event):
        """Auto-fit: resize the figure to match the preview frame size."""
        if getattr(self, "_autosize", True):
            try:
                w_in = max(1, event.width) / float(self.current_dpi)
                h_in = max(1, event.height) / float(self.current_dpi)
                self.current_figsize = (w_in, h_in)
                self.fig.set_size_inches(w_in, h_in, forward=True)
                self.canvas.draw_idle()
            except Exception:
                pass
    
    def _set_autosize(self, enable: bool):
        """Enable/disable auto-fit (<Configure> binding on the preview frame)."""
        self._autosize = bool(enable)
        if enable:
            if not getattr(self, "_cfg_binding", None):
                self._cfg_binding = self.preview_frame.bind("<Configure>", self._on_preview_resize)
        else:
            if getattr(self, "_cfg_binding", None):
                try:
                    self.preview_frame.unbind("<Configure>", self._cfg_binding)
                except Exception:
                    pass
                self._cfg_binding = None

    def plot_all(self):
        if not self.files and not self.references:
            messagebox.showwarning("Warning", "No data or reference files loaded.")
            return
    
        options = self.prepare_options()
        # Reset error buffer for this plotting session
        self._error_buffer = []        
    
        # Update global style
        plt.rcParams.update({
            "xtick.major.width": 1,
            "ytick.major.width": 1,
            "pdf.fonttype": 42,
            "ps.fonttype": 42,
            "font.size": options["default_size"],
            "axes.labelsize": options["label_size"],
            "axes.titlesize": options["title_size"],
            "xtick.labelsize": options["tick_size"],
            "ytick.labelsize": options["tick_size"],
            "legend.fontsize": options["legend_size"],
            "font.family": options["font"],
            "axes.labelcolor": options["textcolor"],
            "axes.titlecolor": options["textcolor"],
            "axes.edgecolor": options["square_color"],
            "xtick.color": options["square_color"],
            "ytick.color": options["square_color"],
            "xtick.labelcolor": options["textcolor"],
            "ytick.labelcolor": options["textcolor"],
            "text.color": options["textcolor"],
        })
    
        # ---- Use the embedded Figure/Axes
        ax = self.ax
        fig = self.fig
        
        # Compute once: apply cm-based sizing; then toggle auto-fit and center if needed.
        fixed_applied = self._apply_physical_size_from_cm(fig)
        self._set_autosize(not fixed_applied)
        if fixed_applied:
            self._center_canvas_for_fixed_size()
        
        ax.clear()

        # --- Apply axes border ("square") color to spines ---
        # Use direct spine styling because rcParams won't retroactively recolor existing axes.
        for sp in ax.spines.values():
            sp.set_edgecolor(options["square_color"])   # set border color
              
        # Enforce square plotting area only when axes_size_cm is square.
        if self._axes_size_is_square():
            ax.set_box_aspect(1)           # force exact square plotting area
        else:
            try:
                ax.set_box_aspect(None)    # release any previous square lock
            except Exception:
                pass
        
        # Make sure previous callbacks are disconnected, avoid stacking.
        try:
            for cid in getattr(self, "_mpl_cids", []):
                self.canvas.mpl_disconnect(cid)
        except Exception:
            pass
        self._mpl_cids = []
    
        # Axes background
        if options["data_bg"].lower() in ("transparent", "none"):
            ax.set_facecolor("none")
        else:
            ax.set_facecolor(options["data_bg"])
        fig.patch.set_alpha(0)
    
        offset = options["offset"]
        n = len(self.files)
    
        pattern_handles, pattern_labels = [], []
        ref_handles, ref_labels = [], []
    
        if not options["yticks"]:
            ax.set_yticks([])
        if not options["xticks"]:
            ax.set_xticks([])
    
        # Colors for data curves
        if options["colormap"]:
            try:
                cmap = plt.get_cmap(options["colormap"])
                col_colors = [cmap(i / max(1, len(self.files) - 1)) for i in range(len(self.files))]
            except Exception:
                col_colors = [self.default_color] * len(self.files)
        else:
            col_colors = [self.default_color] * len(self.files)
    
        line_styles = self.parse_line_styles()
    
        # === Plot DATA files ===
        for i, file_path in enumerate(self.files):
            try:
                ext = os.path.splitext(file_path)[1].lower()
                if ext == '.csv':
                    df = self.robust_read_csv(file_path)
                    r = df.iloc[:, 0].values
                    intensity = df.iloc[:, 1].values
                else:
                    if ext == '.gr':
                        df = self.read_gr_file(file_path)
                        r = df.iloc[:, 0].values
                        intensity = df.iloc[:, 1].values
                    else:
                        data = np.loadtxt(file_path, comments="#", skiprows=1)
                        r, intensity = data[:, 0], data[:, 1]
    
                intensity_norm = intensity if options["normalize"] == "off" else self.normalize(intensity)
                shifted = intensity_norm + (offset * (n - i - 1))
    
                base_name = os.path.splitext(os.path.basename(file_path))[0]
                custom_label = self.custom_names.get(file_path, base_name)
                custom_label = self.commands.get(f"name{i+1}", custom_label)
    
                color = self.commands.get(f"color{i+1}", col_colors[i])
                linewidth = options["linewidth"]
                legend_lw = options["legendlinewidth"]
                linestyle = line_styles.get(f"line{i+1}", "solid")
    
                # plot on embedded axes
                ax.plot(r, shifted, label=custom_label, color=color,
                        linewidth=linewidth, linestyle=linestyle)
    
                # invisible legend line (pattern) for consistent legend thickness
                legend_line, = ax.plot([], [], color=color, label=custom_label,
                                       linewidth=legend_lw, linestyle=linestyle)
                pattern_handles.append(legend_line)
                pattern_labels.append(custom_label)
    
            except Exception as e:
                # Instead of showing one popup per file, collect errors
                self._add_error("DATA", file_path, e)
                continue
    
        # === Plot REFERENCE files ===
        ref_color_list = self.get_distinct_colors(len(self.references))
        for idx, ref_path in enumerate(self.references):
            try:
                ext = os.path.splitext(ref_path)[1].lower()
                color = self.commands.get(f"refcolor{idx+1}", ref_color_list[idx])
                is_peak_list = False
        
                # -- lecture x, y --
                if ext in [".csv", ".xy", ".txt", ".dat"]:
                    df = pd.read_csv(ref_path, sep=r"[,\t; ]+", engine="python", header=None, comment="#")
                elif ext == ".xlsx":
                    df = pd.read_excel(ref_path)
                else:
                    raw = np.loadtxt(ref_path)
                    if raw.ndim == 1:
                        x = raw
                        y = np.ones_like(raw)
                        is_peak_list = True
                    else:
                        x = raw[:, 0]; y = raw[:, 1]
                        is_peak_list = False
                    df = None
        
                if df is not None:
                    if '2Theta (°)' in df.columns:
                        if 'I var' in df.columns:
                            intensity_col = 'I var'
                        elif 'I fix' in df.columns:
                            intensity_col = 'I fix'
                        else:
                            intensity_col = df.columns[1]
                
                        x = df['2Theta (°)'].astype(str).str.replace(',', '.').astype(float).values
                        y = df[intensity_col].astype(str).str.replace(',', '.').astype(float).values
                        is_peak_list = True
                    else:
                        x = df.iloc[:, 0].astype(str).str.replace(',', '.').astype(float).values
                        y = df.iloc[:, 1].astype(str).str.replace(',', '.').astype(float).values
                        is_peak_list = True
        
                y_norm   = y if options["normalizeref"] == "off" else self.normalize(y)
                base_ref = options["refbase"]
                step_ref = options["refoffset"]
                legacy = (self.commands.get("stackrefs", "") or "").strip().lower()
                if legacy in ("off", "no", "false"):
                    step_ref = 0.0
                base_y = base_ref - idx * step_ref
                span_factor = float(self.commands.get("refspan", 0.95))
                direction = 1.0 if step_ref >= 0 else -1.0
                span = (abs(step_ref) * span_factor) if step_ref != 0 else 1.0
                jitter_cmd = (self.commands.get("refxjitter", "auto") or "auto").strip().lower()
                if jitter_cmd == "auto":
                    try:
                        xmin, xmax = ax.get_xlim()
                        xjitter = 0.003 * (xmax - xmin)
                    except Exception:
                        xjitter = 0.5
                else:
                    try:
                        xjitter = float(jitter_cmd.replace(",", "."))
                    except Exception:
                        xjitter = 0.0
        
                if is_peak_list:
                    seen = set()
                    y_max = float(y.max()) if len(y) else 1.0
                    for px, py in zip(x, y):
                        py_norm = (py / y_max)
                        if py_norm > 0:
                            k = round(float(px), 3)
                            if k not in seen:
                                x_shift = (idx - (len(self.references) - 1) / 2.0) * xjitter
                                height = py_norm * span
                                ax.vlines(px + x_shift, base_y, base_y + direction * height,
                                          color=color, linewidth=options["reflinewidth"])
                                seen.add(k)
                else:
                    try:
                        from scipy.signal import find_peaks
                        peaks, _ = find_peaks(y_norm)
                    except Exception:
                        peaks = [i for i in range(1, len(y_norm)-1) if y_norm[i] > y_norm[i-1] and y_norm[i] > y_norm[i+1]]
                    for p in peaks:
                        x_shift = (idx - (len(self.references) - 1) / 2.0) * xjitter
                        height = py_norm * span
                        ax.vlines(px + x_shift, base_y, base_y + direction * height,
                                  color=color, linewidth=options["reflinewidth"])
        
                # -- LÉGENDE REF  --
                base_ref_name = os.path.splitext(os.path.basename(ref_path))[0]
                label = self.custom_ref_names.get(ref_path, base_ref_name)
                label = self.commands.get(f"refname{idx+1}", label)
                ref_line, = ax.plot([], [], color=color, label=label, linewidth=options["legendlinewidthref"])
                ref_handles.append(ref_line)
                ref_labels.append(label)
        
            except Exception as e:
                self._add_error("REF", ref_path, e)
                continue

    
        # Labels / title
        ax.set_xlabel(options["xlabel"])
        ax.set_ylabel(options["ylabel"])
        ax.set_title(options["title"])
    
        # Legend
        if options["legend"]:
            handles = pattern_handles + ref_handles
            labels = pattern_labels + ref_labels
            if options["legendpos"] == "outside":
                legend = ax.legend(
                    handles=handles, labels=labels,
                    loc='upper left', bbox_to_anchor=(1.05, 1),
                    borderaxespad=0., frameon=False,
                    labelspacing=options["legend_labelspacing"]
                )
            else:
                legend = ax.legend(
                    handles=handles, labels=labels,
                    loc=options["legendpos"], frameon=False,
                    labelspacing=options["legend_labelspacing"]
                )
            for text in legend.get_texts():
                text.set_color(options["textcolor"])
    
        # Limits
        if "xlim" in self.commands:
            try:
                x1, x2 = map(float, self.commands["xlim"].split(','))
                ax.set_xlim(x1, x2)
            except Exception:
                pass
        if "ylim" in self.commands:
            try:
                y1, y2 = map(float, self.commands["ylim"].split(','))
                ax.set_ylim(y1, y2)
            except Exception:
                pass

        # --- New tick system (major/minor unified) ---
        def apply_major(axis, val):
            if val == "off":
                axis.set_major_locator(mticker.NullLocator())
            elif val == "auto":
                axis.set_major_locator(mticker.AutoLocator())
            else:  # numeric
                axis.set_major_locator(mticker.MultipleLocator(val))
        
        def apply_minor(axis, val):
            if val == "off":
                axis.set_minor_locator(mticker.NullLocator())
                axis.set_minor_formatter(mticker.NullFormatter())
            elif val == "auto":
                axis.set_minor_locator(mticker.AutoMinorLocator())
                axis.set_minor_formatter(mticker.NullFormatter())
            else:
                axis.set_minor_locator(mticker.MultipleLocator(val))
                axis.set_minor_formatter(mticker.NullFormatter())
        
        apply_major(ax.xaxis, options["xtick_major"])
        apply_major(ax.yaxis, options["ytick_major"])
        apply_minor(ax.xaxis, options["xtick_minor"])
        apply_minor(ax.yaxis, options["ytick_minor"])
        
        # --- Tick mark style (major/minor) ---
        # Use the same color as square_color; minor are shorter and slightly thinner.
        ax.tick_params(axis='both', which='major',
                       color=options["square_color"],
                       width=options["square_width"],
                       length=6)
        
        ax.tick_params(axis='both', which='minor',
                       color=options["square_color"],
                       width=max(0.8, options["square_width"] * 0.8),
                       length=3)

        # After limits are set, remount the cursor if enabled and sync slider range
        self._remount_cursor_after_clear()
        self._kill_mpl_keys()
        self._flush_errors()
        self.canvas.draw()        

if __name__ == "__main__":
    root = tk.Tk()  # simple Tk root; no DnD
    app = Plotter(root)
    root.mainloop()
