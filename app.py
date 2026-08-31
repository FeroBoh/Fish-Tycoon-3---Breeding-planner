"""
app.py
------
Tkinter desktop GUI - Fish Tycoon 3 Breeding Helper.

Layout: one tab per map (Freshwater / Saltwater / Magical - whichever ones
currently have matching sheets in the Excel file). Each tab is split into two
halves:
    - LEFT:  "My Aquarium" - manage the fish you currently own on that map.
    - RIGHT: "Breeding Planner" - pick a target (body and/or fin) and get
      ranked breeding-path suggestions using only fish from that map's
      aquarium.
"""

from __future__ import annotations

import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Callable, Dict, List, Optional

from excel_loader import MAP_NAMES, MapData, load_all_maps
from inventory import load_inventory, save_inventory
from planner import plan_breeding

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_XLSX = os.path.join(HERE, "data", "Breeding_charts.xlsx")

RARITY_COLORS = {
    "Common": "#2e7d32",
    "Uncommon": "#1565c0",
    "Rare": "#c62828",
}

MAGIC_ACCENT = "#6a1b9a"
MAGIC_BG_ACTIVE = "#f3e5f5"
MAGIC_BG_INACTIVE = "#f5f5f5"

DEFAULT_MAX_GENERATIONS = 10
MAX_GENERATIONS_LIMIT = 30


class AquariumPanel(ttk.Frame):
    """Left-hand panel: manage the fish currently owned on this map."""

    def __init__(self, master, map_data: MapData, owned_fish: list, on_change: Callable[[], None]):
        super().__init__(master, padding=10)
        self.map_data = map_data
        self.owned = list(owned_fish)
        self.on_change = on_change

        self._build_ui()
        self._refresh_list()

    def _build_ui(self):
        ttk.Label(self, text="My Aquarium", font=("TkDefaultFont", 12, "bold")).pack(anchor="w", pady=(0, 8))

        form = ttk.Frame(self)
        form.pack(fill="x", pady=(0, 8))

        ttk.Label(form, text="Body:").grid(row=0, column=0, sticky="w", padx=(0, 4), pady=2)
        self.body_var = tk.StringVar()
        self.body_combo = ttk.Combobox(
            form, textvariable=self.body_var, values=self.map_data.body_species,
            state="readonly", width=20,
        )
        self.body_combo.grid(row=0, column=1, pady=2, sticky="w")
        if self.map_data.body_species:
            self.body_combo.current(0)

        ttk.Label(form, text="Fin:").grid(row=1, column=0, sticky="w", padx=(0, 4), pady=2)
        self.fin_var = tk.StringVar()
        self.fin_combo = ttk.Combobox(
            form, textvariable=self.fin_var, values=self.map_data.fin_species,
            state="readonly", width=20,
        )
        self.fin_combo.grid(row=1, column=1, pady=2, sticky="w")
        if self.map_data.fin_species:
            self.fin_combo.current(0)

        add_btn = ttk.Button(form, text="Add to aquarium", command=self._add_fish)
        add_btn.grid(row=2, column=0, columnspan=2, pady=(6, 0), sticky="we")

        ttk.Label(self, text="Fish I currently own:").pack(anchor="w")

        list_frame = ttk.Frame(self)
        list_frame.pack(fill="both", expand=True, pady=(4, 4))

        self.listbox = tk.Listbox(list_frame, height=16, activestyle="dotbox")
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=scrollbar.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        remove_btn = ttk.Button(self, text="Remove selected fish", command=self._remove_selected)
        remove_btn.pack(anchor="e", pady=(4, 0))

    def _add_fish(self):
        body = self.body_var.get()
        fin = self.fin_var.get()
        if not body or not fin:
            return
        self.owned.append((body, fin))
        self._refresh_list()
        self.on_change()

    def _remove_selected(self):
        sel = self.listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        del self.owned[idx]
        self._refresh_list()
        self.on_change()

    def _refresh_list(self):
        self.listbox.delete(0, "end")
        for body, fin in self.owned:
            self.listbox.insert("end", f"{body}  /  {fin}")


class PlannerPanel(ttk.Frame):
    """Right-hand panel: pick a target and find breeding paths for this map."""

    def __init__(self, master, map_data: MapData, get_owned: Callable[[], list]):
        super().__init__(master, padding=10)
        self.map_data = map_data
        self.get_owned = get_owned

        self._build_ui()

    def _build_ui(self):
        ttk.Label(self, text="Breeding Planner", font=("TkDefaultFont", 12, "bold")).pack(anchor="w", pady=(0, 8))

        form = ttk.Frame(self)
        form.pack(fill="x", pady=(0, 4))

        self.body_enabled = tk.BooleanVar(value=True)
        self.fin_enabled = tk.BooleanVar(value=True)

        self.body_check_widget = ttk.Checkbutton(
            form, text="Target body:", variable=self.body_enabled, command=self._update_enabled_state,
        )
        self.body_check_widget.grid(row=0, column=0, sticky="w", pady=2)
        self.target_body_var = tk.StringVar()
        self.target_body_combo = ttk.Combobox(
            form, textvariable=self.target_body_var, values=self.map_data.body_species,
            state="readonly", width=20,
        )
        self.target_body_combo.grid(row=0, column=1, sticky="w", padx=(4, 0), pady=2)
        if self.map_data.body_species:
            self.target_body_combo.current(0)

        self.fin_check_widget = ttk.Checkbutton(
            form, text="Target fin:", variable=self.fin_enabled, command=self._update_enabled_state,
        )
        self.fin_check_widget.grid(row=1, column=0, sticky="w", pady=2)
        self.target_fin_var = tk.StringVar()
        self.target_fin_combo = ttk.Combobox(
            form, textvariable=self.target_fin_var, values=self.map_data.fin_species,
            state="readonly", width=20,
        )
        self.target_fin_combo.grid(row=1, column=1, sticky="w", padx=(4, 0), pady=2)
        if self.map_data.fin_species:
            self.target_fin_combo.current(0)

        hint = ttk.Label(
            form,
            text="Uncheck one side to search for ANY fish matching just the\n"
                 "other attribute (e.g. any fin on a given target body).",
            foreground="#777777", font=("TkDefaultFont", 8),
        )
        hint.grid(row=2, column=0, columnspan=2, sticky="w", pady=(4, 0))

        # --- Magic Fish shortcut (to the right of the body/fin form) ---
        self.magic_frame = tk.Frame(
            form, highlightbackground=MAGIC_ACCENT, highlightthickness=2,
            bg=MAGIC_BG_INACTIVE, padx=8, pady=6,
        )
        self.magic_frame.grid(row=0, column=2, rowspan=3, sticky="ns", padx=(20, 0))

        self.magic_enabled = tk.BooleanVar(value=False)
        self.magic_check = tk.Checkbutton(
            self.magic_frame, text="✨ Magic Fish", variable=self.magic_enabled,
            command=self._on_magic_toggle, bg=MAGIC_BG_INACTIVE,
            font=("TkDefaultFont", 9, "bold"), fg=MAGIC_ACCENT,
            activebackground=MAGIC_BG_INACTIVE, selectcolor=MAGIC_BG_ACTIVE,
        )
        self.magic_check.pack(anchor="w")

        self.magic_titles = [mf.title for mf in self.map_data.magic_fish]
        self.magic_var = tk.StringVar()
        self.magic_combo = ttk.Combobox(
            self.magic_frame, textvariable=self.magic_var, values=self.magic_titles,
            state="disabled", width=24,
        )
        self.magic_combo.pack(anchor="w", pady=(4, 0))
        if self.magic_titles:
            self.magic_combo.current(0)
        self.magic_combo.bind("<<ComboboxSelected>>", lambda e: self._apply_magic_selection())

        if not self.magic_titles:
            self.magic_check.configure(state="disabled")
            ttk.Label(
                self.magic_frame, text="(no magic fish defined for this map)",
                background=MAGIC_BG_INACTIVE, foreground="#999999", font=("TkDefaultFont", 8),
            ).pack(anchor="w", pady=(2, 0))

        gen_frame = ttk.Frame(self)
        gen_frame.pack(fill="x", pady=(10, 4))
        ttk.Label(gen_frame, text="Max breeding generations to simulate:").pack(side="left")
        self.max_gen_var = tk.IntVar(value=DEFAULT_MAX_GENERATIONS)
        gen_spin = ttk.Spinbox(
            gen_frame, from_=1, to=MAX_GENERATIONS_LIMIT, textvariable=self.max_gen_var, width=5,
        )
        gen_spin.pack(side="left", padx=(6, 0))
        ttk.Label(
            gen_frame, text="(higher = explores further, may take longer)",
            foreground="#777777", font=("TkDefaultFont", 8),
        ).pack(side="left", padx=(6, 0))

        plan_btn = ttk.Button(self, text="Plan breeding", command=self._plan)
        plan_btn.pack(anchor="w", pady=(8, 8))

        info = ttk.Label(
            self,
            text=("Shows up to 5 alternative breeding sequences, ranked from most "
                  "reliable (Common) to rarest (Rare) combination. Rarity in "
                  "brackets = [body, fin] of that step's cross."),
            wraplength=420, foreground="#555555",
        )
        info.pack(fill="x", pady=(0, 8))

        self.output = tk.Text(self, wrap="word", height=20)
        self.output.pack(fill="both", expand=True)
        self.output.configure(state="disabled")

        for tag, color in RARITY_COLORS.items():
            self.output.tag_configure(tag, foreground=color, font=("TkDefaultFont", 10, "bold"))
        self.output.tag_configure("heading", font=("TkDefaultFont", 11, "bold"))
        self.output.tag_configure(
            "magic_banner", foreground=MAGIC_ACCENT, font=("TkDefaultFont", 12, "bold"),
        )

    def _update_enabled_state(self):
        self.target_body_combo.configure(state="readonly" if self.body_enabled.get() else "disabled")
        self.target_fin_combo.configure(state="readonly" if self.fin_enabled.get() else "disabled")

    def _on_magic_toggle(self):
        if self.magic_enabled.get():
            # Lock the body/fin checkboxes on and disable them so the user
            # can't accidentally uncheck one - the magic fish target always
            # pins both attributes at once.
            self.body_enabled.set(True)
            self.fin_enabled.set(True)
            self.body_check_widget.configure(state="disabled")
            self.fin_check_widget.configure(state="disabled")
            self.magic_combo.configure(state="readonly")
            self.magic_frame.configure(bg=MAGIC_BG_ACTIVE)
            self.magic_check.configure(bg=MAGIC_BG_ACTIVE, activebackground=MAGIC_BG_ACTIVE)
            self._apply_magic_selection()
        else:
            self.body_check_widget.configure(state="normal")
            self.fin_check_widget.configure(state="normal")
            self.target_body_combo.configure(values=self.map_data.body_species, state="readonly")
            self.target_fin_combo.configure(values=self.map_data.fin_species, state="readonly")
            self._update_enabled_state()
            self.magic_frame.configure(bg=MAGIC_BG_INACTIVE)
            self.magic_check.configure(bg=MAGIC_BG_INACTIVE, activebackground=MAGIC_BG_INACTIVE)

    def _apply_magic_selection(self):
        if not self.magic_enabled.get():
            return
        title = self.magic_var.get()
        match = next((mf for mf in self.map_data.magic_fish if mf.title == title), None)
        if match is None:
            return
        # Restrict the target dropdowns to just this magic fish's body/fin
        # and keep them disabled so nothing else can be picked by mistake.
        self.target_body_combo.configure(values=[match.body], state="disabled")
        self.target_fin_combo.configure(values=[match.fin], state="disabled")
        self.target_body_var.set(match.body)
        self.target_fin_var.set(match.fin)

    def _plan(self):
        use_body = self.body_enabled.get()
        use_fin = self.fin_enabled.get()

        if not use_body and not use_fin:
            messagebox.showwarning(
                "No target selected",
                "Check at least one of 'Target body' or 'Target fin'.",
            )
            return

        target_body = self.target_body_var.get() if use_body else None
        target_fin = self.target_fin_var.get() if use_fin else None

        if use_body and not target_body:
            messagebox.showwarning("Missing target", "Pick a target body.")
            return
        if use_fin and not target_fin:
            messagebox.showwarning("Missing target", "Pick a target fin.")
            return

        owned = self.get_owned()
        if not owned:
            messagebox.showinfo(
                "Empty aquarium",
                "Your aquarium for this map is empty. Add at least one fish "
                "on the left before planning.",
            )
            return

        try:
            max_generations = int(self.max_gen_var.get())
        except (tk.TclError, ValueError):
            max_generations = DEFAULT_MAX_GENERATIONS
        max_generations = max(1, min(MAX_GENERATIONS_LIMIT, max_generations))

        self.output.configure(state="normal")
        self.output.delete("1.0", "end")

        if self.magic_enabled.get() and self.magic_var.get():
            self.output.insert("end", f"✨ {self.magic_var.get()} ✨\n", "magic_banner")
            self.output.insert(
                "end", f"({target_body} / {target_fin})\n\n", "heading",
            )

        if target_body is not None and target_fin is not None and (target_body, target_fin) in owned:
            self.output.insert("end", "You already own this exact fish.\n")
            self.output.configure(state="disabled")
            return

        recipes = plan_breeding(
            self.map_data, owned,
            target_body=target_body, target_fin=target_fin,
            top_n=5, max_generations=max_generations,
        )

        if not recipes:
            self.output.insert(
                "end",
                "I couldn't reach a matching fish from your current aquarium "
                f"within {max_generations} generation(s). Try raising 'Max "
                "breeding generations', adding more/different fish to your "
                "aquarium, or buying some at the market.\n",
            )
            self.output.configure(state="disabled")
            return

        for i, recipe in enumerate(recipes, start=1):
            self.output.insert("end", f"Variant {i}", "heading")
            self.output.insert("end", f"  ->  {recipe.result_fish[0]} / {recipe.result_fish[1]}", "heading")
            self.output.insert("end", "   —   overall rarity: ", "heading")
            self.output.insert("end", f"{recipe.overall_rarity}\n", recipe.overall_rarity)
            if not recipe.steps:
                self.output.insert("end", "  (you already own this fish)\n\n")
                continue
            for j, step in enumerate(recipe.steps, start=1):
                self.output.insert("end", f"  {j}. ")
                self.output.insert("end", f"{step.parent_a[0]} / {step.parent_a[1]}")
                self.output.insert("end", "  x  ")
                self.output.insert("end", f"{step.parent_b[0]} / {step.parent_b[1]}")
                self.output.insert("end", "  ->  ")
                self.output.insert("end", f"{step.result[0]} / {step.result[1]}   ")
                self.output.insert("end", "[body: ")
                self.output.insert("end", step.body_rarity, step.body_rarity)
                self.output.insert("end", ", fin: ")
                self.output.insert("end", step.fin_rarity, step.fin_rarity)
                self.output.insert("end", "]\n")
            self.output.insert("end", "\n")

        self.output.configure(state="disabled")


class MapPage(ttk.Frame):
    """A single map's tab: aquarium panel on the left, planner on the right."""

    def __init__(self, master, map_data: MapData, owned_fish: list, on_inventory_change: Callable[[], None]):
        super().__init__(master)
        self.map_data = map_data

        paned = ttk.PanedWindow(self, orient="horizontal")
        paned.pack(fill="both", expand=True)

        self.aquarium = AquariumPanel(paned, map_data, owned_fish, on_change=self._on_change)
        self.planner = PlannerPanel(paned, map_data, get_owned=lambda: self.aquarium.owned)

        paned.add(self.aquarium, weight=1)
        paned.add(self.planner, weight=1)

        self._on_inventory_change = on_inventory_change

    def _on_change(self):
        self._on_inventory_change()

    @property
    def owned(self) -> list:
        return self.aquarium.owned


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Fish Tycoon 3 — Breeding Planning Helper")
        self.geometry("1200x700")

        self.xlsx_path = DEFAULT_XLSX
        self.maps: Dict[str, MapData] = {}
        self.inventory: Dict[str, list] = load_inventory()

        self._build_menu()

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True)

        self.map_pages: Dict[str, MapPage] = {}

        self._load_and_rebuild()

    def _build_menu(self):
        menubar = tk.Menu(self)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Choose a different Excel file...", command=self._choose_file)
        file_menu.add_command(label="Reload Excel (e.g. after adding sheets)", command=self._reload)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.destroy)
        menubar.add_cascade(label="File", menu=file_menu)
        self.config(menu=menubar)

    def _choose_file(self):
        path = filedialog.askopenfilename(
            title="Select Breeding_charts.xlsx",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if path:
            self.xlsx_path = path
            self._load_and_rebuild()

    def _reload(self):
        self._load_and_rebuild()

    def _load_and_rebuild(self):
        try:
            self.maps = load_all_maps(self.xlsx_path)
        except Exception as e:
            messagebox.showerror("Failed to load file", str(e))
            return

        if not self.maps:
            messagebox.showwarning(
                "No data found",
                "I couldn't find any complete sheet pairs '<Map>_FINS' + "
                "'<Map>_BODIES' in the Excel file.",
            )

        for tab in self.notebook.tabs():
            self.notebook.forget(tab)
        self.map_pages = {}

        for map_name in MAP_NAMES:
            if map_name not in self.maps:
                continue
            owned = self.inventory.get(map_name, [])
            page = MapPage(self.notebook, self.maps[map_name], owned, on_inventory_change=self._save_all)
            self.map_pages[map_name] = page
            self.notebook.add(page, text=map_name)

    def _save_all(self):
        data = {name: page.owned for name, page in self.map_pages.items()}
        self.inventory = data
        save_inventory(data)


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
