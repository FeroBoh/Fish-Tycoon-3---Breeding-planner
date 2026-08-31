# Fish Tycoon 3 — Breeding Planner

A small Tkinter desktop app for planning fish breeding in **Fish Tycoon 3**.
It reads breeding data from an Excel workbook (`Breeding_charts.xlsx`) and
figures out how to breed a target fish from the fish you currently own.

## Features

- One tab per map (Freshwater / Saltwater / Magical), each split into:
  - **My Aquarium** — track which fish you currently own
  - **Breeding Planner** — pick a target and get up to 5 ranked breeding
    paths, from most reliable (Common) to rarest (Rare)
- Target by exact fish (body + fin), by body/fin only, or by picking a named
  **✨ Magic Fish** from the workbook
- Adjustable search depth (max generations) for harder-to-reach targets
- Aquarium contents persist between runs
- New maps are picked up automatically — just add matching sheets to the
  Excel file, no code changes needed

## Setup

```bash
pip install openpyxl
python main.py
```

or

```python -m pip install openpyxl```

Requires Python 3.9+. `tkinter` ships with standard Python on
Windows/macOS; on Linux install `python3-tk` if it's missing.

## Excel file format

Each map needs two sheets: `<Map>_FINS` and `<Map>_BODIES` (a breeding
matrix: species list in row 1 / column A, results in the grid, cell color =
rarity). An optional `<Map>_MAGICFISH` sheet (columns `Fin`, `Body`,
`Magic Fish Title`) adds the named magic fish shortcut.

Put your workbook at `data/Breeding_charts.xlsx`, or point the app at a
different file via **File → Choose a different Excel file...**.

## About

Built as a personal fan tool for Fish Tycoon 3, developed with the
assistance of Claude (Anthropic).

Fish Tycoon 3 is a game by Last Day of Work; this project is unofficial fan
software and isn't affiliated with or endorsed by them.
