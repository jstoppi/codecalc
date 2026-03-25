# CodeCalc MVP

Minimal local MVP for drawing-ready building code calculations.

## Current module
- IBC / NFPA occupant load schedule generator

## Run

```powershell
cd C:\Users\john\.openclaw\workspace\codecalc
py -m pip install -r requirements.txt
py -m streamlit run app.py
```

## What it does
- Enter room story, room number, room name, room area, and occupant-load function of space.
- Choose the standard (currently IBC 2021 or NFPA 101).
- Calculates occupant load from area / factor.
- Shows total occupants per story.
- Exports a clean Excel workbook with:
  - Room schedule
  - Story totals
  - Code basis

## Notes
- Initial code table is a curated MVP set of common IBC 2021 occupant load factors.
- Results are an aid, not a substitute for professional judgment.
