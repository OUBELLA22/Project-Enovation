# Project-Enovation

Automatic Wire Harness Routing & Generation Tool (CATIA V5).

## Documents

- `Cahier_des_Charges_Faisceau_Routing.docx` — full specification document (Cahier des Charges).
- `build_cahier.py` — script that generates the Word document.

Regenerate the document from the repository with:

```bash
python3 build_cahier.py
```

The generated DOCX is written directly into this repository. Version 1 covers mandatory obstacle and collision detection, but not automatic obstacle avoidance or route recalculation.
