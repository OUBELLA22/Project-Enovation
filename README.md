# Project-Enovation

Automatic Wire Harness Routing & Generation Tool (CATIA V5).

## Documents

- `Cahier_des_Charges_Faisceau_Routing.docx` — full routing specification (Cahier des Charges).
- `build_cahier.py` — generates the routing specification.
- `Functional_Specification_AS_IS_Wire_Harness_Routing.docx` — prefilled AS-IS process-discovery document for business validation.
- `build_as_is_spec.py` — generates the AS-IS process-discovery document.

Regenerate both documents from the repository with:

```bash
python3 build_cahier.py
python3 build_as_is_spec.py
```

Both generated DOCX files are written directly into this repository. Version 1 covers mandatory obstacle and collision detection, but not automatic obstacle avoidance or route recalculation.
