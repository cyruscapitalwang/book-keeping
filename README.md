# book-keeping (uv project)

One-flag CLI to build/update your **Check Register-Corp** sheet from a Chase checking PDF.

## Usage

```bash
uv sync
uv run book-keeping --directory "C:\\Users\\admin\\CyrusCapital\\Book Keeping\\2024-07" [--dry-run] [--verbose]
```

## STRICT inputs required in the target folder
- **Excel template**: exactly `Corp Registers_.xlsx`
- **Chase PDF**: filename must contain **`2590`** (e.g., `...2590....pdf`)

## Behavior
- Parses last folder name `YYYY-MM`, copies `Corp Registers_.xlsx` to `Corp Registers_YYYYMM.xlsx` (unless `--dry-run`).
- Uses only the PDF containing `2590`; fails if not present.
- Sets B9 to the first day of that month; writes transactions; formatting & rules as before.
- Prints **section totals** from the PDF and from the data written; **asserts they match** to the cent.
- `--dry-run`: simulate, print plan & totals, **do not write files**.
- `--verbose`: print detailed steps.
