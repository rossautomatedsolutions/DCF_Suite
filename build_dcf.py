from __future__ import annotations

import sys
from pathlib import Path

from dcf_workbook import build_workbook


def main() -> int:
    output_path = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("dcf_model.xlsx")
    workbook = build_workbook()
    workbook.save(output_path)
    print(f"Workbook saved to {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
