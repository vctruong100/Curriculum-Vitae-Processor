#!/usr/bin/env python3
import argparse
import csv
from pathlib import Path

def is_year(value: str) -> bool:
    v = str(value).strip()
    return len(v) == 4 and v.isdigit()

def main():
    parser = argparse.ArgumentParser(
        description="Convert master study list CSV (Phase/Category/Year, Red, No-Red) "
                    "into one or two MASTER text files for the sorter."
    )
    parser.add_argument("--csv", required=True, help="Path to master study list CSV")
    parser.add_argument("--out", help="Output MASTER text from Column C (backwards compatible alias for --out-c)")
    parser.add_argument("--out-c", help="Output MASTER text from Column C (no-red)")
    parser.add_argument("--out-b", help="Output MASTER text from Column B (red-label)")
    parser.add_argument("--has-header", action="store_true", help="Set if first row is a header row to skip")
    args = parser.parse_args()

    csv_path = Path(args.csv)

    if not csv_path.is_file():
        raise SystemExit(f"ERROR: CSV not found: {csv_path}")

    if not args.out_c and not args.out:
        raise SystemExit("ERROR: Either --out or --out-c must be provided.")

    out_c_path = Path(args.out_c or args.out)
    if args.out_b:
        out_b_path = Path(args.out_b)
    else:
        if out_c_path.suffix:
            out_b_path = out_c_path.with_name(out_c_path.stem + "_COLB" + out_c_path.suffix)
        else:
            out_b_path = out_c_path.with_name(out_c_path.name + "_COLB")

    out_c_path.parent.mkdir(parents=True, exist_ok=True)
    out_b_path.parent.mkdir(parents=True, exist_ok=True)

    with csv_path.open("r", encoding="utf-8-sig", newline="") as f_in, \
         out_c_path.open("w", encoding="utf-8", newline="\n") as f_c, \
         out_b_path.open("w", encoding="utf-8", newline="\n") as f_b:

        reader = csv.reader(f_in)
        first = True

        for row in reader:
            if not row:
                continue

            col_a = (row[0] or "").strip()
            col_b = (row[1] or "").strip() if len(row) > 1 else ""
            col_c = (row[2] or "").strip() if len(row) > 2 else ""

            if args.has_header and first:
                first = False
                continue
            first = False

            if not col_a and not col_b and not col_c:
                continue

            if is_year(col_a):
                year = col_a
                desc_c = col_c or col_b
                desc_b = col_b or col_c

                if desc_c:
                    f_c.write(f"{year} {desc_c}\n")
                if desc_b:
                    f_b.write(f"{year} {desc_b}\n")
            else:
                header = col_a.strip()
                if header:
                    f_c.write(header + "\n")
                    f_b.write(header + "\n")

    print(f"Wrote Column C master to: {out_c_path}")
    print(f"Wrote Column B master to: {out_b_path}")

if __name__ == "__main__":
    main()
