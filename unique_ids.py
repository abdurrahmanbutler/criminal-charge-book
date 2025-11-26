#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Check that every `id` in a JSONL file is unique.

Usage:
    python check_unique_ids.py path/to/chunks.jsonl
"""

import argparse
import json
import sys
from collections import defaultdict


def main():
    parser = argparse.ArgumentParser(
        description="Confirm that every `id` in a JSONL file is unique."
    )
    parser.add_argument(
        "jsonl_path",
        help="Path to the JSONL file to check (e.g. chunks.jsonl).",
    )
    args = parser.parse_args()

    jsonl_path = args.jsonl_path

    seen_ids = set()
    duplicates = defaultdict(list)  # id -> list of line numbers
    total_lines = 0
    total_records = 0
    missing_id_lines = []

    try:
        with open(jsonl_path, "r", encoding="utf-8") as f:
            for lineno, line in enumerate(f, start=1):
                total_lines += 1
                stripped = line.strip()
                if not stripped:
                    continue  # skip empty lines

                try:
                    obj = json.loads(stripped)
                except json.JSONDecodeError as e:
                    print(
                        f"ERROR: Invalid JSON on line {lineno}: {e}",
                        file=sys.stderr,
                    )
                    sys.exit(2)

                total_records += 1

                if "id" not in obj:
                    missing_id_lines.append(lineno)
                    continue

                _id = obj["id"]
                if _id in seen_ids:
                    duplicates[_id].append(lineno)
                else:
                    seen_ids.add(_id)

    except FileNotFoundError:
        print(f"ERROR: File not found: {jsonl_path}", file=sys.stderr)
        sys.exit(2)

    # Report results
    print(f"Checked file: {jsonl_path}")
    print(f"Total lines read: {total_lines}")
    print(f"Total JSON records: {total_records}")
    print(f"Total distinct ids: {len(seen_ids)}")

    if missing_id_lines:
        print(
            f"\nWARNING: {len(missing_id_lines)} record(s) missing `id` field "
            f"on line(s): {', '.join(map(str, missing_id_lines))}"
        )

    if duplicates:
        print("\n❌ Duplicate ids found:")
        for _id, lines in duplicates.items():
            # We know it appeared at least twice; the first appearance is
            # implied, subsequent ones are in `lines`.
            line_list = ", ".join(str(l) for l in lines)
            print(f"  id={_id!r} appears more than once (additional lines: {line_list})")
        sys.exit(1)
    else:
        print("\n✅ All ids are unique.")
        sys.exit(0)


if __name__ == "__main__":
    main()
