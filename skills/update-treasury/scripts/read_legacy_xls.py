#!/usr/bin/env python3

import json
import sys

import xlrd


def main() -> None:
    if len(sys.argv) != 2:
        raise SystemExit("Usage: read_legacy_xls.py FILE.xls")
    workbook = xlrd.open_workbook(sys.argv[1])
    sheet = workbook.sheet_by_index(0)
    rows = [sheet.row_values(index) for index in range(sheet.nrows)]
    json.dump(rows, sys.stdout, ensure_ascii=False)


if __name__ == "__main__":
    main()
