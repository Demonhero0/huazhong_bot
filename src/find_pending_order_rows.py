#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import json
import sys

from workbook_query_utils import ORDER_FILE, find_pending_order_rows, get_order_sheet, load_order_workbook


def main():
    parser = argparse.ArgumentParser(description='Find pending supplier-side pickup rows as JSON.')
    parser.add_argument('-s', '--supplier', required=True, help='Supplier worksheet name.')
    parser.add_argument('-n', '--contract', required=True, help='Order contract number.')
    args = parser.parse_args()

    try:
        wb = load_order_workbook()
        ws = get_order_sheet(wb, args.supplier)
    except KeyError as exc:
        print(json.dumps({'ok': False, 'error': str(exc), 'workbook': ORDER_FILE}, ensure_ascii=False, indent=2))
        sys.exit(1)

    rows = find_pending_order_rows(ws, args.contract)
    print(json.dumps({
        'ok': True,
        'workbook': ORDER_FILE,
        'supplier': args.supplier,
        'contract_no': args.contract,
        'pending_rows': rows,
        'pending_count': len(rows),
    }, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
