#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import json
import sys

from workbook_query_utils import ORDER_FILE, get_order_sheet, load_order_workbook, summarize_order_contracts


def main():
    parser = argparse.ArgumentParser(description='List supplier-side order contracts as JSON.')
    parser.add_argument('-s', '--supplier', required=True, help='Supplier worksheet name.')
    args = parser.parse_args()

    try:
        wb = load_order_workbook()
        ws = get_order_sheet(wb, args.supplier)
    except KeyError as exc:
        print(json.dumps({'ok': False, 'error': str(exc), 'workbook': ORDER_FILE}, ensure_ascii=False, indent=2))
        sys.exit(1)

    print(json.dumps({
        'ok': True,
        'workbook': ORDER_FILE,
        'supplier': args.supplier,
        'contracts': summarize_order_contracts(ws),
    }, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
