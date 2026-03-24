#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import json
import sys

from workbook_query_utils import SALES_FILE, find_pending_sales_rows, get_sales_sheet, load_sales_workbook, resolve_sales_sheet_name


def main():
    parser = argparse.ArgumentParser(description='Find pending customer-side delivery rows as JSON.')
    parser.add_argument('-c', '--customer', required=True, help='Customer worksheet name.')
    parser.add_argument('-n', '--contract', required=True, help='Sales contract number.')
    args = parser.parse_args()

    try:
        wb = load_sales_workbook()
        resolved_customer, match_mode = resolve_sales_sheet_name(wb, args.customer)
        ws = get_sales_sheet(wb, args.customer)
    except KeyError as exc:
        print(json.dumps({'ok': False, 'error': str(exc), 'workbook': SALES_FILE}, ensure_ascii=False, indent=2))
        sys.exit(1)

    rows = find_pending_sales_rows(ws, args.contract)
    print(json.dumps({
        'ok': True,
        'workbook': SALES_FILE,
        'customer': args.customer,
        'resolved_customer': resolved_customer,
        'match_mode': match_mode,
        'contract_no': args.contract,
        'pending_rows': rows,
        'pending_count': len(rows),
    }, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
