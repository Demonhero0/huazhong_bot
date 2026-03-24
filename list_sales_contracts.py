#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import json
import sys

from workbook_query_utils import SALES_FILE, get_sales_sheet, load_sales_workbook, resolve_sales_sheet_name, summarize_sales_contracts


def main():
    parser = argparse.ArgumentParser(description='List customer-side sales contracts as JSON.')
    parser.add_argument('-c', '--customer', required=True, help='Customer worksheet name.')
    args = parser.parse_args()

    try:
        wb = load_sales_workbook()
        resolved_customer, match_mode = resolve_sales_sheet_name(wb, args.customer)
        ws = get_sales_sheet(wb, args.customer)
    except KeyError as exc:
        print(json.dumps({'ok': False, 'error': str(exc), 'workbook': SALES_FILE}, ensure_ascii=False, indent=2))
        sys.exit(1)

    print(json.dumps({
        'ok': True,
        'workbook': SALES_FILE,
        'customer': args.customer,
        'resolved_customer': resolved_customer,
        'match_mode': match_mode,
        'contracts': summarize_sales_contracts(ws),
    }, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
