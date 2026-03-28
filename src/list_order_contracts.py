#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import json
import sys

from workbook_query_utils import ORDER_FILE, get_order_sheet, load_order_workbook, summarize_order_contracts


def main():
    parser = argparse.ArgumentParser(description='List supplier-side order contracts as JSON.')
    parser.add_argument('-s', '--supplier', help='Supplier worksheet name. If omitted, list contracts across all suppliers.')
    args = parser.parse_args()

    wb = load_order_workbook()

    if args.supplier:
        try:
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
        return

    suppliers = []
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        contracts = summarize_order_contracts(ws)
        pending_contracts = [contract for contract in contracts if contract['pending_customer_count'] > 0]
        if not pending_contracts:
            continue
        suppliers.append({
            'supplier': sheet_name,
            'contracts': pending_contracts,
        })

    print(json.dumps({
        'ok': True,
        'workbook': ORDER_FILE,
        'supplier': None,
        'suppliers': suppliers,
        'supplier_count': len(suppliers),
        'contract_count': sum(len(item['contracts']) for item in suppliers),
    }, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
