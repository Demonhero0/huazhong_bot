#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Steel Wire Supplier Payment Entry Script
Append a supplier-side payment event into the order workbook.

Usage:
  python3 src/supplier_payment_entry.py -s <supplier> -a <amount> [--payment-date <YYYY.MM.DD>]
"""

import argparse
import os
import sys
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP

from workbook_query_utils import ORDER_FILE, load_order_workbook


DATE_FMT = '%Y.%m.%d'


def format_amount(amount):
    return f'{amount.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP):f}'


def find_last_data_row(ws):
    for row in range(ws.max_row, 0, -1):
        for col in range(1, 20):
            if ws.cell(row=row, column=col).value is not None:
                return row
    return 4


def enter_supplier_payment(supplier, amount, payment_date=None):
    if not payment_date:
        payment_date = datetime.now().strftime(DATE_FMT)

    if amount <= Decimal('0'):
        print('\n❌ Error: Payment amount must be greater than 0!')
        return False

    if not os.path.exists(ORDER_FILE):
        print(f'\n❌ Error: Order file not found: {ORDER_FILE}')
        return False

    wb = load_order_workbook()
    if supplier not in wb.sheetnames:
        print(f'\n❌ Error: Supplier worksheet "{supplier}" not found in order file!')
        print(f'Available worksheets: {", ".join(wb.sheetnames)}')
        return False

    ws = wb[supplier]
    prev_row = find_last_data_row(ws)
    target_row = prev_row + 1
    previous_balance = ws.cell(row=prev_row, column=15).value

    ws.cell(row=target_row, column=13, value=payment_date)
    ws.cell(row=target_row, column=14, value=float(amount))

    if previous_balance is not None:
        ws.cell(row=target_row, column=15, value=f'=O{prev_row}+L{target_row}-N{target_row}')
    else:
        ws.cell(row=target_row, column=15, value=f'=L{target_row}-N{target_row}')

    wb.save(ORDER_FILE)

    print('\n=== Supplier Payment Entry ===')
    print(f'Supplier: {supplier}')
    print(f'Target Row: {target_row}')
    print(f'Payment Date: {payment_date}')
    print(f'Payment Amount: {format_amount(amount)}')
    if previous_balance is not None:
        print(f'Balance Formula: =O{prev_row}+L{target_row}-N{target_row}')
    else:
        print(f'Balance Formula: =L{target_row}-N{target_row}')
    print(f'\n✅ Order file updated: {ORDER_FILE}')
    print('✅ Supplier payment entry completed!')
    return True


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Steel Wire Supplier Payment Entry Script',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
Examples:
  python3 src/supplier_payment_entry.py -s "厦门集金" -a 397.9
  python3 src/supplier_payment_entry.py -s "厦门集金" -a 397.9 --payment-date 2026.03.27
        ''',
    )

    parser.add_argument('-s', '--supplier', required=True, help='Supplier worksheet name in the order workbook.')
    parser.add_argument('-a', '--amount', required=True, type=Decimal, help='Payment amount for this supplier-side entry.')
    parser.add_argument('--payment-date', help='Payment date, format YYYY.MM.DD. Defaults to today.')

    arguments = parser.parse_args()
    success = enter_supplier_payment(
        supplier=arguments.supplier,
        amount=arguments.amount,
        payment_date=arguments.payment_date,
    )
    sys.exit(0 if success else 1)
