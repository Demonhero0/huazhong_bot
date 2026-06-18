#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Steel Wire Delivery Entry Script
Fill delivery execution details into both the supplier workbook and the customer workbook.

Usage:
  python3 delivery_entry.py -s <supplier> -o <order_contract> -c <customer> -n <sales_contract> --pickup-date <pickup_date> --truck-no <truck_no> --factory-weight <factory_weight> --received-weight <received_weight> [--fleet <fleet>] [--freight <freight>] [--freight-tax <freight_tax>] [--dock <dock>] [--delivery-date <delivery_date>]

Examples:
  python3 delivery_entry.py -s "浙江凯航" -o 2026032401 -c "东莞建安" -n 2026032401 --pickup-date 2026.03.24 --truck-no 8888 --factory-weight 31.25 --received-weight 31.20 --fleet "英杰运输" --freight "35元/吨" --freight-tax 含税
  python3 delivery_entry.py -s "浙江凯航" -o 2026032401 -c "东莞建安" -n 2026032401 --pickup-date 2026.03.24 --delivery-date 2026.03.25 --truck-no 8888 --factory-weight 31.25 --received-weight 31.20 --fleet "英杰运输" --freight "1200元" --freight-tax 不含税 --dock 旧港仓
"""

import argparse
import os
import re

from workbook_query_utils import (
    ORDER_FILE,
    SALES_FILE,
    get_order_sheet_layout,
    load_order_workbook,
    load_sales_workbook,
    resolve_order_row_unit_price,
    resolve_sales_row_sell_price,
    resolve_sales_sheet_name,
)


def is_blank(value):
    return value is None or value == ''


def parse_freight_input(freight_text, received_weight):
    """
    Accept either 'X元/吨' or 'X元' and return the actual freight amount.
    """
    if freight_text is None:
        raise ValueError('Freight input is required.')

    text = str(freight_text).strip().replace(' ', '')

    per_ton_match = re.fullmatch(r'([0-9]+(?:\.[0-9]+)?)元/吨', text)
    if per_ton_match:
        unit_price = float(per_ton_match.group(1))
        return round(unit_price * received_weight, 2), 'per_ton'

    total_match = re.fullmatch(r'([0-9]+(?:\.[0-9]+)?)元', text)
    if total_match:
        return float(total_match.group(1)), 'fixed_total'

    raise ValueError('Freight must be in the form "X元/吨" or "X元".')


def is_direct_delivery_text(value):
    text = str(value or '').strip()
    return '直送' in text


def find_pending_order_row(ws, contract_no):
    """
    Find the first supplier workbook row for this order contract where pickup execution
    details have not been filled yet.
    """
    layout = get_order_sheet_layout(ws)
    for row in range(5, ws.max_row + 1):
        if str(ws.cell(row=row, column=2).value) != str(contract_no):
            continue

        pickup_date = ws.cell(row=row, column=layout['pickup_date']).value
        truck_no = ws.cell(row=row, column=layout['truck_no']).value
        factory_weight = ws.cell(row=row, column=layout['factory_weight']).value

        if is_blank(pickup_date) and is_blank(truck_no) and is_blank(factory_weight):
            return row
    return None


def find_pending_sales_row(ws, contract_no):
    """
    Find the first customer workbook row for this sales contract where delivery execution
    details have not been filled yet.
    """
    for row in range(5, ws.max_row + 1):
        if str(ws.cell(row=row, column=2).value) != str(contract_no):
            continue

        delivery_date = ws.cell(row=row, column=8).value  # H 送货日期
        truck_no = ws.cell(row=row, column=16).value      # P 车号
        factory_weight = ws.cell(row=row, column=17).value  # Q 出厂吨数
        received_weight = ws.cell(row=row, column=18).value  # R 实收吨数

        if is_blank(delivery_date) and is_blank(truck_no) and is_blank(factory_weight) and is_blank(received_weight):
            return row
    return None


def enter_delivery(
    supplier,
    order_contract,
    customer,
    sales_contract,
    pickup_date,
    truck_no,
    factory_weight,
    received_weight,
    fleet,
    freight,
    freight_tax,
    dock=None,
    delivery_date=None,
):
    """
    Fill one delivery execution record into both workbooks.
    """
    if not delivery_date:
        delivery_date = pickup_date

    if not os.path.exists(ORDER_FILE):
        print(f'\n❌ Error: Order file not found: {ORDER_FILE}')
        return False

    if not os.path.exists(SALES_FILE):
        print(f'\n❌ Error: Sales file not found: {SALES_FILE}')
        return False

    wb_order = load_order_workbook()
    wb_sales = load_sales_workbook()

    if supplier not in wb_order.sheetnames:
        print(f'\n❌ Error: Supplier worksheet "{supplier}" not found in order file!')
        print(f'Available worksheets: {", ".join(wb_order.sheetnames)}')
        return False

    try:
        resolved_customer, match_mode = resolve_sales_sheet_name(wb_sales, customer)
    except KeyError:
        print(f'\n❌ Error: Customer worksheet "{customer}" not found in sales file!')
        print(f'Available worksheets: {", ".join(wb_sales.sheetnames)}')
        return False

    ws_order = wb_order[supplier]
    ws_sales = wb_sales[resolved_customer]
    order_layout = get_order_sheet_layout(ws_order)

    order_row = find_pending_order_row(ws_order, order_contract)
    if order_row is None:
        print(f'\n❌ Error: No pending supplier row found for supplier "{supplier}" contract "{order_contract}"!')
        print('The contract may not exist, or all rows under it may already have pickup details.')
        return False

    sales_row = find_pending_sales_row(ws_sales, sales_contract)
    if sales_row is None:
        print(f'\n❌ Error: No pending sales row found for customer "{customer}" contract "{sales_contract}"!')
        print('The contract may not exist, or all rows under it may already have delivery details.')
        return False

    is_direct_delivery = (
        is_direct_delivery_text(ws_order.cell(row=order_row, column=7).value)
        or is_direct_delivery_text(ws_sales.cell(row=sales_row, column=14).value)
    )

    if freight in (None, ''):
        if not is_direct_delivery:
            print('\n❌ Error: Freight is required for non-direct-delivery rows!')
            return False
        actual_freight = None
        freight_mode = 'blank'
    else:
        try:
            actual_freight, freight_mode = parse_freight_input(freight, received_weight)
        except ValueError as exc:
            print(f'\n❌ Error: {exc}')
            return False

    if freight_tax in (None, ''):
        if not is_direct_delivery and freight is not None:
            print('\n❌ Error: Freight tax flag must be "含税" or "不含税" for non-direct-delivery rows!')
            return False
    elif freight_tax not in ['含税', '不含税']:
        print('\n❌ Error: Freight tax flag must be "含税" or "不含税"!')
        return False

    if fleet in (None, '') and not is_direct_delivery:
        print('\n❌ Error: Fleet is required for non-direct-delivery rows!')
        return False

    print('\n=== Delivery Entry ===')
    print(f'Supplier: {supplier}')
    print(f'Order Contract: {order_contract} -> Order Row {order_row}')
    print(f'Customer: {customer}')
    if resolved_customer != customer:
        print(f'Resolved Customer Sheet: {resolved_customer} ({match_mode})')
    print(f'Sales Contract: {sales_contract} -> Sales Row {sales_row}')
    print(f'Pickup Date: {pickup_date}')
    print(f'Delivery Date: {delivery_date}')
    print(f'Truck No: {truck_no}')
    print(f'Factory Weight: {factory_weight}')
    print(f'Received Weight: {received_weight}')
    print(f'Direct Delivery: {"yes" if is_direct_delivery else "no"}')
    print(f'Fleet: {fleet}')
    print(f'Freight Input: {freight}')
    print(f'Freight Mode: {freight_mode}')
    print(f'Actual Freight: {actual_freight}')
    print(f'Freight Tax: {freight_tax}')
    if dock:
        print(f'Dock: {dock}')

    # Supplier workbook: E 提货码头, H 提货日期, I 车号, J 出厂吨数
    if dock is not None:
        ws_order.cell(row=order_row, column=5, value=dock)
    ws_order.cell(row=order_row, column=order_layout['pickup_date'], value=pickup_date)
    ws_order.cell(row=order_row, column=order_layout['truck_no'], value=truck_no)
    ws_order.cell(row=order_row, column=order_layout['factory_weight'], value=factory_weight)
    order_unit_price = resolve_order_row_unit_price(ws_order, order_row)
    if order_unit_price is not None:
        ws_order.cell(row=order_row, column=order_layout['unit_price'], value=order_unit_price)
    factory_letter = chr(64 + order_layout['factory_weight'])
    unit_letter = chr(64 + order_layout['unit_price'])
    ws_order.cell(
        row=order_row,
        column=order_layout['pickup_amount'],
        value=f'={factory_letter}{order_row}*{unit_letter}{order_row}',
    )

    # Customer workbook: H 送货日期, I 车队, J 运费, K 运费是否含税, P 车号, Q 出厂吨数, R 实收吨数
    ws_sales.cell(row=sales_row, column=8, value=delivery_date)
    ws_sales.cell(row=sales_row, column=9, value=fleet)
    ws_sales.cell(row=sales_row, column=10, value=actual_freight)
    ws_sales.cell(row=sales_row, column=11, value=freight_tax)
    ws_sales.cell(row=sales_row, column=16, value=truck_no)
    ws_sales.cell(row=sales_row, column=17, value=factory_weight)
    ws_sales.cell(row=sales_row, column=18, value=received_weight)
    if order_unit_price is not None:
        ws_sales.cell(row=sales_row, column=13, value=order_unit_price)
    sell_price = resolve_sales_row_sell_price(ws_sales, sales_row)
    if sell_price is not None:
        ws_sales.cell(row=sales_row, column=19, value=sell_price)
    ws_sales.cell(row=sales_row, column=20, value=f'=R{sales_row}*S{sales_row}')

    wb_order.save(ORDER_FILE)
    wb_sales.save(SALES_FILE)

    print(f'\n✅ Order file updated: {ORDER_FILE}')
    print(f'✅ Sales file updated: {SALES_FILE}')
    print('✅ Delivery entry completed!')
    return True


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Steel Wire Delivery Entry Script',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
Examples:
  python3 delivery_entry.py -s "浙江凯航" -o 2026032401 -c "东莞建安" -n 2026032401 --pickup-date 2026.03.24 --truck-no 8888 --factory-weight 31.25 --received-weight 31.20 --fleet "英杰运输" --freight "35元/吨" --freight-tax 含税
  python3 delivery_entry.py -s "浙江凯航" -o 2026032401 -c "东莞建安" -n 2026032401 --pickup-date 2026.03.24 --delivery-date 2026.03.25 --truck-no 8888 --factory-weight 31.25 --received-weight 31.20 --fleet "英杰运输" --freight "1200元" --freight-tax 不含税 --dock 旧港仓
        '''
    )

    parser.add_argument('-s', '--supplier', required=True, help='Supplier worksheet name in the order workbook (e.g., "浙江凯航").')
    parser.add_argument('-o', '--order-contract', required=True, help='Order contract number in the supplier workbook (e.g., "2026032401").')
    parser.add_argument('-c', '--customer', required=True, help='Customer worksheet name in the sales workbook (e.g., "东莞建安").')
    parser.add_argument('-n', '--sales-contract', required=True, help='Sales contract number in the customer workbook (e.g., "2026032401").')
    parser.add_argument('--pickup-date', required=True, help='Pickup date for the supplier workbook (e.g., "2026.03.24").')
    parser.add_argument('--delivery-date', help='Delivery date for the customer workbook. Defaults to pickup date if omitted.')
    parser.add_argument('--truck-no', required=True, help='Truck number / plate number.')
    parser.add_argument('--factory-weight', type=float, required=True, help='Factory weight / 出厂吨数.')
    parser.add_argument('--received-weight', type=float, required=True, help='Received weight / 实收吨数.')
    parser.add_argument('--fleet', help='Fleet / 车队 (e.g., "英杰运输"). Optional for direct-delivery rows.')
    parser.add_argument('--freight', help='Freight input, either "X元/吨" or "X元". Optional for direct-delivery rows.')
    parser.add_argument('--freight-tax', help='Freight tax flag: 含税 or 不含税. Optional for direct-delivery rows.')
    parser.add_argument('--dock', help='Pickup dock / 提货码头. Optional.')

    args = parser.parse_args()

    enter_delivery(
        supplier=args.supplier,
        order_contract=args.order_contract,
        customer=args.customer,
        sales_contract=args.sales_contract,
        pickup_date=args.pickup_date,
        truck_no=args.truck_no,
        factory_weight=args.factory_weight,
        received_weight=args.received_weight,
        fleet=args.fleet,
        freight=args.freight,
        freight_tax=args.freight_tax,
        dock=args.dock,
        delivery_date=args.delivery_date,
    )
