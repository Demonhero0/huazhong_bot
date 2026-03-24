---
name: steel-wire-order-sales
description: Use when working in this repository on the steel wire Excel workflow for 订货, 销货, 送货, 收款, and 出具对账单. This skill maps natural-language workflow requests to the local scripts and read tools in this repository.
---

# Steel Wire Order Sales

## Overview

Use this skill for the local steel-wire ledger workflow in this repo. It converts user requests about `订货`, `销货`, `送货`, `收款`, and `出具对账单` into concrete commands against the existing Python scripts and Excel files.

Read [references/workflow.md](references/workflow.md) when you need the field mapping, workbook columns, or example commands.

## Workflow Decision

- If the user wants to create or append a supplier order, use the `订货` workflow with `order_entry.py`.
- If the user wants to allocate trucks from an existing supplier contract to a customer, use the `销货` workflow with `sales_entry.py`.
- If the user wants to fill actual truck execution details into both workbooks, use the `送货` workflow with `delivery_entry.py`.
- If the user wants to register customer receipts into the sales workbook, use the `收款` workflow with `receipt_entry.py`.
- If the user wants to generate a customer statement, use the `出具对账单` workflow with `statement_issue.py`. This workflow exports Excel only.
- If the user gives incomplete `销货` data, do not guess missing values such as supplier, contract number, selling price, or delivery mode. Tell the user exactly which fields are missing.

## Environment

- Run commands from the repository root.
- Prefer `./.venv/bin/python` instead of `python3`.
- If dependencies are missing, install with `./.venv/bin/python -m pip install -r requirements.txt`.
- Prefer the JSON read tools before any write that depends on locating contracts or pending rows.

## Read Tools

Use these tools first when the user needs contract discovery or row-level targeting:

- `./.venv/bin/python list_order_contracts.py -s "<supplier>"`
- `./.venv/bin/python list_sales_contracts.py -c "<customer>"`
- `./.venv/bin/python list_receivable_contracts.py -c "<customer>" [--settled no|yes|all]`
- `./.venv/bin/python find_pending_order_rows.py -s "<supplier>" -n <order_contract>`
- `./.venv/bin/python find_pending_sales_rows.py -c "<customer>" -n <sales_contract>`

Rules:

- These tools return JSON and are preferred over scanning Excel manually.
- Before `销货`, use `list_order_contracts.py` if the supplier contract is unknown.
- Before `送货`, use `find_pending_order_rows.py` and `find_pending_sales_rows.py` when there is any ambiguity about which row will be updated.
- Before `收款`, prefer `list_receivable_contracts.py` to confirm the contract is still unsettled and has remaining receipt slots.
- Customer sheet lookup supports exact match first, then unique normalized/fuzzy match for common short names such as `东莞建安 -> 东莞市建安管桩有限公司`.

## 订货 Workflow

Required fields:

- `supplier`
- `price`
- `trucks`
- `brand`
- `spec`
- `transport`

Command template:

```bash
./.venv/bin/python order_entry.py -s "<supplier>" -p <price> -n <trucks> -b "<brand>" -e "<spec>" -t "<transport>"
```

With explicit order date:

```bash
./.venv/bin/python order_entry.py -s "<supplier>" -p <price> -n <trucks> -b "<brand>" -e "<spec>" -t "<transport>" --order-date "<YYYY.MM.DD>"
```

With payment amount only:

```bash
./.venv/bin/python order_entry.py -s "<supplier>" -p <price> -n <trucks> -b "<brand>" -e "<spec>" -t "<transport>" -m <payment_amount>
```

With explicit payment date and amount:

```bash
./.venv/bin/python order_entry.py -s "<supplier>" -p <price> -n <trucks> -b "<brand>" -e "<spec>" -t "<transport>" -m <payment_amount> -d "<YYYY.MM.DD>"
```

Rules:

- `supplier` must match a supplier worksheet in `线材供应商提货明细龙虾版.xlsx`.
- `trucks` is the number of rows to append for the contract.
- `order_date` defaults to today and can be overridden with `--order-date`.
- If `payment_amount` is provided, `付款日期` defaults to `order_date`.
- `payment_date` only matters when `payment_amount` is provided.
- Use the exact spec string the user gives unless there is an obvious local convention already established in the sheet.
- Only include `-c`, `-d`, `-m` if the user explicitly provides customer or payment details.

## 销货 Workflow

Required fields for actual execution:

- `supplier`
- `contract_no`
- `customer`
- `price`
- `delivery`
- `trucks`

If `contract_no` is unknown, first list available contracts:

```bash
./.venv/bin/python list_order_contracts.py -s "<supplier>"
```

Execution template:

```bash
./.venv/bin/python sales_entry.py -s "<supplier>" -n <contract_no> -c "<customer>" -p <price> -d "<delivery>" -t <trucks>
```

With explicit sales date:

```bash
./.venv/bin/python sales_entry.py -s "<supplier>" -n <contract_no> -c "<customer>" -p <price> -d "<delivery>" -t <trucks> --sales-date "<YYYY.MM.DD>"
```

With optional benchmark and price diff:

```bash
./.venv/bin/python sales_entry.py -s "<supplier>" -n <contract_no> -c "<customer>" -p <price> -d "<delivery>" -t <trucks> --benchmark "<benchmark>" --price-diff "<price_diff>"
```

Rules:

- `delivery` must be `自提` or `送到`.
- `sales_date` defaults to today and can be overridden with `--sales-date`.
- `benchmark` maps directly to the sales workbook `对标(F)` column if provided.
- `price_diff` maps directly to the sales workbook `成交价差(G)` column if provided.
- Do not infer `benchmark` or `price_diff`; only write them when the user explicitly provides them.
- Customer names may be short inputs; the script resolves them to an existing customer sheet when there is a unique fuzzy match.
- Sales workbook contract numbers are generated from `sales_date` in `YYYYMMDD01` format.
- `trucks` cannot exceed the number of empty customer rows available under that supplier contract.
- `sales_entry.py` writes to both the supplier workbook and the customer workbook. Treat it as a two-file update.

## 送货 Workflow

Required fields for actual execution:

- `supplier`
- `order_contract`
- `customer`
- `sales_contract`
- `pickup_date`
- `truck_no`
- `factory_weight`
- `received_weight`
- `fleet`
- `freight`
- `freight_tax`

Execution template:

```bash
./.venv/bin/python delivery_entry.py -s "<supplier>" -o <order_contract> -c "<customer>" -n <sales_contract> --pickup-date "<YYYY.MM.DD>" --truck-no "<truck_no>" --factory-weight <factory_weight> --received-weight <received_weight> --fleet "<fleet>" --freight "<X元/吨|X元>" --freight-tax "<含税|不含税>"
```

With explicit dock and delivery date:

```bash
./.venv/bin/python delivery_entry.py -s "<supplier>" -o <order_contract> -c "<customer>" -n <sales_contract> --pickup-date "<YYYY.MM.DD>" --delivery-date "<YYYY.MM.DD>" --truck-no "<truck_no>" --factory-weight <factory_weight> --received-weight <received_weight> --fleet "<fleet>" --freight "<X元/吨|X元>" --freight-tax "<含税|不含税>" --dock "<dock>"
```

Rules:

- The supplier workbook is located by `supplier + order_contract`.
- The customer workbook is located by `customer + sales_contract`.
- The script fills the first pending row under each contract.
- `delivery_date` defaults to `pickup_date` if omitted.
- `dock` is optional and may be left blank.
- Customer names may be short inputs; the script resolves them to an existing customer sheet when there is a unique fuzzy match.
- `freight` accepts either `X元/吨` or `X元`.
- If `freight` is `X元/吨`, the script multiplies it by `received_weight` and writes the computed actual freight.
- `freight_tax` must be `含税` or `不含税`.
- Use the pending-row query tools first if you need to verify which row is next.

## 出具对账单 Workflow

Before running the command, remind the user that these filters are available:

- `customer` (required)
- `sales date range`
- `contract number`
- `paid status`

Paid status rule:

- `收款日期` and `已收款金额` both empty => unpaid
- otherwise => paid

Command template:

```bash
./.venv/bin/python statement_issue.py --customer "<customer_full_name>" [--date-from "<YYYY.MM.DD>"] [--date-to "<YYYY.MM.DD>"] [--contract "<contract_no>"] [--paid yes|no|all] [--statement-date "<YYYY.MM.DD>"]
```

Rules:

- `customer` is required.
- Seller is always `东莞市华众供应链管理有限公司`.
- Buyer is always the resolved full customer name.
- `statement_date` defaults to today when omitted.
- `合同签订日期` is taken from `销售日期`.
- `备注` defaults to `品牌 + 规格`.
- Output goes under `statements/<客户全称>_<销售日期或日期范围>/`.
- Exported file is named `<客户全称>_<销售日期或日期范围>_对账单.xlsx`.
- Each run exports Excel only. Do not describe or promise PDF, HTML, or JSON files for this workflow.
- If multiple contracts match and no mode is specified, the script stops and requires the user to choose `summary` or `split`.
- Tell the agent that the exported statement is a base draft in Excel and the user may request further layout/content tweaks afterward.

## 收款 Workflow

Read-first template:

```bash
./.venv/bin/python list_receivable_contracts.py -c "<customer>" [--contract "<contract_no>"] [--date-from "<YYYY.MM.DD>"] [--date-to "<YYYY.MM.DD>"] [--settled no|yes|all]
```

Execution template:

```bash
./.venv/bin/python receipt_entry.py -c "<customer>" -n <contract_no> -a <amount>
```

With explicit receipt date:

```bash
./.venv/bin/python receipt_entry.py -c "<customer>" -n <contract_no> -a <amount> --receipt-date "<YYYY.MM.DD>"
```

Rules:

- `customer`, `contract_no`, and `amount` are required for actual execution.
- `receipt_date` defaults to today when omitted.
- The script writes only `收款日期(U)` and `已收款金额(V)`.
- Do not modify `未收款金额(W)`; that formula is already in the workbook.
- The first receipt is written into the first contract row.
- Later receipts are written into the next contract rows whose `U/V` are still empty.
- If all contract rows already have receipt entries, stop and tell the user to handle it manually.
- Settlement status and outstanding-balance checks must use the contract's last row as the balance row, but the actual balance should be computed from the contract rows' sales amounts and received amounts instead of relying on stale Excel formula cache values.
- If the input amount exceeds the current outstanding balance, stop and raise an error.
- Do not guess a receipt amount, including for full settlement. The user must provide the amount explicitly.

## Response Style

- When the user asks for a command, return the exact command only, unless a required field is missing.
- When fields are missing, list the missing fields briefly and, if useful, provide the contract-list command.
- When asked to explain the process, summarize at the level of `订货 -> 供应商台账` and `销货 -> 回填供应商台账 + 写客户台账`.
