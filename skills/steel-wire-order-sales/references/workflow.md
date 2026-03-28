# Steel Wire Workflow Reference

## Files

- `order_entry.py`: append supplier-side order rows into `线材供应商提货明细龙虾版.xlsx`
- `sales_entry.py`: fill customer allocation in the supplier workbook and append rows into `线材客户送货明细龙虾版.xlsx`
- `delivery_entry.py`: fill truck execution details into both workbooks
- `receipt_entry.py`: fill customer receipt date and received amount into the customer workbook
- `update_sales_contract.py`: backfill or modify fields in an existing customer sales contract
- `statement_issue.py`: issue customer statements and export Excel only
- `list_order_contracts.py`: list supplier-side contracts as JSON
- `list_sales_contracts.py`: list customer-side contracts as JSON
- `list_receivable_contracts.py`: list customer-side receivable contracts as JSON
- `find_pending_order_rows.py`: list supplier-side pending pickup rows as JSON
- `find_pending_sales_rows.py`: list customer-side pending delivery rows as JSON
- `requirements.txt`: currently only requires `openpyxl`

## Recommended Read-First Flow

Use the JSON query tools before any write that depends on contract lookup or pending-row selection.

Customer lookup supports exact match first, then unique normalized/fuzzy match. This is intended for inputs such as `东莞建安` matching an existing sheet like `东莞市建安管桩有限公司`.

### List supplier-side contracts

```bash
./.venv/bin/python src/list_order_contracts.py -s "浙江凯航"
```

### List customer-side contracts

```bash
./.venv/bin/python src/list_sales_contracts.py -c "东莞建安"
```

### List customer-side receivable contracts

```bash
./.venv/bin/python src/list_receivable_contracts.py -c "东莞建安"
```

### Find pending supplier-side pickup rows

```bash
./.venv/bin/python src/find_pending_order_rows.py -s "浙江凯航" -n 2026032401
```

### Find pending customer-side delivery rows

```bash
./.venv/bin/python src/find_pending_sales_rows.py -c "东莞建安" -n 2026032401
```

## 订货 Inputs

- Supplier sheet, for example `浙江凯航`
- Unit price, for example `3320`
- Truck count, for example `2`
- Brand, for example `迁安`
- Spec, for example `6.5厘`
- Transport mode, for example `自提`

Example:

```bash
./.venv/bin/python src/order_entry.py -s "浙江凯航" -p 3320 -n 2 -b "迁安" -e "6.5厘" -t "自提"
```

With explicit order date:

```bash
./.venv/bin/python src/order_entry.py -s "浙江凯航" -p 3320 -n 2 -b "迁安" -e "6.5厘" -t "自提" --order-date "2026.03.20"
```

With payment amount only:

```bash
./.venv/bin/python src/order_entry.py -s "浙江凯航" -p 3320 -n 2 -b "迁安" -e "6.5厘" -t "自提" -m 181151.6
```

With explicit payment date and amount:

```bash
./.venv/bin/python src/order_entry.py -s "浙江凯航" -p 3320 -n 2 -b "迁安" -e "6.5厘" -t "自提" -m 181151.6 -d "2026.03.23"
```

### Supplier Workbook Columns

- `A`: 订货合同
- `B`: 合同编号
- `C`: 订货日期
- `D`: 品牌
- `F`: 规格
- `G`: 运输方式
- `K`: 单价
- `L`: 提货金额公式 `=J*K`
- `M`: 付款日期
- `N`: 付款金额
- `O`: 余款滚动公式
- `P`: 客户
- `Q`: 出售单价
- `R`: 自提/送到

### Payment Rule

- If `付款金额` is provided, write payment info into the first contract row.
- If `付款日期` is omitted, it defaults to `订货日期`.

## 销货 Inputs

Minimum fields to run:

- Supplier sheet
- Contract number
- Customer name
- Selling price
- Truck count

If the contract number is unknown:

```bash
./.venv/bin/python src/list_order_contracts.py -s "浙江凯航"
```

Example:

```bash
./.venv/bin/python src/sales_entry.py -s "浙江凯航" -n 2026032401 -c "东莞建安" -p 3400 -t 2
```

With optional benchmark and price diff:

```bash
./.venv/bin/python src/sales_entry.py -s "浙江凯航" -n 2026032401 -c "东莞建安" -p 3400 -t 2 --benchmark "迁安自提" --price-diff "80"
```

With explicit sales date:

```bash
./.venv/bin/python src/sales_entry.py -s "浙江凯航" -n 2026032401 -c "东莞建安" -p 3400 -t 2 --sales-date "2026.03.20"
```

If the sales contract already exists and you only need to backfill or modify customer-workbook fields:

```bash
./.venv/bin/python src/update_sales_contract.py -c "东莞建安" -n 2026032401 --set benchmark=迁安自提 --set price_diff=70
```

The customer workbook `合同编号` is generated from `销售日期`, for example `2026.03.20 -> 2026032001`.

If `-d/--delivery` is omitted, the supplier workbook `R` 列 `自提/送到` defaults to `送到`. The customer workbook `N` 列 `运输方式` is copied from the supplier workbook `G` 列 `运输方式`.

### Customer Workbook Columns

- `A`: 销售合同
- `B`: 合同编号
- `C`: 销售日期
- `D`: 品牌
- `E`: 规格
- `F`: 对标
- `G`: 成交价差
- `L`: 供应商
- `M`: 提货价
- `N`: 运输方式
- `S`: 单价
- `T`: 销售金额公式 `=R*S`
- `U`: 收款日期
- `V`: 已收款金额
- `W`: 未收款金额滚动公式

## Practical Guardrails

- Do not invent a contract number.
- Do not invent a sales price.
- Do not infer `自提` versus `送到` unless the user explicitly says so.
- Do not infer `对标` or `成交价差`; only write user-provided values.
- Customer input may be a shorthand; prefer resolving to an existing customer sheet instead of creating a new one when there is a unique match.
- For sales, if the user only gives customer and truck count, ask for or derive supplier first by listing candidate contracts per supplier.
- After a sales contract already exists, use `update_sales_contract.py` instead of rerunning `sales_entry.py` to backfill fields such as `对标` and `成交价差`.

## 送货 Inputs

Minimum fields to run:

- Supplier sheet
- Supplier-side order contract number
- Customer sheet
- Customer-side sales contract number
- Pickup date
- Truck number
- Factory weight
- Received weight

Optional fields:

- Dock / 提货码头
- Delivery date

Required freight fields:

- Fleet / 车队
- Freight input
- Freight tax flag

Example:

```bash
./.venv/bin/python src/delivery_entry.py -s "浙江凯航" -o 2026032401 -c "东莞建安" -n 2026032401 --pickup-date "2026.03.24" --truck-no "8888" --factory-weight 31.25 --received-weight 31.20 --fleet "英杰运输" --freight "35元/吨" --freight-tax "含税"
```

With explicit dock and delivery date:

```bash
./.venv/bin/python src/delivery_entry.py -s "浙江凯航" -o 2026032401 -c "东莞建安" -n 2026032401 --pickup-date "2026.03.24" --delivery-date "2026.03.25" --truck-no "8888" --factory-weight 31.25 --received-weight 31.20 --fleet "英杰运输" --freight "1200元" --freight-tax "不含税" --dock "旧港仓"
```

### Fields Written By 送货

Supplier workbook:

- `E`: 提货码头
- `H`: 提货日期
- `I`: 车号
- `J`: 出厂吨数

Customer workbook:

- `H`: 送货日期
- `I`: 车队
- `J`: 运费
- `K`: 运费是否含税
- `P`: 车号
- `Q`: 出厂吨数
- `R`: 实收吨数

### Freight Rule

- If the input is `X元/吨`, actual freight = `X * 实收吨数`
- If the input is `X元`, actual freight = `X`
- `运费是否含税` must be `含税` or `不含税`

### Matching Rule

- Supplier side: first pending row under `supplier + order_contract`
- Customer side: first pending row under `customer + sales_contract`
- Customer sheet lookup supports unique fuzzy matching for shorthand customer names.
- Use the pending-row query tools to inspect the exact row sequence before writing.

## 出具对账单 Inputs

Available filters:

- `客户`，必填
- `销售日期区间`
- `合同编号`
- `是否已经收款`

Paid status rule:

- `收款日期` and `已收款金额` both empty => `未收款`
- otherwise => `已收款`

Command:

```bash
./.venv/bin/python src/statement_issue.py --customer "东莞市建安管桩有限公司" --date-from "2026.03.01" --date-to "2026.03.31" --paid all --statement-date "2026.03.31"
```

By contract:

```bash
./.venv/bin/python src/statement_issue.py --customer "东莞市建安管桩有限公司" --contract "2026032401" --statement-date "2026.03.31"
```

If multiple contracts match, rerun with one of:

```bash
./.venv/bin/python src/statement_issue.py --customer "东莞市建安管桩有限公司" --date-from "2026.03.01" --date-to "2026.03.31" --multi-contract-mode summary
./.venv/bin/python src/statement_issue.py --customer "东莞市建安管桩有限公司" --date-from "2026.03.01" --date-to "2026.03.31" --multi-contract-mode split
```

### Statement Rules

- Seller is fixed as `东莞市华众供应链管理有限公司`
- Buyer is the resolved full customer name
- `对账日期` defaults to today
- `合同签订日期` is taken from `销售日期`
- `备注` is `品牌 + 规格`
- The statement is exported as `Excel` only
- Output is stored under `statements/<客户全称>_<销售日期或日期范围>/`
- Exported files are named `<客户全称>_<销售日期或日期范围>_对账单.xlsx`
- Do not expect PDF, HTML, or JSON files from this workflow
- The generated statement is a base Excel draft and may be further adjusted by later user instructions

## 收款 Inputs

Minimum fields to run:

- Customer sheet
- Customer-side sales contract number
- Received amount

Optional fields:

- Receipt date

Read-first example:

```bash
./.venv/bin/python src/list_receivable_contracts.py -c "东莞建安" --settled no
```

Filtered examples:

```bash
./.venv/bin/python src/list_receivable_contracts.py -c "东莞建安" -n 2026032401
./.venv/bin/python src/list_receivable_contracts.py -c "东莞建安" --date-from "2026.03.01" --date-to "2026.03.31" --settled all
```

Write example:

```bash
./.venv/bin/python src/receipt_entry.py -c "东莞建安" -n 2026032401 -a 50000
```

With explicit receipt date:

```bash
./.venv/bin/python src/receipt_entry.py -c "东莞建安" -n 2026032401 -a 50000 --receipt-date "2026.03.24"
```

### Receipt Rules

- `收款日期` defaults to today
- `已收款金额` must always be provided by the user
- `未收款金额` is not edited by the script
- The first receipt is written into the first row of the contract
- The second, third, and later receipts are written into the later contract rows whose `U/V` are still blank
- Whether the contract is settled, and whether a new receipt amount is too large, must be judged with the contract's last row as the balance row, but the actual balance should be computed from the contract rows' sales amounts and received amounts instead of relying on stale Excel formula cache values
- If receipt count exceeds the number of contract rows, raise an exception and handle manually
- If the input amount is greater than the current outstanding balance, raise an exception and do not write anything
- In this workbook, `未收款金额` is negative while the contract is still unpaid, and `0` means fully settled
