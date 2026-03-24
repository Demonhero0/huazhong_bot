# openclaw_huazhong

用于维护线材业务台账的本地脚本仓库，覆盖以下流程：

- 订货
- 销货
- 送货
- 收款
- 出具对账单

仓库通过 Python 脚本直接读写两本 Excel 台账：

- `线材供应商提货明细龙虾版.xlsx`
- `线材客户送货明细龙虾版.xlsx`

这两个 Excel 文件当前被 `.gitignore` 忽略，不会上传到 Git。

## 目录说明

- [`order_entry.py`](/Users/jianzhongsu/Desktop/openclaw_huazhong/order_entry.py)：订货
- [`sales_entry.py`](/Users/jianzhongsu/Desktop/openclaw_huazhong/sales_entry.py)：销货
- [`delivery_entry.py`](/Users/jianzhongsu/Desktop/openclaw_huazhong/delivery_entry.py)：送货
- [`receipt_entry.py`](/Users/jianzhongsu/Desktop/openclaw_huazhong/receipt_entry.py)：收款
- [`statement_issue.py`](/Users/jianzhongsu/Desktop/openclaw_huazhong/statement_issue.py)：出具对账单
- [`list_order_contracts.py`](/Users/jianzhongsu/Desktop/openclaw_huazhong/list_order_contracts.py)：查询供应商合同
- [`list_sales_contracts.py`](/Users/jianzhongsu/Desktop/openclaw_huazhong/list_sales_contracts.py)：查询客户销货合同
- [`find_pending_order_rows.py`](/Users/jianzhongsu/Desktop/openclaw_huazhong/find_pending_order_rows.py)：查询待送货的供应商行
- [`find_pending_sales_rows.py`](/Users/jianzhongsu/Desktop/openclaw_huazhong/find_pending_sales_rows.py)：查询待送货的客户行
- [`list_receivable_contracts.py`](/Users/jianzhongsu/Desktop/openclaw_huazhong/list_receivable_contracts.py)：查询待收款合同
- [`skills/steel-wire-order-sales/SKILL.md`](/Users/jianzhongsu/Desktop/openclaw_huazhong/skills/steel-wire-order-sales/SKILL.md)：给 OpenClaw 使用的 skill

## 环境准备

推荐使用仓库里的 `.venv`。

安装依赖：

```bash
source .venv/bin/activate
pip install -r requirements.txt
```

或者直接：

```bash
./.venv/bin/python -m pip install -r requirements.txt
```

当前依赖只有：

- `openpyxl==3.1.5`

## 使用前提

运行脚本前，请确认当前目录下有这两本 Excel：

- `线材供应商提货明细龙虾版.xlsx`
- `线材客户送货明细龙虾版.xlsx`

所有命令都建议在仓库根目录执行，并优先使用：

```bash
./.venv/bin/python
```

## 业务流程

### 1. 订货

作用：
- 往供应商台账新增一笔订货合同

基础命令：

```bash
./.venv/bin/python order_entry.py -s "浙江凯航" -p 3320 -n 2 -b "迁安" -e "6.5厘" -t "自提"
```

常用可选参数：

- `--order-date "YYYY.MM.DD"`：指定订货日期，默认今天
- `-m <付款金额>`：填写付款金额
- `-d "YYYY.MM.DD"`：填写付款日期

示例：

```bash
./.venv/bin/python order_entry.py -s "浙江凯航" -p 3320 -n 2 -b "迁安" -e "6.5厘" -t "自提" --order-date "2026.03.24" -m 181151.6
```

规则：

- 付款日期默认跟订货日期一致
- 如果只传付款日期、不传付款金额，脚本不会写付款信息

### 2. 销货

作用：
- 把供应商合同下的车次分配给客户
- 同时写入客户台账

如果不知道可销合同，先查：

```bash
./.venv/bin/python list_order_contracts.py -s "浙江凯航"
```

基础命令：

```bash
./.venv/bin/python sales_entry.py -s "浙江凯航" -n 2026032401 -c "东莞建安" -p 3400 -d "送到" -t 2
```

常用可选参数：

- `--sales-date "YYYY.MM.DD"`：销售日期，默认今天
- `--benchmark "<对标>"`
- `--price-diff "<成交价差>"`

示例：

```bash
./.venv/bin/python sales_entry.py -s "浙江凯航" -n 2026032401 -c "东莞建安" -p 3400 -d "送到" -t 2 --sales-date "2026.03.24" --benchmark "迁安自提" --price-diff "80"
```

规则：

- 客户简称支持模糊解析，例如 `东莞建安 -> 东莞市建安管桩有限公司`
- 销货合同编号会跟销售日期同步生成
- `对标` 和 `成交价差` 只有你显式提供时才会写入

### 3. 送货

作用：
- 同时补全供应商台账和客户台账的执行信息

如果不知道待送货行，先查：

```bash
./.venv/bin/python find_pending_order_rows.py -s "浙江凯航" -n 2026032401
./.venv/bin/python find_pending_sales_rows.py -c "东莞建安" -n 2026032401
```

基础命令：

```bash
./.venv/bin/python delivery_entry.py -s "浙江凯航" -o 2026032401 -c "东莞建安" -n 2026032401 --pickup-date "2026.03.24" --truck-no "682" --factory-weight 31.24 --received-weight 31.16 --fleet "货主帮" --freight "1000元" --freight-tax "含税"
```

常用可选参数：

- `--delivery-date "YYYY.MM.DD"`：送货日期，默认跟提货日期一致
- `--dock "<提货码头>"`

运费输入支持两种形式：

- `35元/吨`
- `1000元`

规则：

- `车号` 必填，不能猜
- `运费是否含税` 只能是 `含税` 或 `不含税`
- 如果运费写成 `X元/吨`，脚本会按 `实收吨数` 自动换算

### 4. 收款

作用：
- 在客户台账中登记收款日期和已收款金额

先查待收款合同：

```bash
./.venv/bin/python list_receivable_contracts.py -c "东莞建安"
```

按合同筛选：

```bash
./.venv/bin/python list_receivable_contracts.py -c "东莞建安" -n 2026032401
```

基础命令：

```bash
./.venv/bin/python receipt_entry.py -c "东莞建安" -n 2026032401 -a 200000
```

指定收款日期：

```bash
./.venv/bin/python receipt_entry.py -c "东莞建安" -n 2026032401 -a 200000 --receipt-date "2026.03.24"
```

规则：

- `收款日期` 默认今天
- `已收款金额` 必须由你输入
- 脚本只写 `U:收款日期`、`V:已收款金额`
- 不修改 `W:未收款金额` 的公式
- 第一次收款写合同第一条空白收款行
- 第二次、第三次收款依次往后写
- 如果付款次数大于车数，会报错让你手动处理
- 如果本次收款金额超过当前未收款金额，会直接报错

注意：

- 收款余额判断不依赖 Excel 公式缓存
- 脚本会按合同各行的销售金额和已收款金额自行计算真实余额

### 5. 出具对账单

作用：
- 按客户筛选客户台账数据
- 导出一份 Excel 对账单

基础命令：

```bash
./.venv/bin/python statement_issue.py --customer "东莞建安"
```

按合同导出：

```bash
./.venv/bin/python statement_issue.py --customer "东莞建安" --contract "2026032401"
```

按时间区间导出：

```bash
./.venv/bin/python statement_issue.py --customer "东莞建安" --date-from "2026.03.01" --date-to "2026.03.31" --paid no
```

如果筛中多个合同，可选：

```bash
./.venv/bin/python statement_issue.py --customer "东莞建安" --date-from "2026.03.01" --date-to "2026.03.31" --multi-contract-mode summary
./.venv/bin/python statement_issue.py --customer "东莞建安" --date-from "2026.03.01" --date-to "2026.03.31" --multi-contract-mode split
```

规则：

- 只导出 Excel，不导出 PDF
- 输出目录在 `statements/`
- 文件名格式为：`<客户全称>_<销售日期或日期范围>_对账单.xlsx`
- 客户支持简称匹配

## 常用查询命令

查询供应商可销合同：

```bash
./.venv/bin/python list_order_contracts.py -s "浙江凯航"
```

查询客户销货合同：

```bash
./.venv/bin/python list_sales_contracts.py -c "东莞建安"
```

查询客户待收款合同：

```bash
./.venv/bin/python list_receivable_contracts.py -c "东莞建安"
```

查询待送货的供应商行：

```bash
./.venv/bin/python find_pending_order_rows.py -s "浙江凯航" -n 2026032401
```

查询待送货的客户行：

```bash
./.venv/bin/python find_pending_sales_rows.py -c "东莞建安" -n 2026032401
```

## OpenClaw Skill

如果你要让 OpenClaw 直接调用这套流程，核心 skill 在：

- [`skills/steel-wire-order-sales/SKILL.md`](/Users/jianzhongsu/Desktop/openclaw_huazhong/skills/steel-wire-order-sales/SKILL.md)

### 部署步骤

#### 1. 保持仓库本地可用

OpenClaw 最终调用的是这个仓库里的本地脚本，所以仓库目录本身要保留在本机，例如：

```bash
/Users/你的用户名/Desktop/openclaw_huazhong
```

并且目录下要有：

- Python 脚本
- `.venv`
- 两本实际业务 Excel

#### 2. 安装 Python 依赖

```bash
./.venv/bin/python -m pip install -r requirements.txt
```

#### 3. 把 skill 安装到 OpenClaw 可读取的位置

如需安装到本机 skills 目录，可手动复制：

```bash
mkdir -p ~/.codex/skills
cp -R skills/steel-wire-order-sales ~/.codex/skills/
```

安装完成后，通常会得到：

```bash
~/.codex/skills/steel-wire-order-sales/
```

里面至少应包含：

- `SKILL.md`
- `agents/openai.yaml`
- `references/workflow.md`

#### 4. 在 OpenClaw 中使用这个 skill

进入这个仓库目录后，让 OpenClaw 在当前仓库里工作，并显式调用这个 skill。

常见用法示例：

```text
Use $steel-wire-order-sales to handle 订货、销货、送货、收款和出具对账单 in this repository.
```

或者直接下业务指令：

```text
Use $steel-wire-order-sales. 帮我把浙江凯航3320元2车迁安6.5厘自提录成订货。
```

```text
Use $steel-wire-order-sales. 帮我查询东莞建安当前未结清的合同。
```

```text
Use $steel-wire-order-sales. 帮我出具中山豪强未付款的对账单。
```

#### 5. OpenClaw 的运行前提

为了让 skill 正常工作，OpenClaw 运行时需要满足：

- 当前工作目录就是这个仓库根目录
- 仓库里存在 `.venv`
- 两本 Excel 文件就在仓库根目录
- OpenClaw 对这个仓库有读写权限

### 部署后的建议验证

部署完后，建议先做一次只读验证：

```bash
./.venv/bin/python list_order_contracts.py -s "浙江凯航"
./.venv/bin/python list_sales_contracts.py -c "东莞建安"
./.venv/bin/python list_receivable_contracts.py -c "东莞建安"
```

如果这 3 条都能正常返回，说明：

- skill 对应的脚本在
- Excel 文件在
- Python 环境在
- OpenClaw 后续就可以直接调用这些流程

## 注意事项

- 脚本默认假设 Excel 表头结构稳定，不要随意改列顺序
- 所有日期格式统一使用 `YYYY.MM.DD`
- 客户简称匹配必须唯一；如果不唯一，脚本会报错
- 当前 `.gitignore` 会忽略所有 `.xlsx` 文件，请自行保管业务数据
- 在正式批量使用前，建议先复制 Excel 做一次完整演练

## Git 初始化

如果你要把仓库推到 GitHub，基础命令：

```bash
git init
git add .
git commit -m "Initial commit"
```
